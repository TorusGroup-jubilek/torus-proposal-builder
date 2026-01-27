import os
import json
import datetime
from io import BytesIO
from dataclasses import dataclass, asdict
from typing import List, Dict, Any

import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from pypdf import PdfReader
from openai import OpenAI

CHECK = "✓"
COMPANY_NAME = "Torus Group"


# =========================
# DATA MODEL
# =========================
@dataclass
class ProposalInputs:
    client: str
    facility_name: str
    service_begin_date: str
    service_end_date: str
    service_addresses: List[str]
    days_per_week: int
    cleaning_times: str

    num_offices: int
    num_conference_rooms: int
    num_break_rooms: int
    num_bathrooms: int

    hand_soap: str
    paper_towels: str
    toilet_paper: str

    include_cover_page: bool
    cover_letter_body: str
    cleaning_plan: str
    notes: str


# =========================
# OPENAI CLIENT
# =========================
def get_openai_client() -> OpenAI:
    key = st.secrets.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("Missing OPENAI_API_KEY in Streamlit secrets.")
    return OpenAI(api_key=key)


# =========================
# AI RFP ANALYSIS (stable)
# =========================
def analyze_rfp_with_ai(text: str) -> Dict[str, Any]:
    client = get_openai_client()

    instructions = """
Return ONLY valid JSON with this exact structure:

{
  "cleaning_plan_draft": "string",
  "scope_of_work_draft": "string",
  "schedule_rows": [
    {"task": "string", "daily": true, "weekly": false, "monthly": false}
  ],
  "clarifying_questions": ["string"]
}

Rules:
- JSON only (no markdown, no explanations)
- Keep schedule_rows ~12–30 items
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": instructions},
            {"role": "user", "content": text[:120000]},
        ],
        response_format={"type": "json_object"},
        temperature=0.2,
    )
    return json.loads(resp.choices[0].message.content)


# =========================
# FILE EXTRACTION
# =========================
def extract_text(uploaded_file) -> str:
    name = (uploaded_file.name or "").lower()
    data = uploaded_file.read()

    if name.endswith(".pdf"):
        reader = PdfReader(BytesIO(data))
        return "\n".join((p.extract_text() or "") for p in reader.pages).strip()

    if name.endswith(".docx"):
        doc = Document(BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs).strip()

    try:
        return data.decode("utf-8", errors="ignore").strip()
    except Exception:
        return ""


# =========================
# WORD HELPERS
# =========================
def add_bullet_paragraph(doc: Document, text: str):
    """Use a bullet style if present; otherwise safe manual bullet so template differences never crash."""
    for style_name in ("List Bullet", "List Paragraph", "Bullet List"):
        try:
            doc.add_paragraph(text, style=style_name)
            return
        except KeyError:
            continue
    doc.add_paragraph(f"• {text}")


def add_cover_page(doc: Document, client: str, body: str):
    doc.add_paragraph(client)
    doc.add_paragraph("")
    doc.add_paragraph("Attn: ______________________")
    doc.add_paragraph("")
    doc.add_paragraph("Re: Janitorial Services Proposal")
    doc.add_paragraph("")
    doc.add_paragraph(f"Dear {client},")
    doc.add_paragraph("")
    doc.add_paragraph(body or "")
    doc.add_paragraph("")
    doc.add_paragraph("Respectfully,")
    doc.add_paragraph("")
    doc.add_paragraph("Kara Jubilee")
    doc.add_paragraph("Owner")
    doc.add_paragraph("Torus Cleaning Services")
    doc.add_page_break()


def add_scope_table(doc: Document, rows: List[tuple]):
    title = doc.add_paragraph("SCOPE OF WORK – CLEANING SCHEDULE")
    (title.runs[0] if title.runs else title.add_run("SCOPE OF WORK – CLEANING SCHEDULE")).bold = True

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr = table.rows[0].cells
    hdr[0].text = "Task"
    hdr[1].text = "Daily"
    hdr[2].text = "Weekly"
    hdr[3].text = "Monthly"

    for task, daily, weekly, monthly in rows:
        row = table.add_row().cells
        row[0].text = str(task)
        row[1].text = CHECK if bool(daily) else ""
        row[2].text = CHECK if bool(weekly) else ""
        row[3].text = CHECK if bool(monthly) else ""

    doc.add_paragraph("")


def build_doc(p: ProposalInputs, schedule_rows: List[tuple]) -> bytes:
    template_path = "proposal_template.docx"
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    for s in doc.sections:
        s.different_first_page_header_footer = False

    if p.include_cover_page:
        add_cover_page(doc, p.client, p.cover_letter_body)

    title = doc.add_paragraph("CLEANING SERVICE AGREEMENT")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    (title.runs[0] if title.runs else title.add_run("CLEANING SERVICE AGREEMENT")).bold = True

    doc.add_paragraph(f"Client: {p.client}")
    doc.add_paragraph(f"Facility: {p.facility_name}")

    doc.add_paragraph("Service Address(es):")
    for a in (p.service_addresses or []):
        a2 = (a or "").strip()
        if a2:
            add_bullet_paragraph(doc, a2)

    doc.add_paragraph("")
    doc.add_paragraph(
        f"{p.client}, ('Client'), enters into this agreement on this date ______________ "
        f"for Torus Cleaning Services ('Contractor'), to provide janitorial services for facility/facilities "
        f"located at the addresses listed above."
    )
    doc.add_paragraph(
        f"Contractor shall provide janitorial services {p.days_per_week} per week between the hours of "
        f"{p.cleaning_times}."
    )
    doc.add_paragraph(
        f"The contract period is as follows {p.service_begin_date} to {p.service_end_date}."
    )

    add_scope_table(doc, schedule_rows)

    if (p.cleaning_plan or "").strip():
        h = doc.add_paragraph("CLEANING PLAN")
        (h.runs[0] if h.runs else h.add_run("CLEANING PLAN")).bold = True
        doc.add_paragraph(p.cleaning_plan.strip())
        doc.add_paragraph("")

    h = doc.add_paragraph("GENERAL REQUIREMENTS")
    (h.runs[0] if h.runs else h.add_run("GENERAL REQUIREMENTS")).bold = True
    doc.add_paragraph(
        "Contractor shall provide all labor, supervision, and personnel necessary to perform the services described.\n\n"
        "Consumable supplies:\n"
        f"• Hand soap: {p.hand_soap}\n"
        f"• Paper towels: {p.paper_towels}\n"
        f"• Toilet paper: {p.toilet_paper}\n"
    )

    doc.add_paragraph("NOTES")
    doc.add_paragraph((p.notes or "").strip() or "(none)")

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# =========================
# APP UI
# =========================
st.set_page_config(layout="wide")
st.title("Torus Group – Cleaning Proposal Builder")

# Hide Streamlit UI elements that often create bottom “white bar” / layout shifts
st.markdown(
    """
<style>
/* Hide Streamlit footer/menu/toolbar/status */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
div[data-testid="stToolbar"] { visibility: hidden; height: 0px; }
div[data-testid="stStatusWidget"] { visibility: hidden; height: 0px; }

/* Reduce bottom padding that can look like a blank bar */
div.block-container { padding-bottom: 1rem; }

/* In some Streamlit builds, this bottom container shows as a white block */
div[data-testid="stBottomBlockContainer"] { display: none; }
</style>
""",
    unsafe_allow_html=True,
)

# Session defaults
st.session_state.setdefault("ai", None)

default_schedule_rows = [
    ("Empty trash & replace liners", True, False, False),
    ("Clean & disinfect restrooms", True, False, False),
    ("Vacuum carpet / sweep hard floors", True, False, False),
    ("Wipe high-touch points (handles, switches)", True, False, False),
    ("Dust reachable surfaces", False, True, False),
    ("Mop hard floors (as applicable)", False, True, False),
    ("Clean break room counters & sink", False, True, False),
    ("Glass/mirrors touch-up", False, True, False),
    ("High dusting (vents/ledges)", False, False, True),
    ("Detail baseboards/edges", False, False, True),
]

if "schedule_df" not in st.session_state:
    st.session_state["schedule_df"] = pd.DataFrame(
        default_schedule_rows, columns=["Task", "Daily", "Weekly", "Monthly"]
    )

# --- ONE FORM (inputs + schedule + AI upload) ---
with st.form("proposal_form", clear_on_submit=False):
    st.subheader("Client & Contract")
    c1, c2, c3 = st.columns(3)
    with c1:
        client = st.text_input("Client (legal name)")
        facility = st.text_input("Facility / Location name")
    with c2:
        service_begin_date = st.text_input("Service begin date")
        service_end_date = st.text_input("Service end date")
    with c3:
        days = st.number_input("Days per week", min_value=1, value=5)
        times = st.text_input("Cleaning times (e.g., 6 PM – 10 PM)")

    st.subheader("Service Addresses")
    addresses_text = st.text_area("One address per line", height=120)

    st.subheader("Room Counts")
    r1, r2, r3, r4 = st.columns(4)
    with r1:
        offices = st.number_input("Offices", min_value=0)
    with r2:
        conference = st.number_input("Conference rooms", min_value=0)
    with r3:
        breaks = st.number_input("Break rooms", min_value=0)
    with r4:
        baths = st.number_input("Bathrooms", min_value=0)

    st.subheader("Consumables")
    s1, s2, s3 = st.columns(3)
    with s1:
        soap = st.selectbox("Hand soap", ["Contractor", "Client"])
    with s2:
        towels = st.selectbox("Paper towels", ["Contractor", "Client"])
    with s3:
        tp = st.selectbox("Toilet paper", ["Contractor", "Client"])

    st.subheader("Cover Page")
    include_cover = st.checkbox("Include cover page", True)
    cover_body = st.text_area("Cover letter body", height=120)

    st.subheader("Cleaning Plan & Notes")
    cleaning_plan = st.text_area("Cleaning Plan (optional)", height=140)
    notes = st.text_area("Notes", height=100)

    st.subheader("Cleaning Schedule")
    st.caption("Edit tasks freely. Add new rows at the bottom.")
    schedule_df = st.data_editor(
        st.session_state["schedule_df"],
        num_rows="dynamic",
        use_container_width=True,
        height=320,
    )

    st.subheader("RFP / PWS Analyzer (optional)")
    uploads = st.file_uploader(
        "Upload RFP/PWS (PDF, DOCX, TXT)",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
    )

    colA, colB, colC = st.columns([1, 1, 2])
    with colA:
        run_ai = st.form_submit_button("Analyze with AI")
    with colB:
        generate_doc = st.form_submit_button("Generate Proposal")
    with colC:
        st.caption(f"Template found: {os.path.exists('proposal_template.docx')}")

# Keep schedule edits in session
st.session_state["schedule_df"] = schedule_df

# Parse addresses
addresses = [x.strip() for x in (addresses_text or "").splitlines() if x.strip()]

# Analyze AI on submit
if run_ai:
    if not uploads:
        st.error("Please upload at least one RFP/PWS file.")
    else:
        try:
            full_text = "\n\n".join(extract_text(f) for f in uploads)
            if not full_text.strip():
                st.error("Could not extract text from the upload(s). If PDF is scanned, OCR is needed.")
            else:
                with st.spinner("Analyzing…"):
                    st.session_state["ai"] = analyze_rfp_with_ai(full_text)
                st.success("AI analysis complete.")
        except Exception as e:
            st.exception(e)

# Show AI results (outside the form, so it doesn’t shift inputs)
if st.session_state.get("ai"):
    ai = st.session_state["ai"]
    st.divider()
    st.subheader("AI Results")
    st.text_area("AI Cleaning Plan", ai.get("cleaning_plan_draft", ""), height=160)
    st.text_area("AI Scope of Work", ai.get("scope_of_work_draft", ""), height=160)

    qs = ai.get("clarifying_questions", [])
    if qs:
        st.write("**Clarifying questions**")
        for q in qs:
            st.write(f"- {q}")

    # Apply AI to schedule/editor
    if st.button("Apply AI schedule to table"):
        rows = []
        for r in ai.get("schedule_rows", []):
            task = (r.get("task") or "").strip()
            if not task:
                continue
            rows.append((task, bool(r.get("daily")), bool(r.get("weekly")), bool(r.get("monthly"))))
        if rows:
            st.session_state["schedule_df"] = pd.DataFrame(rows, columns=["Task", "Daily", "Weekly", "Monthly"])
            st.success("Applied AI schedule. Scroll up to see it in the table.")
        else:
            st.warning("AI did not return usable schedule rows.")

# Generate proposal
if generate_doc:
    p = ProposalInputs(
        client=client,
        facility_name=facility,
        service_begin_date=service_begin_date,
        service_end_date=service_end_date,
        service_addresses=addresses,
        days_per_week=int(days),
        cleaning_times=times,
        num_offices=int(offices),
        num_conference_rooms=int(conference),
        num_break_rooms=int(breaks),
        num_bathrooms=int(baths),
        hand_soap=soap,
        paper_towels=towels,
        toilet_paper=tp,
        include_cover_page=include_cover,
        cover_letter_body=cover_body,
        cleaning_plan=cleaning_plan,
        notes=notes,
    )

    # Convert schedule DF to tuples and skip blanks
    schedule_rows = [
        (r.Task, r.Daily, r.Weekly, r.Monthly)
        for r in st.session_state["schedule_df"].itertuples()
        if str(r.Task).strip()
    ]

    docx_bytes = build_doc(p, schedule_rows)

    st.success("Proposal generated.")
    st.download_button(
        "Download Word Proposal",
        data=docx_bytes,
        file_name=f"Torus_Cleaning_Agreement_{datetime.date.today().isoformat()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.download_button(
        "Download Inputs (JSON)",
        data=json.dumps(asdict(p), indent=2).encode("utf-8"),
        file_name=f"Torus_Inputs_{datetime.date.today().isoformat()}.json",
        mime="application/json",
    )
