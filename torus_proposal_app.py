# torus_proposal_app.py
# Torus Group – Cleaning Service Agreement Builder with AI RFP/PWS Analyzer (Streamlit Cloud-safe)

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
    net_terms: int
    sales_tax_percent: float

    square_footage: int
    floor_types: str

    num_offices: int
    num_conference_rooms: int
    num_break_rooms: int
    num_bathrooms: int

    hand_soap: str
    paper_towels: str
    toilet_paper: str

    pricing_mode: str
    monthly_price: float

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
# AI RFP ANALYSIS (STABLE ON STREAMLIT CLOUD)
# =========================
def analyze_rfp_with_ai(text: str) -> Dict[str, Any]:
    client = get_openai_client()

    instructions = """
You are assisting a janitorial contractor responding to an RFP or PWS.

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
- Include realistic janitorial tasks
- Keep schedule_rows to about 12–30 rows
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

    content = resp.choices[0].message.content
    return json.loads(content)


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
    """
    Adds a bullet paragraph. Uses a bullet style if it exists in the template,
    otherwise falls back to a manual bullet so it never crashes.
    """
    bullet_style_candidates = ("List Bullet", "List Paragraph", "Bullet List")

    for style_name in bullet_style_candidates:
        try:
            doc.add_paragraph(text, style=style_name)
            return
        except KeyError:
            continue

    # Template has no bullet style → safe fallback
    doc.add_paragraph(f"• {text}")


def add_cover_page(doc: Document, client: str, body: str):
    # Letterhead/header should already be in proposal_template.docx
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
    if title.runs:
        title.runs[0].bold = True
    else:
        title.add_run("SCOPE OF WORK – CLEANING SCHEDULE").bold = True

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

    # Ensure headers show on first page even if template has "Different First Page"
    for s in doc.sections:
        s.different_first_page_header_footer = False

    if p.include_cover_page:
        add_cover_page(doc, p.client, p.cover_letter_body)

    title = doc.add_paragraph("CLEANING SERVICE AGREEMENT")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title.runs:
        title.runs[0].bold = True
    else:
        title.add_run("CLEANING SERVICE AGREEMENT").bold = True

    doc.add_paragraph(f"Client: {p.client}")
    doc.add_paragraph(f"Facility: {p.facility_name}")

    doc.add_paragraph("Service Address(es):")
    for a in (p.service_addresses or []):
        a2 = (a or "").strip()
        if a2:
            # ✅ FIX: bullet helper prevents KeyError if template lacks 'List Bullet'
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
        if h.runs:
            h.runs[0].bold = True
        else:
            h.add_run("CLEANING PLAN").bold = True
        doc.add_paragraph(p.cleaning_plan.strip())
        doc.add_paragraph("")

    h = doc.add_paragraph("GENERAL REQUIREMENTS")
    if h.runs:
        h.runs[0].bold = True
    else:
        h.add_run("GENERAL REQUIREMENTS").bold = True

    doc.add_paragraph(
        "Contractor shall provide all labor, supervision, and personnel necessary to perform the services "
        "described in this agreement. Unless otherwise stated, Contractor shall provide all standard equipment "
        "and cleaning supplies.\n\n"
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
# STREAMLIT UI
# =========================
st.set_page_config(layout="wide")
st.title("Torus Group – Cleaning Proposal Builder")

# Persist “apply AI” results across reruns
st.session_state.setdefault("cleaning_plan_prefill", "")
st.session_state.setdefault("schedule_rows_prefill", None)

with st.sidebar:
    st.header("Proposal Inputs")
    st.caption(f"OpenAI key loaded: {bool(st.secrets.get('OPENAI_API_KEY'))}")
    st.caption(f"Template found: {os.path.exists('proposal_template.docx')}")

    client = st.text_input("Client")
    facility = st.text_input("Facility name")

    service_begin_date = st.text_input("Service begin date")
    service_end_date = st.text_input("Service end date")

    days = st.number_input("Days per week", min_value=1, value=5)
    times = st.text_input("Cleaning times (e.g., 6 PM – 10 PM)")

    st.subheader("Service Addresses")
    addresses = st.text_area("One address per line").splitlines()

    st.subheader("Room Counts")
    offices = st.number_input("Offices", min_value=0)
    conference = st.number_input("Conference rooms", min_value=0)
    breaks = st.number_input("Break rooms", min_value=0)
    baths = st.number_input("Bathrooms", min_value=0)

    st.subheader("Consumables")
    soap = st.selectbox("Hand soap", ["Contractor", "Client"])
    towels = st.selectbox("Paper towels", ["Contractor", "Client"])
    tp = st.selectbox("Toilet paper", ["Contractor", "Client"])

    st.subheader("Cover Page")
    include_cover = st.checkbox("Include cover page", True)
    cover_body = st.text_area("Cover letter body")

st.subheader("Cleaning Plan")
cleaning_plan = st.text_area("Cleaning Plan (optional)", value=st.session_state["cleaning_plan_prefill"])

st.subheader("Notes")
notes = st.text_area("Notes")

# Schedule editor (dynamic; you can add/edit/remove tasks)
st.subheader("Cleaning Schedule")
st.caption("Add, remove, or edit tasks below. Use the last row to add new cleaning tasks.")

default_rows = [
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

prefill_rows = st.session_state["schedule_rows_prefill"] or default_rows
df = pd.DataFrame(prefill_rows, columns=["Task", "Daily", "Weekly", "Monthly"])
edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

# Convert edited schedule to tuples; ignore blank tasks
schedule_rows = [
    (r.Task, r.Daily, r.Weekly, r.Monthly)
    for r in edited.itertuples()
    if str(r.Task).strip()
]

# =========================
# AI ANALYZER
# =========================
st.divider()
st.subheader("RFP / PWS Analyzer (AI)")

uploads = st.file_uploader("Upload RFP/PWS", type=["pdf", "docx", "txt"], accept_multiple_files=True)

colA, colB = st.columns([1, 2])
with colA:
    run_ai = st.button("Analyze with AI")
with colB:
    if st.button("Clear AI results"):
        for k in ["ai", "cleaning_plan_prefill", "schedule_rows_prefill"]:
            if k in st.session_state:
                del st.session_state[k]
        st.session_state["cleaning_plan_prefill"] = ""
        st.session_state["schedule_rows_prefill"] = None
        st.success("Cleared.")
        st.rerun()

if run_ai and uploads:
    try:
        full_text = "\n\n".join(extract_text(f) for f in uploads)
        if not full_text.strip():
            st.error("Could not extract any text from the upload(s). If PDF is scanned, OCR is needed.")
        else:
            with st.spinner("Analyzing…"):
                result = analyze_rfp_with_ai(full_text)
            st.session_state["ai"] = result
            st.success("Analysis complete.")
    except Exception as e:
        st.exception(e)

if "ai" in st.session_state:
    ai = st.session_state["ai"]
    st.text_area("AI Cleaning Plan", ai.get("cleaning_plan_draft", ""), height=160)
    st.text_area("AI Scope of Work", ai.get("scope_of_work_draft", ""), height=160)

    qs = ai.get("clarifying_questions", [])
    if qs:
        st.write("**Clarifying questions**")
        for q in qs:
            st.write(f"- {q}")

    if st.button("Apply AI to proposal"):
        st.session_state["cleaning_plan_prefill"] = ai.get("cleaning_plan_draft", "")
        st.session_state["schedule_rows_prefill"] = [
            (r.get("task", ""), bool(r.get("daily", False)), bool(r.get("weekly", False)), bool(r.get("monthly", False)))
            for r in ai.get("schedule_rows", [])
            if (r.get("task") or "").strip()
        ] or None
        st.success("Applied. Scroll up—your Cleaning Plan and Schedule are now prefilled.")
        st.rerun()

# =========================
# BUILD + DOWNLOAD
# =========================
p = ProposalInputs(
    client=client,
    facility_name=facility,
    service_begin_date=service_begin_date,
    service_end_date=service_end_date,
    service_addresses=addresses,
    days_per_week=int(days),
    cleaning_times=times,
    net_terms=30,
    sales_tax_percent=0.0,
    square_footage=0,
    floor_types="",
    num_offices=int(offices),
    num_conference_rooms=int(conference),
    num_break_rooms=int(breaks),
    num_bathrooms=int(baths),
    hand_soap=soap,
    paper_towels=towels,
    toilet_paper=tp,
    pricing_mode="Monthly",
    monthly_price=0.0,
    include_cover_page=include_cover,
    cover_letter_body=cover_body,
    cleaning_plan=cleaning_plan,
    notes=notes,
)

docx_bytes = build_doc(p, schedule_rows)

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
