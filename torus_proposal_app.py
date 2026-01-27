# torus_proposal_app.py
# Torus Group – Cleaning Service Agreement Builder with AI RFP/PWS Analyzer (Streamlit Cloud-safe)
# NORMAL LAYOUT (no form), with:
# ✅ Dynamic “Additional Room Types” (name + count)
# ✅ Standard Room Counts (offices, conference rooms, break rooms, bathrooms)
# ✅ Standard cover letter (auto-fills Client Name) + toggle + editable
# ✅ Dynamic cleaning schedule table (add/edit rows)
# ✅ Word template support (proposal_template.docx)
# ✅ Bullet-style fallback (prevents KeyError on List Bullet)

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

    # Standard room counts
    num_offices: int
    num_conference_rooms: int
    num_break_rooms: int
    num_bathrooms: int

    # Custom room types
    custom_rooms: List[Dict[str, int]]

    # Consumables
    hand_soap: str
    paper_towels: str
    toilet_paper: str

    # Cover page
    include_cover_page: bool
    cover_letter_body: str

    # Optional sections
    cleaning_plan: str
    notes: str


# =========================
# COVER LETTER DEFAULT
# =========================
def default_cover_letter(client_name: str) -> str:
    cn = client_name.strip() or "[Client Name]"
    return f"""Hello {cn},

I want to personally take the opportunity to say thank you for considering Torus Cleaning as an option for your commercial cleaning needs. We pride ourselves as a core value based business and seek to partner with those that align with the culture we continue to build. Our capable crews of background checked staff are constantly growing to accommodate the needs of our customers. 

As the President, I have over 20 years of project and program management in both the military and corporate settings. Believe when I tell you I am no stranger to long, busy days that carry over to the next. We strive to deliver a professional, trustworthy service that affords our customers the peace of mind to know their spaces are well maintained. 

Thank you again for the opportunity to partner with you to take your spaces beyond clean enough!

Respectfully,

Kary Jubilee - President
Torus Cleaning Services
"""


# =========================
# OPENAI CLIENT
# =========================
def get_openai_client() -> OpenAI:
    key = st.secrets.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("Missing OPENAI_API_KEY in Streamlit secrets.")
    return OpenAI(api_key=key)


# =========================
# AI RFP ANALYSIS (Stable)
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
    for style_name in ("List Bullet", "List Paragraph", "Bullet List"):
        try:
            doc.add_paragraph(text, style=style_name)
            return
        except KeyError:
            continue
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
    doc.add_paragraph("Kary Jubilee")
    doc.add_paragraph("President")
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

    # Optional cover page
    if p.include_cover_page:
        add_cover_page(doc, p.client or "[Client Name]", p.cover_letter_body)

    # Agreement title
    title = doc.add_paragraph("CLEANING SERVICE AGREEMENT")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title.runs:
        title.runs[0].bold = True
    else:
        title.add_run("CLEANING SERVICE AGREEMENT").bold = True

    # Basic info
    doc.add_paragraph(f"Client: {p.client}")
    doc.add_paragraph(f"Facility: {p.facility_name}")

    # Addresses
    doc.add_paragraph("Service Address(es):")
    for a in (p.service_addresses or []):
        a2 = (a or "").strip()
        if a2:
            add_bullet_paragraph(doc, a2)

    doc.add_paragraph("")

    # Required paragraphs (date is blank for typing)
    client_name = p.client.strip() if p.client.strip() else "[Client Name]"
    doc.add_paragraph(
        f"{client_name}, ('Client'), enters into this agreement on this date ______________ "
        f"for Torus Cleaning Services ('Contractor'), to provide janitorial services for facility/facilities "
        f"located at the addresses listed above."
    )
    doc.add_paragraph(
        f"Contractor shall provide janitorial services {p.days_per_week} per week between the hours of "
        f"{p.cleaning_times} for the facility/facilities located at the addresses listed above."
    )
    doc.add_paragraph(
        f"The contract period is as follows {p.service_begin_date} to {p.service_end_date}."
    )

    doc.add_paragraph("")

    # Room counts section
    h = doc.add_paragraph("ROOM COUNTS")
    if h.runs:
        h.runs[0].bold = True
    else:
        h.add_run("ROOM COUNTS").bold = True

    add_bullet_paragraph(doc, f"Offices: {p.num_offices}")
    add_bullet_paragraph(doc, f"Conference rooms: {p.num_conference_rooms}")
    add_bullet_paragraph(doc, f"Break rooms: {p.num_break_rooms}")
    add_bullet_paragraph(doc, f"Bathrooms: {p.num_bathrooms}")

    # Custom room types (named in-app)
    custom = []
    for r in (p.custom_rooms or []):
        rt = str(r.get("type", "")).strip()
        rc = int(r.get("count", 0) or 0)
        if rt and rc > 0:
            custom.append((rt, rc))

    if custom:
        doc.add_paragraph("Additional room types:")
        for rt, rc in custom:
            add_bullet_paragraph(doc, f"{rt}: {rc}")

    doc.add_paragraph("")

    # Scope of work schedule table
    add_scope_table(doc, schedule_rows)

    # Optional Cleaning Plan
    if (p.cleaning_plan or "").strip():
        h = doc.add_paragraph("CLEANING PLAN")
        (h.runs[0] if h.runs else h.add_run("CLEANING PLAN")).bold = True
        doc.add_paragraph(p.cleaning_plan.strip())
        doc.add_paragraph("")

    # General requirements from consumables
    h = doc.add_paragraph("GENERAL REQUIREMENTS")
    (h.runs[0] if h.runs else h.add_run("GENERAL REQUIREMENTS")).bold = True
    doc.add_paragraph(
        "Contractor shall provide all labor, supervision, and personnel necessary to perform the services "
        "described in this agreement. Unless otherwise stated, Contractor shall provide all standard equipment "
        "and cleaning supplies.\n\n"
        "Consumable supplies:\n"
        f"• Hand soap: {p.hand_soap}\n"
        f"• Paper towels: {p.paper_towels}\n"
        f"• Toilet paper: {p.toilet_paper}\n"
    )

    # Notes
    doc.add_paragraph("NOTES")
    doc.add_paragraph((p.notes or "").strip() or "(none)")

    # Save
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# =========================
# STREAMLIT UI (Normal layout)
# =========================
st.set_page_config(layout="wide")
st.title("Torus Group – Cleaning Proposal Builder")

# Session state defaults
st.session_state.setdefault("ai", None)
st.session_state.setdefault("cleaning_plan_prefill", "")
st.session_state.setdefault("schedule_rows_prefill", None)

# Standard cover letter behavior
st.session_state.setdefault("cover_body_custom", "")

# Dynamic custom rooms
if "custom_rooms" not in st.session_state:
    st.session_state["custom_rooms"] = [{"type": "", "count": 0}]

# Schedule DF
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

# Sidebar inputs (normal layout)
with st.sidebar:
    st.header("Proposal Inputs")
    st.caption(f"OpenAI key loaded: {bool(st.secrets.get('OPENAI_API_KEY'))}")
    st.caption(f"Template found: {os.path.exists('proposal_template.docx')}")

    client = st.text_input("Client", value="")
    facility = st.text_input("Facility name", value="")

    service_begin_date = st.text_input("Service begin date", value="")
    service_end_date = st.text_input("Service end date", value="")

    days = st.number_input("Days per week", min_value=1, value=5)
    times = st.text_input("Cleaning times (e.g., 6 PM – 10 PM)", value="")

    st.subheader("Service Addresses")
    addresses = st.text_area("One address per line").splitlines()

    st.subheader("Room Counts (Standard)")
    offices = st.number_input("Offices", min_value=0, value=0)
    conference = st.number_input("Conference rooms", min_value=0, value=0)
    breaks = st.number_input("Break rooms", min_value=0, value=0)
    baths = st.number_input("Bathrooms", min_value=0, value=0)

    st.subheader("Additional Room Types")
    st.caption("Add any room type and count (e.g., Exam rooms, Classrooms, Server rooms).")

    for i, room in enumerate(st.session_state["custom_rooms"]):
        c1, c2, c3 = st.columns([3, 1, 1])
        with c1:
            st.session_state["custom_rooms"][i]["type"] = st.text_input(
                f"room_type_{i}",
                value=room.get("type", ""),
                placeholder="e.g., Exam Rooms",
                label_visibility="collapsed",
            )
        with c2:
            st.session_state["custom_rooms"][i]["count"] = st.number_input(
                f"room_count_{i}",
                min_value=0,
                step=1,
                value=int(room.get("count", 0) or 0),
                label_visibility="collapsed",
            )
        with c3:
            if st.button("Remove", key=f"remove_room_{i}") and len(st.session_state["custom_rooms"]) > 1:
                st.session_state["custom_rooms"].pop(i)
                st.rerun()

    if st.button("Add another room type"):
        st.session_state["custom_rooms"].append({"type": "", "count": 0})
        st.rerun()

    st.subheader("Consumables")
    soap = st.selectbox("Hand soap", ["Contractor", "Client"])
    towels = st.selectbox("Paper towels", ["Contractor", "Client"])
    tp = st.selectbox("Toilet paper", ["Contractor", "Client"])

    st.subheader("Cover Page")
    include_cover = st.checkbox("Include cover page", value=True)

    use_standard_cover = st.checkbox("Use Torus standard cover letter", value=True)
    if use_standard_cover:
        cover_body = default_cover_letter(client)
    else:
        cover_body = st.text_area(
            "Cover letter body",
            value=st.session_state.get("cover_body_custom", ""),
            height=260
        )
        st.session_state["cover_body_custom"] = cover_body

# Main page optional sections
st.subheader("Cleaning Plan")
cleaning_plan = st.text_area("Cleaning Plan (optional)", value=st.session_state["cleaning_plan_prefill"], height=160)

st.subheader("Notes")
notes = st.text_area("Notes", height=120)

# Cleaning schedule editor (dynamic)
st.subheader("Cleaning Schedule")
st.caption("Add, remove, or edit tasks below. Use the last row to add new cleaning tasks.")

st.session_state["schedule_df"] = st.data_editor(
    st.session_state["schedule_df"],
    num_rows="dynamic",
    use_container_width=True,
    height=320,
)

# Convert schedule DF to tuples; ignore blank tasks
schedule_rows = [
    (r.Task, r.Daily, r.Weekly, r.Monthly)
    for r in st.session_state["schedule_df"].itertuples()
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
        st.session_state["ai"] = None
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
                st.session_state["ai"] = analyze_rfp_with_ai(full_text)
            st.success("Analysis complete.")
    except Exception as e:
        st.exception(e)

if st.session_state.get("ai"):
    ai = st.session_state["ai"]
    st.subheader("AI Results")
    st.text_area("AI Cleaning Plan", ai.get("cleaning_plan_draft", ""), height=160)
    st.text_area("AI Scope of Work", ai.get("scope_of_work_draft", ""), height=160)

    qs = ai.get("clarifying_questions", [])
    if qs:
        st.write("**Clarifying questions**")
        for q in qs:
            st.write(f"- {q}")

    if st.button("Apply AI to proposal"):
        st.session_state["cleaning_plan_prefill"] = ai.get("cleaning_plan_draft", "")

        # Apply schedule rows into the editor table
        rows = []
        for r in ai.get("schedule_rows", []):
            task = (r.get("task") or "").strip()
            if not task:
                continue
            rows.append((task, bool(r.get("daily", False)), bool(r.get("weekly", False)), bool(r.get("monthly", False))))
        if rows:
            st.session_state["schedule_df"] = pd.DataFrame(rows, columns=["Task", "Daily", "Weekly", "Monthly"])
        st.success("Applied. Scroll up—Cleaning Plan and Schedule updated.")
        st.rerun()

# =========================
# BUILD + DOWNLOAD
# =========================
st.divider()
st.subheader("Generate Proposal")

# Package inputs
p = ProposalInputs(
    client=client.strip(),
    facility_name=facility.strip(),
    service_begin_date=service_begin_date.strip(),
    service_end_date=service_end_date.strip(),
    service_addresses=addresses,
    days_per_week=int(days),
    cleaning_times=times.strip(),

    num_offices=int(offices),
    num_conference_rooms=int(conference),
    num_break_rooms=int(breaks),
    num_bathrooms=int(baths),

    custom_rooms=st.session_state.get("custom_rooms", []),

    hand_soap=soap,
    paper_towels=towels,
    toilet_paper=tp,

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
