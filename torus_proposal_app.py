import os
import json
import datetime
from io import BytesIO
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional

import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from pypdf import PdfReader
from openai import OpenAI

CHECK = "✓"
TEMPLATE_FILE = "Torus_Template.docx"


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
    custom_rooms: List[Dict[str, int]]

    hand_soap: Optional[str]
    paper_towels: Optional[str]
    toilet_paper: Optional[str]

    include_cover_page: bool
    cover_letter_body: str

    cleaning_plan: str
    notes: str

    compensation_amount: Optional[float]
    compensation_basis: str
    net_terms_days: Optional[int]
    late_interest_percent: Optional[float]

    include_employee_conduct: bool
    include_on_site_storage: bool
    include_compensation_section: bool
    include_modification: bool
    include_access: bool
    include_cancellation: bool

    contractor_printed_name: str
    contractor_title: str


# =========================
# DEFAULT COVER LETTER
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
def get_openai_client():
    key = st.secrets.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("Missing OPENAI_API_KEY in Streamlit secrets.")
    return OpenAI(api_key=key)


def analyze_rfp_with_ai(text: str) -> dict:
    client = get_openai_client()

    instructions = """
Return ONLY valid JSON with this structure:
{
  "cleaning_plan_draft": "string",
  "scope_of_work_draft": "string",
  "schedule_rows": [
    {"task": "string", "daily": true, "weekly": false, "monthly": false}
  ],
  "clarifying_questions": ["string"]
}
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
    name = uploaded_file.name.lower()
    data = uploaded_file.read()

    if name.endswith(".pdf"):
        reader = PdfReader(BytesIO(data))
        return "\n".join(p.extract_text() or "" for p in reader.pages)

    if name.endswith(".docx"):
        doc = Document(BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)

    return data.decode("utf-8", errors="ignore")


# =========================
# WORD HELPERS
# =========================
def add_heading(doc, text):
    p = doc.add_paragraph(text)
    (p.runs[0] if p.runs else p.add_run(text)).bold = True


def add_bullet(doc, text):
    for style in ("List Bullet", "List Paragraph"):
        try:
            doc.add_paragraph(text, style=style)
            return
        except:
            pass
    doc.add_paragraph(f"• {text}")


def add_scope_table(doc, rows):
    add_heading(doc, "SCOPE OF WORK – CLEANING SCHEDULE")
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Task", "Daily", "Weekly", "Monthly"

    for task, d, w, m in rows:
        row = table.add_row().cells
        row[0].text = task
        row[1].text = CHECK if d else ""
        row[2].text = CHECK if w else ""
        row[3].text = CHECK if m else ""


def add_signatures(doc, name, title):
    add_heading(doc, "SIGNATURES")
    doc.add_paragraph("Date: ___________________")
    doc.add_paragraph("__________________________________")
    doc.add_paragraph("Contractor Signature")
    doc.add_paragraph(name)
    doc.add_paragraph(title)

    doc.add_paragraph("")
    doc.add_paragraph("Date: ___________________")
    doc.add_paragraph("__________________________________")
    doc.add_paragraph("Client Signature")
    doc.add_paragraph("Client Printed Name: __________________________")
    doc.add_paragraph("Client Title: _________________________________")


# =========================
# BUILD WORD DOC
# =========================
def build_doc(p: ProposalInputs, schedule_rows):
    doc = Document(TEMPLATE_FILE) if os.path.exists(TEMPLATE_FILE) else Document()

    for s in doc.sections:
        s.different_first_page_header_footer = False

    if p.include_cover_page:
        doc.add_paragraph(p.client)
        doc.add_paragraph("")
        doc.add_paragraph(p.cover_letter_body)
        doc.add_page_break()

    title = doc.add_paragraph("CLEANING SERVICE AGREEMENT")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    doc.add_paragraph(f"Client: {p.client}")
    doc.add_paragraph(f"Facility: {p.facility_name}")

    doc.add_paragraph("Service Addresses:")
    for a in p.service_addresses:
        add_bullet(doc, a)

    doc.add_paragraph(
        f"{p.client} enters into this agreement on __________ for janitorial services "
        f"between {p.service_begin_date} and {p.service_end_date}."
    )

    add_heading(doc, "ROOM COUNTS")
    add_bullet(doc, f"Offices: {p.num_offices}")
    add_bullet(doc, f"Conference rooms: {p.num_conference_rooms}")
    add_bullet(doc, f"Break rooms: {p.num_break_rooms}")
    add_bullet(doc, f"Bathrooms: {p.num_bathrooms}")

    for r in p.custom_rooms:
        if r["type"] and r["count"] > 0:
            add_bullet(doc, f"{r['type']}: {r['count']}")

    add_scope_table(doc, schedule_rows)

    if p.cleaning_plan:
        add_heading(doc, "CLEANING PLAN")
        doc.add_paragraph(p.cleaning_plan)

    add_heading(doc, "GENERAL REQUIREMENTS")
    doc.add_paragraph("Contractor shall provide all labor, supervision, and personnel.")

    if p.hand_soap or p.paper_towels or p.toilet_paper:
        doc.add_paragraph("Consumables:")
        if p.hand_soap:
            doc.add_paragraph(f"- Hand soap: {p.hand_soap}")
        if p.paper_towels:
            doc.add_paragraph(f"- Paper towels: {p.paper_towels}")
        if p.toilet_paper:
            doc.add_paragraph(f"- Toilet paper: {p.toilet_paper}")

    if p.include_compensation_section and p.compensation_amount:
        add_heading(doc, "COMPENSATION")
        doc.add_paragraph(
            f"${p.compensation_amount:,.2f} ({p.compensation_basis})"
        )

    if p.notes:
        add_heading(doc, "NOTES")
        doc.add_paragraph(p.notes)

    add_signatures(doc, p.contractor_printed_name, p.contractor_title)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(layout="wide")
st.title("Torus Group – Cleaning Proposal Builder")

st.session_state.setdefault("schedule_df", pd.DataFrame(
    [("Empty trash", True, False, False)],
    columns=["Task", "Daily", "Weekly", "Monthly"]
))
st.session_state.setdefault("custom_rooms", [{"type": "", "count": 0}])

with st.form("proposal_form"):
    client = st.text_input("Client")
    facility = st.text_input("Facility")
    begin = st.text_input("Service Begin Date")
    end = st.text_input("Service End Date")
    days = st.number_input("Days per week", 1, 7, 5)
    times = st.text_input("Cleaning Times")

    addresses = st.text_area("Addresses (one per line)").splitlines()

    offices = st.number_input("Offices", 0)
    conference = st.number_input("Conference Rooms", 0)
    breaks = st.number_input("Break Rooms", 0)
    baths = st.number_input("Bathrooms", 0)

    for i, r in enumerate(st.session_state["custom_rooms"]):
        r["type"] = st.text_input(f"Room Type {i+1}", r["type"])
        r["count"] = st.number_input(f"Count {i+1}", 0, value=r["count"])

    if st.form_submit_button("Add Room Type"):
        st.session_state["custom_rooms"].append({"type": "", "count": 0})

    schedule_df = st.data_editor(st.session_state["schedule_df"], num_rows="dynamic")
    st.session_state["schedule_df"] = schedule_df

    amount = st.text_input("Compensation Amount")
    basis = st.selectbox("Basis", ["monthly", "annual", "per visit", "one-time clean"])

    submit = st.form_submit_button("Generate Proposal")

if submit:
    p = ProposalInputs(
        client, facility, begin, end,
        [a for a in addresses if a.strip()],
        days, times,
        offices, conference, breaks, baths,
        st.session_state["custom_rooms"],
        None, None, None,
        True, default_cover_letter(client),
        "", "",
        float(amount) if amount else None,
        basis, None, None,
        True, True, True, True, True,
        "Kary Jubilee", "President"
    )

    schedule_rows = [
        (r.Task, r.Daily, r.Weekly, r.Monthly)
        for r in st.session_state["schedule_df"].itertuples()
    ]

    docx = build_doc(p, schedule_rows)
    st.download_button("Download Proposal", docx, file_name="Torus_Proposal.docx")
