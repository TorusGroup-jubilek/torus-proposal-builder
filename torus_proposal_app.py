# torus_proposal_app.py
# Torus Group – Cleaning Service Agreement Builder with AI RFP/PWS Analyzer

import os
import json
import datetime
from io import BytesIO
from dataclasses import dataclass, asdict
from typing import List

import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from pypdf import PdfReader
from openai import OpenAI
from pydantic import BaseModel

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
# AI RFP ANALYSIS (STABLE)
# =========================
class ScheduleRow(BaseModel):
    task: str
    daily: bool
    weekly: bool
    monthly: bool


class RfpAnalysis(BaseModel):
    cleaning_plan_draft: str
    scope_of_work_draft: str
    schedule_rows: List[ScheduleRow]
    clarifying_questions: List[str]


def analyze_rfp_with_ai(text: str) -> dict:
    client = get_openai_client()

    instructions = """
You are assisting a janitorial contractor responding to an RFP or PWS.
Extract requirements and draft:
1) A cleaning plan
2) A scope of work
3) A cleaning schedule (daily/weekly/monthly)
4) Clarifying questions

Return structured data only.
"""

    response = client.responses.parse(
        model="gpt-4o-2024-08-06",
        input=[
            {"role": "system", "content": instructions},
            {"role": "user", "content": text[:120000]},
        ],
        text_format=RfpAnalysis,
        store=False,
    )

    return response.output_parsed.model_dump()


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
def add_cover_page(doc, client, body):
    doc.add_paragraph(client)
    doc.add_paragraph("")
    doc.add_paragraph("Re: Janitorial Services Proposal")
    doc.add_paragraph("")
    doc.add_paragraph(f"Dear {client},")
    doc.add_paragraph("")
    doc.add_paragraph(body)
    doc.add_paragraph("")
    doc.add_paragraph("Respectfully,")
    doc.add_paragraph("Kara Jubilee\nOwner\nTorus Cleaning Services")
    doc.add_page_break()


def add_scope_table(doc, rows):
    p = doc.add_paragraph("SCOPE OF WORK – CLEANING SCHEDULE")
    p.runs[0].bold = True

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr = table.rows[0].cells
    hdr[0].text = "Task"
    hdr[1].text = "Daily"
    hdr[2].text = "Weekly"
    hdr[3].text = "Monthly"

    for r in rows:
        row = table.add_row().cells
        row[0].text = r[0]
        row[1].text = CHECK if r[1] else ""
        row[2].text = CHECK if r[2] else ""
        row[3].text = CHECK if r[3] else ""


def build_doc(p: ProposalInputs, schedule_rows):
    doc = Document("Torus_Template.docx") if os.path.exists("Torus_Template.docx") else Document()

    for s in doc.sections:
        s.different_first_page_header_footer = False

    if p.include_cover_page:
        add_cover_page(doc, p.client, p.cover_letter_body)

    title = doc.add_paragraph("CLEANING SERVICE AGREEMENT")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    doc.add_paragraph(f"Client: {p.client}")
    doc.add_paragraph(f"Facility: {p.facility_name}")
    doc.add_paragraph("Service Addresses:")
    for a in p.service_addresses:
        doc.add_paragraph(a, style="List Bullet")

    doc.add_paragraph("")
    doc.add_paragraph(
        f"Contract period: {p.service_begin_date} to {p.service_end_date}. "
        f"Cleaning {p.days_per_week} days per week between {p.cleaning_times}."
    )

    add_scope_table(doc, schedule_rows)

    if p.cleaning_plan:
        h = doc.add_paragraph("CLEANING PLAN")
        h.runs[0].bold = True
        doc.add_paragraph(p.cleaning_plan)

    h = doc.add_paragraph("GENERAL REQUIREMENTS")
    h.runs[0].bold = True
    doc.add_paragraph(
        "Contractor shall provide all labor, supervision, personnel, and standard cleaning supplies.\n"
        f"Hand soap: {p.hand_soap}\n"
        f"Paper towels: {p.paper_towels}\n"
        f"Toilet paper: {p.toilet_paper}"
    )

    doc.add_paragraph("")
    doc.add_paragraph("NOTES")
    doc.add_paragraph(p.notes or "(none)")

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(layout="wide")
st.title("Torus Group – Cleaning Proposal Builder")

st.sidebar.caption(f"OpenAI key loaded: {bool(st.secrets.get('OPENAI_API_KEY'))}")

client = st.text_input("Client")
facility = st.text_input("Facility name")

c1, c2 = st.columns(2)
with c1:
    start = st.text_input("Service begin date")
    days = st.number_input("Days per week", min_value=1, value=5)
with c2:
    end = st.text_input("Service end date")
    times = st.text_input("Cleaning times")

addresses = st.text_area("Service addresses (one per line)").splitlines()

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
cleaning_plan = st.text_area("Cleaning Plan (optional)")

st.subheader("Notes")
notes = st.text_area("Notes")

# Schedule editor
st.subheader("Cleaning Schedule")
default_rows = [
    ("Empty trash", True, False, False),
    ("Clean restrooms", True, False, False),
    ("Vacuum floors", True, False, False),
    ("Dust surfaces", False, True, False),
    ("Deep clean", False, False, True),
]

df = pd.DataFrame(default_rows, columns=["Task", "Daily", "Weekly", "Monthly"])
edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
schedule_rows = [(r.Task, r.Daily, r.Weekly, r.Monthly) for r in edited.itertuples()]

# AI ANALYZER
st.divider()
st.subheader("RFP / PWS Analyzer (AI)")

uploads = st.file_uploader("Upload RFP/PWS", type=["pdf", "docx", "txt"], accept_multiple_files=True)

if st.button("Analyze with AI") and uploads:
    text = "\n".join(extract_text(f) for f in uploads)
    with st.spinner("Analyzing…"):
        result = analyze_rfp_with_ai(text)

    st.session_state["ai"] = result

if "ai" in st.session_state:
    ai = st.session_state["ai"]
    st.text_area("AI Cleaning Plan", ai["cleaning_plan_draft"], height=150)
    st.text_area("AI Scope", ai["scope_of_work_draft"], height=150)

    if st.button("Apply AI to proposal"):
        cleaning_plan = ai["cleaning_plan_draft"]
        schedule_rows = [(r["task"], r["daily"], r["weekly"], r["monthly"]) for r in ai["schedule_rows"]]
        st.success("Applied.")

# BUILD PROPOSAL
p = ProposalInputs(
    client, facility, start, end, addresses, days, times, 30, 0.0,
    0, "", offices, conference, breaks, baths,
    soap, towels, tp,
    "Monthly", 0.0,
    include_cover, cover_body, cleaning_plan, notes
)

docx = build_doc(p, schedule_rows)

st.download_button(
    "Download Word Proposal",
    data=docx,
    file_name=f"Torus_Cleaning_Agreement_{datetime.date.today()}.docx",
)
