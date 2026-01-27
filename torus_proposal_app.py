# torus_proposal_app.py
# Torus Group — Cleaning Service Agreement Builder + Optional Cover Page + AI RFP/PWS Analyzer
#
# ✅ Template support (proposal_template.docx)
# ✅ Optional cover page toggle
# ✅ Multi-address bullets in Word
# ✅ Dynamic Scope of Work schedule table (editable)
# ✅ Optional Cleaning Plan section
# ✅ AI RFP/PWS upload + analysis + “Apply to Proposal” (requires OPENAI_API_KEY in Streamlit secrets)

import datetime
import json
import os
from dataclasses import dataclass, asdict
from io import BytesIO
from typing import List, Dict, Any

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
from pypdf import PdfReader

COMPANY_NAME = "Torus Group"
CHECK = "✓"


# =========================
# Data Model
# =========================
@dataclass
class ProposalInputs:
    # Agreement / client
    client: str
    facility_name: str
    service_begin_date: str
    service_end_date: str
    service_addresses: List[str]
    days_per_week: int
    cleaning_times: str
    net_terms: int

    # Tax
    sales_tax_percent: float

    # Facility details
    space_type: str
    square_footage: int
    floor_types: str

    # Room counts
    num_offices: int
    num_conference_rooms: int
    num_break_rooms: int
    num_bathrooms: int
    num_kitchens: int
    num_locker_rooms: int

    # Ops
    cleaning_frequency: str
    day_porter_needed: str  # Yes/No
    trash_pickup: str
    restocking_needed: str  # Yes/No

    # Consumables
    hand_soap: str
    paper_towels: str
    toilet_paper: str

    # Pricing
    pricing_mode: str  # Monthly Fixed | Per Sq Ft | Per Visit
    monthly_fixed_price: float
    rate_per_sqft: float
    rate_per_visit: float
    visits_per_week: float
    visits_per_month: int  # computed

    # Deep clean
    deep_clean_option: str  # None | One-time | Quarterly
    deep_clean_price: float
    deep_clean_includes: List[str]

    # Add-ons
    additional_services: List[Dict[str, float]]
    include_addons_in_total: str  # Yes/No

    # Compensation
    compensation_mode: str  # Auto (calculated) | Override
    compensation_override: float

    # Optional section
    cleaning_plan: str

    # Cover page options
    include_cover_page: bool
    cover_letter_body: str

    # Notes
    notes: str


# =========================
# Helpers
# =========================
def money(x: float) -> str:
    return f"${x:,.2f}"


def compute_visits_per_month(visits_per_week: float) -> int:
    return int(round(float(visits_per_week) * (52.0 / 12.0)))


def clean_list(items: List[str]) -> List[str]:
    out = []
    for s in items:
        s2 = (s or "").strip()
        if s2:
            out.append(s2)
    return out


def _has_keyword(text: str, keywords) -> bool:
    t = (text or "").lower()
    return any(k.lower() in t for k in keywords)


def build_totals(p: ProposalInputs) -> dict:
    if p.pricing_mode == "Monthly Fixed":
        base_monthly = float(p.monthly_fixed_price)
        base_explain = f"Monthly fixed price: {money(base_monthly)}"
    elif p.pricing_mode == "Per Sq Ft":
        base_monthly = float(p.rate_per_sqft) * float(p.square_footage)
        base_explain = f"Rate: {money(p.rate_per_sqft)}/sqft × {p.square_footage:,} sqft = {money(base_monthly)} per month"
    else:
        base_monthly = float(p.rate_per_visit) * float(p.visits_per_month)
        base_explain = (
            f"Rate: {money(p.rate_per_visit)}/visit × {p.visits_per_month} visits/month "
            f"({p.visits_per_week:g}/week) = {money(base_monthly)} per month"
        )

    addons_total = 0.0
    addons_lines = []
    for item in p.additional_services:
        name = str(item.get("name", "")).strip()
        price = float(item.get("price", 0.0) or 0.0)
        if name and price > 0:
            addons_total += price
            addons_lines.append(f"• {name}: {money(price)}")

    include_addons = (p.include_addons_in_total == "Yes")
    addons_included_monthly = addons_total if include_addons else 0.0

    deep_clean_one_time = 0.0
    deep_clean_quarterly = 0.0
    deep_clean_monthly_equiv = 0.0
    if p.deep_clean_option == "One-time":
        deep_clean_one_time = float(p.deep_clean_price)
    elif p.deep_clean_option == "Quarterly":
        deep_clean_quarterly = float(p.deep_clean_price)
        deep_clean_monthly_equiv = deep_clean_quarterly / 3.0

    monthly_subtotal = base_monthly + addons_included_monthly + deep_clean_monthly_equiv

    tax_rate = max(0.0, float(p.sales_tax_percent)) / 100.0
    monthly_tax = monthly_subtotal * tax_rate
    monthly_total_with_tax = monthly_subtotal + monthly_tax

    if p.compensation_mode == "Override":
        compensation_monthly = float(p.compensation_override)
        compensation_explain = f"Compensation (override): {money(compensation_monthly)}"
    else:
        compensation_monthly = monthly_total_with_tax
        compensation_explain = f"Compensation (calculated): {money(compensation_monthly)}"

    return {
        "base_monthly": base_monthly,
        "base_explain": base_explain,
        "addons_total": addons_total,
        "addons_lines": addons_lines,
        "include_addons": include_addons,
        "addons_included_monthly": addons_included_monthly,
        "deep_clean_one_time": deep_clean_one_time,
        "deep_clean_quarterly": deep_clean_quarterly,
        "deep_clean_monthly_equiv": deep_clean_monthly_equiv,
        "monthly_subtotal": monthly_subtotal,
        "monthly_tax": monthly_tax,
        "monthly_total_with_tax": monthly_total_with_tax,
        "compensation_monthly": compensation_monthly,
        "compensation_explain": compensation_explain,
    }


# =========================
# Schedule: recommended rows + editable table
# =========================
def compute_cleaning_schedule(p: ProposalInputs) -> list:
    d = int(p.days_per_week or 0)

    has_carpet = _has_keyword(p.floor_types, ["carpet"])
    has_vct = _has_keyword(p.floor_types, ["vct", "vinyl"])
    has_epoxy = _has_keyword(p.floor_types, ["epoxy"])
    has_tile = _has_keyword(p.floor_types, ["tile"])
    hard_floors = has_vct or has_epoxy or has_tile or _has_keyword(p.floor_types, ["hard", "concrete"])

    rows = []
    rows.append(("Empty trash & replace liners", CHECK if d >= 3 else "", CHECK if d in (1, 2) else "", ""))
    rows.append(("Clean/disinfect high-touch points (handles, switches, rails)", CHECK if d >= 3 else "", CHECK if d in (1, 2) else "", ""))

    if int(p.num_bathrooms or 0) > 0:
        rows.append(("Clean & disinfect restrooms; restock as applicable", CHECK if d >= 3 else "", CHECK if d in (1, 2) else "", ""))

    if int(p.num_break_rooms or 0) > 0 or int(p.num_kitchens or 0) > 0:
        rows.append(("Break rooms/kitchens: counters, sinks, exterior appliances", CHECK if d >= 5 else "", CHECK if d in (1, 2, 3, 4) else "", ""))

    rows.append(("Dust horizontal surfaces (accessible)", CHECK if d >= 5 else "", CHECK if d in (1, 2, 3, 4) else "", ""))
    rows.append(("Spot clean glass & mirrors (interior)", CHECK if d >= 5 else "", CHECK if d in (1, 2, 3, 4) else "", ""))

    if has_carpet:
        rows.append(("Vacuum carpeted areas (as applicable)", CHECK if d >= 3 else "", CHECK if d in (1, 2) else "", ""))
        rows.append(("Spot treat carpet stains (as needed)", CHECK if d >= 5 else "", CHECK if d in (1, 2, 3, 4) else "", ""))
        rows.append(("Carpet extraction / shampoo (as scheduled)", "", "", CHECK))

    if hard_floors:
        rows.append(("Damp mop hard floors (as applicable)", CHECK if d >= 5 else "", CHECK if d in (1, 2, 3, 4) else "", ""))
        rows.append(("Detail floor scrubbing / machine scrub", "", CHECK if d >= 3 else "", CHECK))

    rows.append(("High dusting (vents, ledges, corners)", "", "", CHECK))
    rows.append(("Baseboards / detail edges (as applicable)", "", "", CHECK))

    if has_vct:
        rows.append(("VCT maintenance (buff/burnish if applicable)", "", "", CHECK))
        rows.append(("Strip & wax (as quoted/needed)", "", "", CHECK))

    if int(p.num_locker_rooms or 0) > 0:
        rows.append(("Locker rooms: clean/disinfect & mop", CHECK if d >= 3 else "", CHECK if d in (1, 2) else "", ""))

    if p.day_porter_needed == "Yes":
        rows.append(("Day porter tasks (restroom checks, spills, touch-ups)", CHECK, "", ""))

    if p.deep_clean_option != "None":
        rows.append(("Deep clean tasks (per agreement)", "", "", CHECK))

    return rows


def schedule_rows_to_df(rows: list) -> pd.DataFrame:
    return pd.DataFrame(
        [(t, d == CHECK, w == CHECK, m == CHECK) for (t, d, w, m) in rows],
        columns=["Task", "Daily", "Weekly", "Monthly"],
    )


def df_to_schedule_rows(df: pd.DataFrame) -> list:
    rows = []
    for _, r in df.iterrows():
        task = str(r.get("Task", "")).strip()
        if not task:
            continue
        daily = CHECK if bool(r.get("Daily", False)) else ""
        weekly = CHECK if bool(r.get("Weekly", False)) else ""
        monthly = CHECK if bool(r.get("Monthly", False)) else ""
        rows.append((task, daily, weekly, monthly))
    return rows


# =========================
# Word helpers
# =========================
def add_scope_of_work_table(doc: Document, schedule_rows: list):
    title_p = doc.add_paragraph()
    title_run = title_p.add_run("SCOPE OF WORK – CLEANING SCHEDULE")
    title_run.bold = True

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"

    hdr = table.rows[0].cells
    hdr[0].text = "Task"
    hdr[1].text = "Daily"
    hdr[2].text = "Weekly"
    hdr[3].text = "Monthly"

    for task, daily, weekly, monthly in schedule_rows:
        row = table.add_row().cells
        row[0].text = task
        row[1].text = daily
        row[2].text = weekly
        row[3].text = monthly

    doc.add_paragraph("")


def add_address_bullets(doc: Document, addresses: List[str]):
    addresses = clean_list(addresses)
    if not addresses:
        doc.add_paragraph("(not provided)")
        return
    for a in addresses:
        doc.add_paragraph(a, style="List Bullet")


def add_cover_page(doc: Document, client_name: str, body: str):
    # Letterhead style should be in the template header. This function adds the cover letter content.
    paragraphs = [
        client_name,
        "",
        "Attn: ______________________",
        "",
        "Re: Janitorial Services Proposal",
        "",
        f"Dear {client_name},",
        "",
        body if body else "",
        "",
        "Respectfully,",
        "",
        "Kara Jubilee",
        "Owner",
        "Torus Cleaning Services",
    ]
    for line in paragraphs:
        doc.add_paragraph(line)
    doc.add_page_break()


# =========================
# Agreement text
# =========================
def build_agreement_text(p: ProposalInputs) -> str:
    totals = build_totals(p)
    addresses = clean_list(p.service_addresses)
    address_inline = "; ".join(addresses) if addresses else "(service address not provided)"
    today = datetime.date.today().strftime("%B %d, %Y")
    floor_note = p.floor_types.strip() if p.floor_types.strip() else "N/A"

    para1 = (
        f"{p.client}, ('Client'), enters into this agreement on this date "
        f"______________ for Torus Cleaning Services ('Contractor'), to provide "
        f"janitorial services for facility/facilities located at the following locations: {address_inline}"
    )
    para2 = (
        f"Contractor shall provide janitorial services {p.days_per_week} per week between the hours of "
        f"{p.cleaning_times} for the facility/facilities located at {address_inline}."
    )
    para3 = f"The contract period is as follows {p.service_begin_date} to {p.service_end_date}."

    cleaning_plan_block = ""
    if p.cleaning_plan.strip():
        cleaning_plan_block = (
            "CLEANING PLAN\n"
            f"{p.cleaning_plan.strip()}\n"
        )

    general_requirements_block = (
        "GENERAL REQUIREMENTS\n"
        "Contractor shall provide all labor, supervision, and personnel necessary to perform the janitorial services "
        "described in this agreement.\n\n"
        "Unless otherwise stated below, all cleaning equipment and standard janitorial supplies required to perform "
        "the services shall be provided by the Contractor.\n\n"
        "Consumable supplies:\n"
        f"• Hand soap: {p.hand_soap}\n"
        f"• Paper towels: {p.paper_towels}\n"
        f"• Toilet paper: {p.toilet_paper}\n"
    )

    insurance_block = (
        "INSURANCE & LIABILITY\n"
        "Contractor shall maintain insurance customary for janitorial service providers, including general liability "
        "and workers’ compensation as required by law.\n\n"
        "Upon request, Contractor may provide a certificate of insurance.\n\n"
        "Each party shall be responsible for its own acts and omissions and those of its employees and subcontractors.\n"
    )

    access_security_block = (
        "ACCESS & SECURITY\n"
        "Client shall provide Contractor with reasonable access to the facility/facilities during the agreed cleaning times, "
        "including access to water and electrical service as needed.\n\n"
        "If keys, fobs, alarm codes, or badges are issued, Contractor will take reasonable care to safeguard them "
        "and will return them upon termination of this agreement.\n\n"
        "Client shall notify Contractor of any site-specific security procedures, restricted areas, or check-in/check-out requirements.\n"
    )

    damages_exclusions_block = (
        "DAMAGES, CLIENT PROPERTY & EXCLUSIONS\n"
        "Contractor shall exercise reasonable care while performing services. Client agrees to secure or remove fragile, "
        "high-value, or sensitive items. Contractor is not responsible for normal wear and tear.\n\n"
        "Contractor is not responsible for pre-existing conditions (including but not limited to stained carpet, damaged flooring, "
        "peeling finishes, or cracked tile) or damage resulting from defective surfaces/materials.\n\n"
        "Services do not include hazardous materials handling, mold remediation, biohazard cleanup, or specialized restoration work "
        "unless specifically listed in writing as an additional service.\n"
    )

    termination_block = (
        "TERM, TERMINATION & CHANGES\n"
        "This agreement remains in effect for the contract period stated above unless terminated earlier in accordance with this section.\n\n"
        "Either party may terminate this agreement with written notice. Unless otherwise agreed in writing, a notice period of 30 days applies.\n\n"
        "Client may request changes to scope, frequency, or locations. Any material change may require a written adjustment to pricing.\n\n"
        "Contractor may suspend services for non-payment after providing written notice and a reasonable opportunity to cure.\n"
    )

    payment_block = (
        "PAYMENT TERMS, TAXES & LATE FEES\n"
        f"Payment terms are Net {p.net_terms}. Sales tax will be applied where required at {p.sales_tax_percent:.2f}%.\n\n"
        "Past due balances may be subject to a late charge of 1.5% per month (or the maximum allowed by law, whichever is less), "
        "plus reasonable collection costs.\n"
    )

    entire_agreement_block = (
        "ENTIRE AGREEMENT\n"
        "This document constitutes the entire agreement between the parties regarding the services described and supersedes all prior "
        "discussions or representations. Any amendments must be in writing and signed by both parties.\n"
    )

    pricing_block = (
        "PRICING & COMPENSATION\n"
        f"Base service pricing: {totals['base_explain']}\n"
        f"Additional services total: {money(totals['addons_total'])} "
        f"({'included in monthly total' if totals['include_addons'] else 'not included in monthly total'})\n"
        f"Monthly subtotal (pre-tax): {money(totals['monthly_subtotal'])}\n"
        f"Sales tax ({p.sales_tax_percent:.2f}%): {money(totals['monthly_tax'])}\n"
        f"Monthly total (with tax): {money(totals['monthly_total_with_tax'])}\n"
        f"{totals['compensation_explain']}\n"
    )

    deep_clean_block = ""
    if p.deep_clean_option != "None":
        if p.deep_clean_option == "One-time":
            deep_clean_block += f"Deep clean (one-time): {money(totals['deep_clean_one_time'])}\n"
        else:
            deep_clean_block += f"Deep clean (quarterly): {money(totals['deep_clean_quarterly'])} per quarter\n"
        if p.deep_clean_includes:
            deep_clean_block += "Deep clean includes:\n" + "\n".join([f"• {x}" for x in p.deep_clean_includes]) + "\n"

    add_on_block = ""
    if totals["addons_lines"]:
        add_on_block = "ADDITIONAL SERVICES (LINE ITEMS)\n" + "\n".join(totals["addons_lines"]) + "\n"

    return f"""
\n\nCLEANING SERVICE AGREEMENT

Date prepared: {today}

Client / Location Name: {p.client}
Facility/Location Name: {p.facility_name}
Service Address(es):
SERVICE_ADDRESSES_BULLETS

{para1}

{para2}

{para3}

SCOPE_OF_WORK_TABLE

{cleaning_plan_block}{general_requirements_block}

{pricing_block}
{deep_clean_block}
{add_on_block}

Facility details:
• Space type: {p.space_type}
• Approx. square footage: {p.square_footage:,} sqft
• Floor types/notes: {floor_note}

{insurance_block}

{access_security_block}

{damages_exclusions_block}

{termination_block}

{payment_block}

{entire_agreement_block}

NOTES
{p.notes if p.notes.strip() else "(none)"}

ACCEPTANCE

Client Authorized Signature: _______________________________    Date: _______________

Contractor Authorized Signature: ___________________________   Date: _______________
""".strip()


def docx_from_agreement(text: str, schedule_rows: list, addresses: List[str], p: ProposalInputs) -> bytes:
    template_path = "Torus_Template.docx"
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    # Prevent blank header on cover page if template has "Different First Page"
    for section in doc.sections:
        section.different_first_page_header_footer = False

    # Optional cover page
    if p.include_cover_page:
        add_cover_page(doc, p.client, p.cover_letter_body)

    # Write agreement content
    for line in text.splitlines():
        s = line.strip()

        if not s:
            doc.add_paragraph("")
            continue

        if s == "CLEANING SERVICE AGREEMENT":
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(s)
            run.bold = True
            continue

        if s == "SERVICE_ADDRESSES_BULLETS":
            add_address_bullets(doc, addresses)
            continue

        if s == "SCOPE_OF_WORK_TABLE":
            add_scope_of_work_table(doc, schedule_rows)
            continue

        # Bold section headings (all caps)
        if s.isupper() and len(s) <= 60:
            para = doc.add_paragraph()
            run = para.add_run(s)
            run.bold = True
            continue

        doc.add_paragraph(s)

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# =========================
# RFP / PWS AI Analyzer
# =========================
def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(BytesIO(file_bytes))
    parts = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return "\n".join(parts).strip()


def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(BytesIO(file_bytes))
    return "\n".join([p.text for p in doc.paragraphs]).strip()


def extract_text_from_upload(uploaded_file) -> str:
    name = (uploaded_file.name or "").lower()
    data = uploaded_file.read()

    if name.endswith(".pdf"):
        return extract_text_from_pdf(data)
    if name.endswith(".docx"):
        return extract_text_from_docx(data)

    # txt or fallback
    try:
        return data.decode("utf-8", errors="ignore").strip()
    except Exception:
        return ""


def get_openai_client() -> OpenAI:
    key = None
    try:
        key = st.secrets.get("OPENAI_API_KEY")
    except Exception:
        key = None

    if not key:
        raise RuntimeError("Missing OPENAI_API_KEY in Streamlit secrets.")
    return OpenAI(api_key=key)


def analyze_rfp_with_ai(rfp_text: str, company_profile: str = "") -> Dict[str, Any]:
    client = get_openai_client()

    # JSON schema for consistent output
    schema = {
        "name": "rfp_analysis",
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "opportunity_summary": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "title": {"type": "string"},
                        "agency_or_client": {"type": "string"},
                        "submission_deadline": {"type": "string"},
                        "submission_method": {"type": "string"},
                        "site_visit_required": {"type": "boolean"},
                        "period_of_performance": {"type": "string"},
                    },
                    "required": [
                        "title",
                        "agency_or_client",
                        "submission_deadline",
                        "submission_method",
                        "site_visit_required",
                        "period_of_performance",
                    ],
                },
                "compliance_checklist": {"type": "array", "items": {"type": "string"}},
                "cleaning_plan_draft": {"type": "string"},
                "scope_of_work_draft": {"type": "string"},
                "schedule_rows": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "properties": {
                            "task": {"type": "string"},
                            "daily": {"type": "boolean"},
                            "weekly": {"type": "boolean"},
                            "monthly": {"type": "boolean"},
                        },
                        "required": ["task", "daily", "weekly", "monthly"],
                    },
                },
                "add_on_suggestions": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "properties": {
                            "name": {"type": "string"},
                            "notes": {"type": "string"},
                        },
                        "required": ["name", "notes"],
                    },
                },
                "clarifying_questions": {"type": "array", "items": {"type": "string"}},
                "assumptions": {"type": "array", "items": {"type": "string"}},
            },
            "required": [
                "opportunity_summary",
                "compliance_checklist",
                "cleaning_plan_draft",
                "scope_of_work_draft",
                "schedule_rows",
                "add_on_suggestions",
                "clarifying_questions",
                "assumptions",
            ],
        },
    }

    instructions = f"""
You are assisting a janitorial contractor preparing an RFP/PWS response.
Extract mandatory requirements, compliance items, and draft a cleaning plan and scope of work.
Also propose a cleaning schedule table (daily/weekly/monthly) suitable for a janitorial SOW.
Return ONLY JSON that matches the provided schema.

Company profile/context:
{company_profile}
""".strip()

    payload = rfp_text[:120000]  # keep bounded

    resp = client.responses.create(
        model="gpt-4.1-mini",
        input=[
            {"role": "system", "content": instructions},
            {"role": "user", "content": payload},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "rfp_analysis",
                "schema": "schema",
            }
        },
        store=False,
    )

    return json.loads(resp.output_text)


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title=f"{COMPANY_NAME} Agreement Builder", layout="wide")
st.title(f"{COMPANY_NAME} — Cleaning Service Agreement Builder")

st.sidebar.caption(f"OpenAI key loaded: {bool(st.secrets.get('OPENAI_API_KEY', ''))}")

# Quick template indicator
st.sidebar.caption(f"Template found: {os.path.exists('proposal_template.docx')}")

# Session defaults for prefills
st.session_state.setdefault("cover_letter_body_prefill", "")
st.session_state.setdefault("cleaning_plan_prefill", "")
st.session_state.setdefault("notes_prefill", "")

with st.sidebar:
    st.header("Client / Contract")
    client_name = st.text_input("Client (legal name)", value="")
    facility_name = st.text_input("Facility/Location name", value="")

    service_begin_date = st.text_input("Service begin date")
    service_end_date = st.text_input("Service end date")

    net_terms = st.selectbox("Net pay terms", [15, 30, 45, 60], index=1)
    days_per_week = st.number_input("Number of days per week", min_value=0, step=1, value=5)
    cleaning_times = st.text_input("Cleaning times (e.g., 6:00 PM – 10:00 PM)")

    st.header("Sales Tax")
    sales_tax_percent = st.number_input("Sales tax (%)", min_value=0.0, step=0.25, value=0.0)

    st.header("Operations")
    cleaning_frequency = st.text_input("Cleaning frequency (e.g., 5x/week)", value="5x/week")
    trash_pickup = st.text_input("Trash pickup schedule", value="daily")
    day_porter_needed = st.selectbox("Day porter needed", ["No", "Yes"])
    restocking_needed = st.selectbox("Restocking needed", ["No", "Yes"])

    st.header("Consumable Supplies")
    hand_soap = st.selectbox("Hand soap", ["Provided by Contractor", "Provided by Client"])
    paper_towels = st.selectbox("Paper towels", ["Provided by Contractor", "Provided by Client"])
    toilet_paper = st.selectbox("Toilet paper", ["Provided by Contractor", "Provided by Client"])

    st.header("Cover Page")
    include_cover_page = st.checkbox("Include cover page", value=True)

c1, c2, c3 = st.columns(3)

with c1:
    st.subheader("Facility details")
    space_type = st.text_input("Type of space (Office/Medical/etc.)", value="Office")
    square_footage = st.number_input("Square footage", min_value=0, step=100, value=0)
    floor_types = st.text_area(
        "Floor types (optional)",
        placeholder="Carpet 7,600 sqft; VCT 30,000 sqft; Epoxy 50,000 sqft"
    )

with c2:
    st.subheader("Room counts")
    num_offices = st.number_input("Offices", min_value=0, step=1, value=0)
    num_conference_rooms = st.number_input("Conference rooms", min_value=0, step=1, value=0)
    num_break_rooms = st.number_input("Break rooms", min_value=0, step=1, value=0)
    num_bathrooms = st.number_input("Bathrooms", min_value=0, step=1, value=0)
    num_kitchens = st.number_input("Kitchens", min_value=0, step=1, value=0)
    num_locker_rooms = st.number_input("Locker rooms", min_value=0, step=1, value=0)

with c3:
    st.subheader("Pricing")
    pricing_mode = st.selectbox("Pricing method", ["Monthly Fixed", "Per Sq Ft", "Per Visit"])

    monthly_fixed_price = 0.0
    rate_per_sqft = 0.0
    rate_per_visit = 0.0
    visits_per_week = 0.0
    visits_per_month = 0

    if pricing_mode == "Monthly Fixed":
        monthly_fixed_price = st.number_input("Monthly fixed price ($)", min_value=0.0, step=50.0, value=0.0)
    elif pricing_mode == "Per Sq Ft":
        rate_per_sqft = st.number_input(
            "Rate per sq ft ($/sqft)",
            min_value=0.0,
            step=0.001,
            format="%.4f",
            value=0.0000
        )
        st.caption("Calculates monthly base: rate × square footage.")
    else:
        rate_per_visit = st.number_input("Rate per visit ($/visit)", min_value=0.0, step=25.0, value=0.0)
        visits_per_week = st.number_input("Visits per week", min_value=0.0, step=0.5, value=5.0)
        visits_per_month = compute_visits_per_month(float(visits_per_week))
        st.caption(f"Estimated visits/month: {visits_per_month} (based on {visits_per_week:g}/week)")

    st.subheader("Deep clean")
    deep_clean_option = st.selectbox("Deep clean option", ["None", "One-time", "Quarterly"])
    deep_clean_price = 0.0
    deep_clean_includes: List[str] = []

    if deep_clean_option != "None":
        deep_clean_price = st.number_input("Deep clean price ($)", min_value=0.0, step=50.0, value=0.0)
        st.caption("Choose what the deep clean includes:")

        d1, d2 = st.columns(2)
        with d1:
            if st.checkbox("Carpet extraction"):
                deep_clean_includes.append("Carpet extraction (as applicable)")
            if st.checkbox("Strip & wax (VCT)"):
                deep_clean_includes.append("Strip & wax for VCT (as applicable)")
            if st.checkbox("High dusting"):
                deep_clean_includes.append("High dusting (vents, ledges, hard-to-reach areas)")
        with d2:
            if st.checkbox("Disinfection service"):
                deep_clean_includes.append("Disinfection of high-touch and common areas")
            if st.checkbox("Detail floor scrubbing"):
                deep_clean_includes.append("Detail floor scrubbing / machine scrub (as applicable)")
            if st.checkbox("Glass detailing"):
                deep_clean_includes.append("Glass detailing (interior as applicable)")

    st.subheader("Cover Letter (optional)")
    cover_letter_body = st.text_area(
        "Cover Letter Body",
        height=220,
        value=st.session_state.get("cover_letter_body_prefill", ""),
        placeholder="This text will appear on the cover page introduction letter..."
    )

    st.subheader("Cleaning Plan (optional)")
    cleaning_plan = st.text_area(
        "Cleaning Plan",
        height=160,
        value=st.session_state.get("cleaning_plan_prefill", ""),
        placeholder="Add any special instructions, priorities, staffing approach, or site-specific plan..."
    )

    st.subheader("Notes")
    notes = st.text_area(
        "Notes (optional)",
        height=140,
        value=st.session_state.get("notes_prefill", "")
    )

st.divider()

# Multi-address UI
st.subheader("Service Address(es)")
st.caption("Add one or more addresses for this agreement.")

if "service_addresses" not in st.session_state:
    st.session_state.service_addresses = [""]

for i, addr in enumerate(st.session_state.service_addresses):
    ca, cb = st.columns([6, 1])
    with ca:
        st.session_state.service_addresses[i] = st.text_input(
            f"service_addr_{i}",
            value=addr,
            label_visibility="collapsed",
            placeholder="Street, City, State ZIP"
        )
    with cb:
        if st.button("Remove", key=f"remove_addr_{i}") and len(st.session_state.service_addresses) > 1:
            st.session_state.service_addresses.pop(i)
            st.rerun()

if st.button("Add another address"):
    st.session_state.service_addresses.append("")
    st.rerun()

st.divider()

# Add-ons UI
st.subheader("Additional services (add-ons)")
st.caption("Add line items (example: Day porter hours, Event cleanup, Carpet extraction add-on).")

if "addons" not in st.session_state:
    st.session_state.addons = [{"name": "", "price": 0.0}]

for i, item in enumerate(st.session_state.addons):
    ca, cb, cc = st.columns([3, 1, 1])
    with ca:
        st.session_state.addons[i]["name"] = st.text_input(
            f"addon_name_{i}",
            value=item["name"],
            placeholder="Service name",
            label_visibility="collapsed"
        )
    with cb:
        st.session_state.addons[i]["price"] = st.number_input(
            f"addon_price_{i}",
            min_value=0.0,
            step=25.0,
            value=float(item["price"]),
            label_visibility="collapsed"
        )
    with cc:
        if st.button("Remove", key=f"remove_addon_{i}") and len(st.session_state.addons) > 1:
            st.session_state.addons.pop(i)
            st.rerun()

cbtn1, cbtn2, _ = st.columns([1, 2, 3])
with cbtn1:
    if st.button("Add another add-on"):
        st.session_state.addons.append({"name": "", "price": 0.0})
        st.rerun()
with cbtn2:
    include_addons_in_total = st.selectbox("Include add-ons in monthly total?", ["Yes", "No"], index=0)

st.divider()

# Compensation UI
st.subheader("Compensation")
compensation_mode = st.selectbox("Compensation mode", ["Auto (calculated)", "Override"])
compensation_override = 0.0
if compensation_mode == "Override":
    compensation_override = st.number_input("Compensation override ($ per month)", min_value=0.0, step=50.0, value=0.0)

# Build ProposalInputs
p = ProposalInputs(
    client=client_name.strip(),
    facility_name=facility_name.strip(),
    service_begin_date=service_begin_date.strip(),
    service_end_date=service_end_date.strip(),
    service_addresses=st.session_state.service_addresses,
    days_per_week=int(days_per_week),
    cleaning_times=cleaning_times.strip(),
    net_terms=int(net_terms),
    sales_tax_percent=float(sales_tax_percent),
    space_type=space_type.strip(),
    square_footage=int(square_footage),
    floor_types=floor_types.strip(),
    num_offices=int(num_offices),
    num_conference_rooms=int(num_conference_rooms),
    num_break_rooms=int(num_break_rooms),
    num_bathrooms=int(num_bathrooms),
    num_kitchens=int(num_kitchens),
    num_locker_rooms=int(num_locker_rooms),
    cleaning_frequency=cleaning_frequency.strip(),
    day_porter_needed=day_porter_needed,
    trash_pickup=trash_pickup.strip(),
    restocking_needed=restocking_needed,
    hand_soap=hand_soap,
    paper_towels=paper_towels,
    toilet_paper=toilet_paper,
    pricing_mode=pricing_mode,
    monthly_fixed_price=float(monthly_fixed_price),
    rate_per_sqft=float(rate_per_sqft),
    rate_per_visit=float(rate_per_visit),
    visits_per_week=float(visits_per_week),
    visits_per_month=int(visits_per_month),
    deep_clean_option=deep_clean_option,
    deep_clean_price=float(deep_clean_price),
    deep_clean_includes=deep_clean_includes,
    additional_services=st.session_state.addons,
    include_addons_in_total=include_addons_in_total,
    compensation_mode=compensation_mode,
    compensation_override=float(compensation_override),
    cleaning_plan=cleaning_plan.strip(),
    include_cover_page=include_cover_page,
    cover_letter_body=cover_letter_body.strip(),
    notes=notes.strip(),
)

# Totals display
totals = build_totals(p)
st.subheader("Calculated totals")
t1, t2, t3, t4 = st.columns(4)
t1.metric("Monthly subtotal (pre-tax)", money(totals["monthly_subtotal"]))
t2.metric("Sales tax (monthly)", money(totals["monthly_tax"]))
t3.metric("Monthly total (with tax)", money(totals["monthly_total_with_tax"]))
t4.metric("Compensation (monthly)", money(totals["compensation_monthly"]))

# Schedule tuning
st.divider()
st.subheader("Scope of Work — Schedule Tuning")
st.caption("Edit tasks/frequencies here. The Word schedule table will match what you set.")

default_rows = compute_cleaning_schedule(p)
default_df = schedule_rows_to_df(default_rows)

if "schedule_df" not in st.session_state:
    st.session_state.schedule_df = default_df.copy()

if st.button("Reset schedule to recommended defaults"):
    st.session_state.schedule_df = default_df.copy()
    st.rerun()

edited_df = st.data_editor(
    st.session_state.schedule_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Task": st.column_config.TextColumn("Task"),
        "Daily": st.column_config.CheckboxColumn("Daily"),
        "Weekly": st.column_config.CheckboxColumn("Weekly"),
        "Monthly": st.column_config.CheckboxColumn("Monthly"),
    },
)
st.session_state.schedule_df = edited_df
tuned_schedule_rows = df_to_schedule_rows(edited_df)

# =========================
# RFP / PWS Analyzer (AI)
# =========================
st.divider()
st.subheader("RFP / PWS Analyzer (AI)")
st.caption("Upload an RFP/PWS to extract requirements and draft a response you can apply into this proposal.")

rfp_files = st.file_uploader(
    "Upload RFP/PWS files (PDF, DOCX, TXT)",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True
)

company_profile = st.text_area(
    "Optional: company profile context",
    height=120,
    placeholder="Example: Licensed/insured; background-checked staff; QA inspections; SDVOSB; etc."
)

col_an1, col_an2 = st.columns([1, 2])
with col_an1:
    run_analysis = st.button("Analyze RFP/PWS with AI")
with col_an2:
    if st.button("Clear analysis / prefills"):
        for k in ["rfp_analysis", "cleaning_plan_prefill", "notes_prefill", "cover_letter_body_prefill"]:
            if k in st.session_state:
                del st.session_state[k]
        st.success("Cleared.")
        st.rerun()

if run_analysis:
    if not rfp_files:
        st.error("Please upload at least one RFP/PWS file.")
    else:
        combined = []
        for f in rfp_files:
            text = extract_text_from_upload(f)
            if text:
                combined.append(f"\n\n===== {f.name} =====\n{text}")
        rfp_text = "\n".join(combined).strip()

        if not rfp_text:
            st.error("Could not extract text from the uploads (PDF may be scanned).")
        else:
            try:
                with st.spinner("Analyzing with AI..."):
                    analysis = analyze_rfp_with_ai(rfp_text, company_profile=company_profile)
                st.session_state["rfp_analysis"] = analysis
                st.success("Analysis complete.")
            except Exception as e:
                st.error(f"AI analysis failed: {e}")

analysis = st.session_state.get("rfp_analysis")
if analysis:
    st.subheader("Opportunity Summary")
    st.json(analysis["opportunity_summary"])

    st.subheader("Compliance Checklist")
    st.write("\n".join([f"- {x}" for x in analysis["compliance_checklist"]]) or "(none)")

    st.subheader("Drafts")
    st.text_area("AI Cleaning Plan draft", analysis["cleaning_plan_draft"], height=180)
    st.text_area("AI Scope of Work draft", analysis["scope_of_work_draft"], height=180)

    st.subheader("Clarifying Questions")
    st.write("\n".join([f"- {x}" for x in analysis["clarifying_questions"]]) or "(none)")

    st.subheader("Assumptions")
    st.write("\n".join([f"- {x}" for x in analysis["assumptions"]]) or "(none)")

    st.subheader("Add-on Suggestions")
    for s in analysis["add_on_suggestions"]:
        st.write(f"- **{s['name']}** — {s['notes']}")

    if st.button("Apply AI drafts to proposal fields + schedule"):
        st.session_state["cleaning_plan_prefill"] = analysis.get("cleaning_plan_draft", "")
        st.session_state["notes_prefill"] = analysis.get("scope_of_work_draft", "")

        if not st.session_state.get("cover_letter_body_prefill"):
            st.session_state["cover_letter_body_prefill"] = (
                "Thank you for the opportunity to submit our proposal. "
                "Please find our Cleaning Service Agreement, scope of work, and cleaning schedule enclosed. "
                "We look forward to supporting your facility with consistent, high-quality service."
            )

        ai_rows = []
        for r in analysis.get("schedule_rows", []):
            ai_rows.append(
                (r.get("task", ""), bool(r.get("daily", False)), bool(r.get("weekly", False)), bool(r.get("monthly", False)))
            )

        if ai_rows:
            st.session_state["schedule_df"] = pd.DataFrame(ai_rows, columns=["Task", "Daily", "Weekly", "Monthly"])

        st.success("Applied. Scroll up to see updated fields and schedule.")
        st.rerun()

# =========================
# Preview + downloads
# =========================
st.divider()
st.subheader("Preview")
agreement_text = build_agreement_text(p)
st.text_area("Agreement text (preview)", agreement_text, height=520)

colA, colB, colC = st.columns(3)

with colA:
    st.download_button(
        "Download .txt",
        data=agreement_text.encode("utf-8"),
        file_name=f"TorusGroup_Cleaning_Service_Agreement_{datetime.date.today().isoformat()}.txt",
        mime="text/plain",
    )

with colB:
    docx_data = docx_from_agreement(
        agreement_text,
        tuned_schedule_rows,
        clean_list(st.session_state.service_addresses),
        p
    )
    st.download_button(
        "Download .docx (Word)",
        data=docx_data,
        file_name=f"TorusGroup_Cleaning_Service_Agreement_{datetime.date.today().isoformat()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

with colC:
    st.download_button(
        "Download inputs (.json)",
        data=json.dumps(asdict(p), indent=2).encode("utf-8"),
        file_name=f"TorusGroup_Agreement_Inputs_{datetime.date.today().isoformat()}.json",
        mime="application/json",
    )
