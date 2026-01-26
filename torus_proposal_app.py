import datetime
import os
from dataclasses import dataclass, asdict
from io import BytesIO
from typing import List, Dict, Optional

import streamlit as st
from docx import Document


COMPANY_NAME = "Torus Group"


# ---------------- Data ----------------

@dataclass
class ProposalInputs:
    # Core
    client: str
    facility_name: str
    space_type: str
    square_footage: int
    floor_types: str

    # Service details
    service_begin_date: str
    service_end_date: str
    service_addresses: List[str]
    days_per_week: int
    cleaning_times: str

    # Rooms
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
    supplies_included: str  # Yes/No

    # Pricing
    pricing_mode: str  # Monthly Fixed | Per Sq Ft | Per Visit
    monthly_fixed_price: float
    rate_per_sqft: float

    rate_per_visit: float
    visits_per_week: float
    visits_per_month: int  # computed/displayed

    # Deep clean
    deep_clean_option: str  # None | One-time | Quarterly
    deep_clean_price: float
    deep_clean_includes: List[str]

    # Add-ons
    additional_services: List[Dict[str, float]]  # [{"name": str, "price": float}]
    include_addons_in_total: str  # Yes/No

    # Tax / payment
    sales_tax_percent: float
    net_terms: int  # 15/30/45/60

    # Compensation
    compensation_mode: str  # Auto (calculated) | Override
    compensation_override: float

    # Notes
    notes: str


# ---------------- Helpers ----------------

def money(x: float) -> str:
    return f"${x:,.2f}"


def compute_visits_per_month(visits_per_week: float) -> int:
    # 52 weeks / 12 months ≈ 4.3333
    return int(round(visits_per_week * (52.0 / 12.0)))


def clean_address_list(items: List[str]) -> List[str]:
    out = []
    for s in items:
        s2 = (s or "").strip()
        if s2:
            out.append(s2)
    return out


def build_totals(p: ProposalInputs) -> dict:
    # Base monthly
    if p.pricing_mode == "Monthly Fixed":
        base_monthly = float(p.monthly_fixed_price)
        base_explain = f"Monthly fixed price: {money(base_monthly)}"

    elif p.pricing_mode == "Per Sq Ft":
        base_monthly = float(p.rate_per_sqft) * float(p.square_footage)
        base_explain = f"Rate: {money(p.rate_per_sqft)}/sqft × {p.square_footage:,} sqft = {money(base_monthly)} per month"

    else:  # Per Visit
        base_monthly = float(p.rate_per_visit) * float(p.visits_per_month)
        base_explain = (
            f"Rate: {money(p.rate_per_visit)}/visit × {p.visits_per_month} visits/month "
            f"({p.visits_per_week:g}/week) = {money(base_monthly)} per month"
        )

    # Add-ons
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

    # Deep clean handling
    deep_clean_one_time = 0.0
    deep_clean_quarterly = 0.0
    deep_clean_monthly_equiv = 0.0

    if p.deep_clean_option == "One-time":
        deep_clean_one_time = float(p.deep_clean_price)
    elif p.deep_clean_option == "Quarterly":
        deep_clean_quarterly = float(p.deep_clean_price)
        deep_clean_monthly_equiv = deep_clean_quarterly / 3.0

    # Subtotal (monthly)
    monthly_subtotal = base_monthly + addons_included_monthly + deep_clean_monthly_equiv

    # Sales tax (monthly)
    tax_rate = max(0.0, float(p.sales_tax_percent)) / 100.0
    monthly_tax = monthly_subtotal * tax_rate
    monthly_total_with_tax = monthly_subtotal + monthly_tax

    # Compensation display (monthly)
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


def build_proposal_text(p: ProposalInputs) -> str:
    today = datetime.date.today().strftime("%B %d, %Y")
    totals = build_totals(p)

    addresses = clean_address_list(p.service_addresses)
    address_block = "\n".join([f"• {a}" for a in addresses]) if addresses else "• (not provided)"

    floor_note = f"Floor types/notes: {p.floor_types}" if p.floor_types.strip() else "Floor types/notes: N/A"

    scope_lines = [
        "• Empty trash/recycling; replace liners as needed",
        "• Dust and wipe accessible surfaces (desks, ledges, counters)",
        "• Clean and disinfect high-touch points (door handles, switches, etc.)",
        "• Vacuum carpeted areas; spot clean as needed",
        "• Damp mop hard floors; detail as required by floor type",
        "• Clean and disinfect bathrooms; refill soap/paper as applicable",
        "• Break rooms/kitchens: wipe counters, clean sinks, exterior of appliances",
    ]
    if p.num_conference_rooms > 0:
        scope_lines.append("• Conference rooms: wipe tables, straighten, vacuum/mop")
    if p.num_locker_rooms > 0:
        scope_lines.append("• Locker rooms: clean/disinfect, mop floors, touchpoint disinfection")
    if p.day_porter_needed == "Yes":
        scope_lines.append("• Day porter support (restroom checks, spills, touchpoints, common areas)")
    if p.restocking_needed == "Yes":
        scope_lines.append("• Restocking of consumables (client-provided unless supplies included)")

    # Pricing summary
    pricing_lines = [f"• Base service: {totals['base_explain']}"]

    if totals["include_addons"]:
        pricing_lines.append(f"• Additional services (included): {money(totals['addons_total'])} per month")
    else:
        pricing_lines.append(
            f"• Additional services (not included in total): {money(totals['addons_total'])}"
            if totals["addons_total"] > 0 else
            "• Additional services: None"
        )

    if p.deep_clean_option == "One-time":
        pricing_lines.append(f"• One-time deep clean: {money(totals['deep_clean_one_time'])} (one-time)")
    elif p.deep_clean_option == "Quarterly":
        pricing_lines.append(f"• Quarterly deep clean: {money(totals['deep_clean_quarterly'])} per quarter")
        pricing_lines.append(f"• Quarterly deep clean monthly equivalent: {money(totals['deep_clean_monthly_equiv'])}/month")

    pricing_lines.append(f"• Monthly subtotal (pre-tax): {money(totals['monthly_subtotal'])}")
    pricing_lines.append(f"• Sales tax ({p.sales_tax_percent:.2f}%): {money(totals['monthly_tax'])}")
    pricing_lines.append(f"• Monthly total (with tax): {money(totals['monthly_total_with_tax'])}")
    pricing_lines.append(f"• {totals['compensation_explain']}")

    deep_clean_block = ""
    if p.deep_clean_option != "None" and p.deep_clean_includes:
        deep_clean_block = "\nDEEP CLEAN INCLUDES\n" + "\n".join([f"• {x}" for x in p.deep_clean_includes]) + "\n"

    addon_detail_block = ""
    if totals["addons_lines"]:
        addon_detail_block = "\nADDITIONAL SERVICES (LINE ITEMS)\n" + "\n".join(totals["addons_lines"]) + "\n"

    return f"""JANITORIAL SERVICES PROPOSAL
{COMPANY_NAME}
Date: {today}

CLIENT / SERVICE INFO
Client: {p.client}
Facility/Location Name: {p.facility_name}
Service begin date: {p.service_begin_date}
Service end date: {p.service_end_date}
Number of days per week: {p.days_per_week}
Cleaning times: {p.cleaning_times}
Net pay terms: Net {p.net_terms}

SERVICE ADDRESSES
{address_block}

FACILITY OVERVIEW
• Space type: {p.space_type}
• Approx. square footage: {p.square_footage:,} sqft
• Offices: {p.num_offices}
• Conference rooms: {p.num_conference_rooms}
• Break rooms: {p.num_break_rooms}
• Bathrooms: {p.num_bathrooms}
• Kitchens: {p.num_kitchens}
• Locker rooms: {p.num_locker_rooms}
• {floor_note}

SCOPE OF WORK (SUMMARY)
{chr(10).join(scope_lines)}

PRICING SUMMARY
{chr(10).join(pricing_lines)}
{deep_clean_block}{addon_detail_block}
NOTES
{p.notes if p.notes.strip() else "(none)"}

ACCEPTANCE
Authorized Signature: ___________________________    Date: _______________
"""


def docx_bytes_from_text(text: str) -> bytes:
    """
    Uses proposal_template.docx if it exists (keeps your footer/header exactly).
    If not present, generates a normal doc.
    """
    template_path = "proposal_template.docx"
    if os.path.exists(template_path):
        doc = Document(template_path)
    else:
        doc = Document()

    for line in text.splitlines():
        if line.strip() and line == line.upper() and len(line) <= 90:
            para = doc.add_paragraph()
            run = para.add_run(line)
            run.bold = True
        else:
            doc.add_paragraph(line)

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ---------------- UI ----------------

st.set_page_config(page_title=f"{COMPANY_NAME} Proposal Builder", layout="wide")
st.title(f"{COMPANY_NAME} — Proposal Builder")

# Sidebar
with st.sidebar:
    st.header("Client / Contract")
    client = st.text_input("Client (legal name)")
    facility_name = st.text_input("Facility/Location name")

    service_begin_date = st.text_input("Service begin date")
    service_end_date = st.text_input("Service end date")

    net_terms = st.selectbox("Net pay terms", [15, 30, 45, 60], index=1)

    days_per_week = st.number_input("Number of days per week", min_value=0, step=1, value=5)
    cleaning_times = st.text_input("Cleaning times (e.g., 6:00 PM – 10:00 PM)")

    st.header("Sales Tax")
    sales_tax_percent = st.number_input("Sales tax (%)", min_value=0.0, step=0.25, value=0.0)

    st.header("Service (ops)")
    cleaning_frequency = st.text_input("Cleaning frequency (e.g., 5x/week)", value="5x/week")
    trash_pickup = st.text_input("Trash pickup schedule", value="daily")
    day_porter_needed = st.selectbox("Day porter needed", ["No", "Yes"])
    restocking_needed = st.selectbox("Restocking needed", ["No", "Yes"])
    supplies_included = st.selectbox("Supplies included", ["No", "Yes"])


# Main columns
c1, c2, c3 = st.columns(3)

with c1:
    st.subheader("Facility details")
    space_type = st.text_input("Type of space (Office/Medical/etc.)", value="Office")
    square_footage = st.number_input("Square footage", min_value=0, step=100, value=0)
    floor_types = st.text_area("Floor types (optional)", placeholder="Carpet 7,600 sqft; VCT 30,000 sqft; Epoxy 50,000 sqft")

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
        rate_per_sqft = st.number_input("Rate per sq ft ($/sqft)", min_value=0.0, step=0.001, format="%.4f", value=0.0000)
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

    st.subheader("Notes")
    notes = st.text_area("Notes (optional)", height=120)


st.divider()

# Multi-address section
st.subheader("Service Address(es)")
st.caption("Add one or more addresses for this proposal.")

if "service_addresses" not in st.session_state:
    st.session_state.service_addresses = [""]  # start with one line

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

# Additional services line items
st.subheader("Additional services (add-ons)")
st.caption("Add line items (example: Day porter hours, Event cleanup, Carpet extraction add-on).")

if "addons" not in st.session_state:
    st.session_state.addons = [{"name": "", "price": 0.0}]

for i, item in enumerate(st.session_state.addons):
    ca, cb, cc = st.columns([3, 1, 1])
    with ca:
        st.session_state.addons[i]["name"] = st.text_input(f"addon_name_{i}", value=item["name"], placeholder="Service name", label_visibility="collapsed")
    with cb:
        st.session_state.addons[i]["price"] = st.number_input(f"addon_price_{i}", min_value=0.0, step=25.0, value=float(item["price"]), label_visibility="collapsed")
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


# Compensation controls
st.divider()
st.subheader("Compensation")
compensation_mode = st.selectbox("Compensation mode", ["Auto (calculated)", "Override"])
compensation_override = 0.0
if compensation_mode == "Override":
    compensation_override = st.number_input("Compensation override ($ per month)", min_value=0.0, step=50.0, value=0.0)
    st.caption("This is what will be shown as Compensation in the proposal, regardless of totals.")


# Build proposal inputs
p = ProposalInputs(
    client=client.strip(),
    facility_name=facility_name.strip(),
    space_type=space_type.strip(),
    square_footage=int(square_footage),
    floor_types=floor_types.strip(),

    service_begin_date=service_begin_date.strip(),
    service_end_date=service_end_date.strip(),
    service_addresses=st.session_state.service_addresses,
    days_per_week=int(days_per_week),
    cleaning_times=cleaning_times.strip(),

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
    supplies_included=supplies_included,

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

    sales_tax_percent=float(sales_tax_percent),
    net_terms=int(net_terms),

    compensation_mode=compensation_mode,
    compensation_override=float(compensation_override),

    notes=notes.strip(),
)

totals = build_totals(p)

# Totals display
st.subheader("Calculated totals")
t1, t2, t3, t4 = st.columns(4)
t1.metric("Monthly subtotal (pre-tax)", money(totals["monthly_subtotal"]))
t2.metric("Sales tax (monthly)", money(totals["monthly_tax"]))
t3.metric("Monthly total (with tax)", money(totals["monthly_total_with_tax"]))
t4.metric("Compensation (monthly)", money(totals["compensation_monthly"]))

if p.deep_clean_option == "One-time":
    st.info(f"One-time deep clean (separate): {money(totals['deep_clean_one_time'])}")
elif p.deep_clean_option == "Quarterly":
    st.info(f"Quarterly deep clean: {money(totals['deep_clean_quarterly'])} per quarter (monthly equivalent {money(totals['deep_clean_monthly_equiv'])})")

# Preview + downloads
st.divider()
st.subheader("Preview")
proposal_text = build_proposal_text(p)
st.text_area("Generated proposal text", proposal_text, height=520)

colA, colB, colC = st.columns(3)
with colA:
    st.download_button(
        "Download .txt",
        data=proposal_text.encode("utf-8"),
        file_name=f"TorusGroup_Proposal_{datetime.date.today().isoformat()}.txt",
        mime="text/plain",
    )
with colB:
    docx_data = docx_bytes_from_text(proposal_text)
    st.download_button(
        "Download .docx (Word)",
        data=docx_data,
        file_name=f"TorusGroup_Proposal_{datetime.date.today().isoformat()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
with colC:
    st.download_button(
        "Download inputs (.json)",
        data=str(asdict(p)).encode("utf-8"),
        file_name=f"TorusGroup_Inputs_{datetime.date.today().isoformat()}.json",
        mime="application/json",
    )
