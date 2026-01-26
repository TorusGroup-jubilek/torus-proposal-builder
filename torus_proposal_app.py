import datetime
from dataclasses import dataclass, asdict
from io import BytesIO
from typing import List, Dict

import streamlit as st
from docx import Document


COMPANY_NAME = "Torus Group"


@dataclass
class ProposalInputs:
    client_name: str
    facility_name: str
    facility_address: str
    space_type: str
    square_footage: int
    num_offices: int
    num_conference_rooms: int
    num_break_rooms: int
    num_bathrooms: int
    num_kitchens: int
    num_locker_rooms: int
    floor_types: str
    cleaning_frequency: str
    day_porter_needed: str  # Yes/No
    trash_pickup: str
    restocking_needed: str  # Yes/No
    supplies_included: str  # Yes/No
    start_date: str
    notes: str

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

    # Admin / terms
    payment_terms: str  # Net 15 / Net 30 / Due on receipt / Custom
    custom_payment_terms: str
    walkthrough_date: str  # free text or date string


def money(x: float) -> str:
    return f"${x:,.2f}"


def compute_visits_per_month(visits_per_week: float) -> int:
    # 52 weeks / 12 months ≈ 4.3333 weeks/month
    vpm = visits_per_week * (52.0 / 12.0)
    return int(round(vpm))


def build_totals(p: ProposalInputs):
    base_monthly = 0.0
    base_explain = ""

    if p.pricing_mode == "Monthly Fixed":
        base_monthly = float(p.monthly_fixed_price)
        base_explain = f"Monthly fixed price: {money(base_monthly)}"

    elif p.pricing_mode == "Per Sq Ft":
        base_monthly = float(p.rate_per_sqft) * float(p.square_footage)
        base_explain = f"Rate: {money(p.rate_per_sqft)}/sqft × {p.square_footage:,} sqft = {money(base_monthly)} per month"

    elif p.pricing_mode == "Per Visit":
        base_monthly = float(p.rate_per_visit) * float(p.visits_per_month)
        base_explain = (
            f"Rate: {money(p.rate_per_visit)}/visit × {p.visits_per_month} visits/month "
            f"({p.visits_per_week:g}/week) = {money(base_monthly)} per month"
        )
    else:
        base_explain = "Pricing: (not set)"

    addons_total = 0.0
    addons_lines = []
    for item in p.additional_services:
        name = str(item.get("name", "")).strip()
        price = float(item.get("price", 0.0) or 0.0)
        if name and price > 0:
            addons_total += price
            addons_lines.append(f"• {name}: {money(price)}")

    deep_clean_one_time = 0.0
    deep_clean_quarterly = 0.0
    deep_clean_monthly_equiv = 0.0

    if p.deep_clean_option == "One-time":
        deep_clean_one_time = float(p.deep_clean_price)
    elif p.deep_clean_option == "Quarterly":
        deep_clean_quarterly = float(p.deep_clean_price)
        deep_clean_monthly_equiv = deep_clean_quarterly / 3.0

    include_addons = (p.include_addons_in_total == "Yes")
    monthly_total = base_monthly + (addons_total if include_addons else 0.0) + deep_clean_monthly_equiv

    return {
        "base_monthly": base_monthly,
        "base_explain": base_explain,
        "addons_total": addons_total,
        "addons_lines": addons_lines,
        "include_addons": include_addons,
        "deep_clean_one_time": deep_clean_one_time,
        "deep_clean_quarterly": deep_clean_quarterly,
        "deep_clean_monthly_equiv": deep_clean_monthly_equiv,
        "monthly_total": monthly_total,
    }


def build_proposal_text(p: ProposalInputs) -> str:
    today = datetime.date.today().strftime("%B %d, %Y")
    totals = build_totals(p)

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

    floor_note = f"Floor types/notes: {p.floor_types}" if p.floor_types.strip() else "Floor types/notes: N/A"

    # Pricing summary formatting
    pricing_lines = [f"• Base service: {totals['base_explain']}"]

    if totals["include_addons"]:
        pricing_lines.append(f"• Additional services (included): {money(totals['addons_total'])} per month")
    else:
        if totals["addons_total"] > 0:
            pricing_lines.append(f"• Additional services (not included in total): {money(totals['addons_total'])}")
        else:
            pricing_lines.append("• Additional services: None")

    deep_clean_block = ""
    if p.deep_clean_option == "One-time":
        pricing_lines.append(f"• One-time deep clean: {money(totals['deep_clean_one_time'])} (one-time)")
    elif p.deep_clean_option == "Quarterly":
        pricing_lines.append(f"• Quarterly deep clean: {money(totals['deep_clean_quarterly'])} per quarter")
        pricing_lines.append(f"• Quarterly deep clean monthly equivalent: {money(totals['deep_clean_monthly_equiv'])}/month")

    pricing_lines.append(f"• Estimated monthly total: {money(totals['monthly_total'])}")

    if p.deep_clean_option != "None":
        includes = p.deep_clean_includes[:] if p.deep_clean_includes else []
        if includes:
            deep_clean_block = (
                "\nDEEP CLEAN INCLUDES\n"
                + "\n".join([f"• {x}" for x in includes])
                + "\n"
            )

    addon_detail_block = ""
    if totals["addons_lines"]:
        addon_detail_block = "\nADDITIONAL SERVICES (LINE ITEMS)\n" + "\n".join(totals["addons_lines"]) + "\n"

    # Payment terms + walkthrough
    if p.payment_terms == "Custom":
        pay_terms = p.custom_payment_terms.strip() or "Custom (to be specified)"
    else:
        pay_terms = p.payment_terms

    walkthrough_line = f"Walkthrough date: {p.walkthrough_date}" if p.walkthrough_date.strip() else "Walkthrough date: (to be scheduled)"

    terms = [
        f"Service frequency: {p.cleaning_frequency}",
        f"Trash pickup schedule: {p.trash_pickup}",
        f"Supplies included: {p.supplies_included}",
        f"Target start date: {p.start_date}",
        walkthrough_line,
        f"Payment terms: {pay_terms}",
        "Pricing assumes normal access, utilities available, and standard soil levels.",
        "Specialty services (strip & wax, carpet extraction, high dusting) can be added by request.",
    ]

    if p.notes.strip():
        terms.append(f"Additional notes: {p.notes}")

    return f"""JANITORIAL SERVICES PROPOSAL
{COMPANY_NAME}
Date: {today}

Prepared For:
{p.client_name}
Facility: {p.facility_name}
Address: {p.facility_address}

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
SERVICE TERMS / ASSUMPTIONS
{chr(10).join("• " + t for t in terms)}

ACCEPTANCE
Authorized Signature: ___________________________    Date: _______________
"""


def docx_bytes_from_text(text: str) -> bytes:
    doc = Document()
    for line in text.splitlines():
        if line.strip() and line == line.upper() and len(line) <= 80:
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
st.title(f"{COMPANY_NAME} — Janitorial Proposal Builder")

# Sidebar inputs
with st.sidebar:
    st.header("Client / Site")
    client_name = st.text_input("Client name")
    facility_name = st.text_input("Facility name")
    facility_address = st.text_area("Facility address", height=80)

    st.header("Service")
    cleaning_frequency = st.text_input("Cleaning frequency (e.g., 5x/week)", value="5x/week")
    trash_pickup = st.text_input("Trash pickup schedule", value="daily")
    day_porter_needed = st.selectbox("Day porter needed", ["No", "Yes"])
    restocking_needed = st.selectbox("Restocking needed", ["No", "Yes"])
    supplies_included = st.selectbox("Supplies included", ["No", "Yes"])
    start_date = st.text_input("Target start date", value="")

    st.header("Admin / Terms")
    payment_terms = st.selectbox("Payment terms", ["Net 15", "Net 30", "Due on receipt", "Custom"])
    custom_payment_terms = ""
    if payment_terms == "Custom":
        custom_payment_terms = st.text_input("Custom payment terms")
    walkthrough_date = st.text_input("Walkthrough date (optional)", value="")

# Main layout
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
        rate_per_sqft = st.number_input("Rate per sq ft ($/sqft)", min_value=0.0, step=0.001, format="%.4f", value=0.0000)
        st.caption("This calculates: rate × square footage (as monthly base).")

    else:
        rate_per_visit = st.number_input("Rate per visit ($/visit)", min_value=0.0, step=25.0, value=0.0)
        visits_per_week = st.number_input("Visits per week", min_value=0.0, step=0.5, value=0.0)
        visits_per_month = compute_visits_per_month(float(visits_per_week))
        st.caption(f"Estimated visits/month based on {visits_per_week:g}/week: {visits_per_month} (52 weeks / 12 months)")

    st.subheader("Deep clean")
    deep_clean_option = st.selectbox("Deep clean option", ["None", "One-time", "Quarterly"])
    deep_clean_price = 0.0

    deep_clean_includes = []
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

        if deep_clean_option == "Quarterly":
            st.caption("Quarterly deep clean is shown as quarterly + monthly equivalent in the proposal.")

    st.subheader("Notes")
    notes = st.text_area("Notes (optional)", height=120)

st.divider()

# Additional services (line items)
st.subheader("Additional services (add-ons)")
st.caption("Add line items (example: Day porter hours, Strip & wax add-on, Carpet extraction add-on, Event cleanup).")

if "addons" not in st.session_state:
    st.session_state.addons = [{"name": "", "price": 0.0}]

col_add_a, col_add_b, col_add_c = st.columns([3, 1, 1])
with col_add_a:
    st.write("Service name")
with col_add_b:
    st.write("Price ($)")
with col_add_c:
    st.write("")

for i, item in enumerate(st.session_state.addons):
    ca, cb, cc = st.columns([3, 1, 1])
    with ca:
        st.session_state.addons[i]["name"] = st.text_input(f"addon_name_{i}", value=item["name"], label_visibility="collapsed")
    with cb:
        st.session_state.addons[i]["price"] = st.number_input(
            f"addon_price_{i}",
            min_value=0.0,
            step=25.0,
            value=float(item["price"]),
            label_visibility="collapsed"
        )
    with cc:
        if st.button("Remove", key=f"remove_addon_{i}"):
            st.session_state.addons.pop(i)
            st.rerun()

col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
with col_btn1:
    if st.button("Add another line"):
        st.session_state.addons.append({"name": "", "price": 0.0})
        st.rerun()
with col_btn2:
    include_addons_in_total = st.selectbox("Include add-ons in monthly total?", ["Yes", "No"])
with col_btn3:
    st.write("")

# Build inputs object
p = ProposalInputs(
    client_name=client_name.strip(),
    facility_name=facility_name.strip(),
    facility_address=facility_address.strip(),
    space_type=space_type.strip(),
    square_footage=int(square_footage),
    num_offices=int(num_offices),
    num_conference_rooms=int(num_conference_rooms),
    num_break_rooms=int(num_break_rooms),
    num_bathrooms=int(num_bathrooms),
    num_kitchens=int(num_kitchens),
    num_locker_rooms=int(num_locker_rooms),
    floor_types=floor_types.strip(),
    cleaning_frequency=cleaning_frequency.strip(),
    day_porter_needed=day_porter_needed,
    trash_pickup=trash_pickup.strip(),
    restocking_needed=restocking_needed,
    supplies_included=supplies_included,
    start_date=start_date.strip(),
    notes=notes.strip(),
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
    payment_terms=payment_terms,
    custom_payment_terms=custom_payment_terms.strip(),
    walkthrough_date=walkthrough_date.strip(),
)

totals = build_totals(p)

# Show computed totals
st.subheader("Calculated totals")
ct1, ct2, ct3 = st.columns(3)
ct1.metric("Base monthly", money(totals["base_monthly"]))
ct2.metric("Add-ons total", money(totals["addons_total"]))
ct3.metric("Estimated monthly total", money(totals["monthly_total"]))

if p.deep_clean_option == "One-time":
    st.info(f"One-time deep clean (separate): {money(totals['deep_clean_one_time'])}")
elif p.deep_clean_option == "Quarterly":
    st.info(
        f"Quarterly deep clean: {money(totals['deep_clean_quarterly'])} per quarter "
        f"(monthly equivalent {money(totals['deep_clean_monthly_equiv'])})"
    )

st.divider()
st.subheader("Preview")

proposal_text = build_proposal_text(p)
st.text_area("Generated proposal text", proposal_text, height=480)

# Downloads
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
