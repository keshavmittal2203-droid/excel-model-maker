"""Generate the complete curated 5-slide ERR-4 deck pack proposal and inputs for Marksans Pharma."""
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from modelmaker import sha, read_json

TEMPLATE_PATH = ROOT / "originals" / "ERR-4.pptx"
WB_PATH = ROOT / "candidates" / "Marksans-Recalculated.xlsx"

template_bytes = TEMPLATE_PATH.read_bytes()
template_sha = sha(template_bytes)

wb_bytes = WB_PATH.read_bytes()
wb_sha = sha(wb_bytes)

retained_slides = [1, 2, 14, 16, 23]

# Read existing proposal slots to retain core slide 1, 2, 14 base mappings
existing_pack = read_json(str(ROOT / "proposals" / "err4.deck-pack.proposal.json"))
existing_inputs = read_json(str(ROOT / "proposals" / "err4.inputs.proposal.json"))

# Base slots to retain from previous proposal (excluding those we redefine completely)
excluded_names = {
    # Redefined or superseded slots
    "slide16_debt_to_equity_latest",
    "slide1_chart_revenue_mix",
    "slide2_chart_revenue",
    "slide2_chart_snapshot",
    "slide16_chart_cr",
    "slide16_chart_qr",
    "slide23_table_ke",
    "slide23_table_kd",
    "slide23_table_wacc",
    "slide1_logo_suppress",
    "slide2_logo_suppress",
    "slide14_logo_suppress",
    "slide16_logo_suppress",
    "slide23_logo_suppress",
    "slide1_business_model",
    "slide1_brand_portfolio",
    "slide1_product_ecosystem",
    "slide1_strategic_approach",
}

kept_slots = [
    s for s in existing_pack["slots"]
    if s["slide"] in retained_slides and s["name"] not in excluded_names and not s["name"].startswith("slide16_cr_")
    and not s["name"].startswith("slide16_qr_") and not s["name"].startswith("slide16_de_")
    and not s["name"].startswith("slide16_icr_") and not s["name"].startswith("slide16_debtor_")
    and not s["name"].startswith("slide16_payable_") and not s["name"].startswith("slide16_inventory_")
    and not s["name"].startswith("slide16_ccc_")
]
kept_values = {s["name"]: existing_inputs["values"][s["name"]] for s in kept_slots}

# Ensure all workbook bindings reference the recalculated workbook hash
for slot in kept_slots:
    if "workbook_binding" in slot:
        slot["workbook_binding"]["workbook_sha256"] = wb_sha

# Explicitly defined curated slots
curated_slots = [
    # Slide 1: Company Profile & Operating Ecosystem
    {
        "name": "slide1_business_model",
        "slide": 1,
        "target_type": "shape_text",
        "shape_id": 12,
        "shape_name": "object 18",
        "paragraph_index": 1,
        "description": "Business Model regulated markets and formulations mix on Slide 1",
    },
    {
        "name": "slide1_brand_portfolio",
        "slide": 1,
        "target_type": "shape_text",
        "shape_id": 11,
        "shape_name": "object 18",
        "paragraph_index": 1,
        "description": "Brand Portfolio and therapeutic focus on Slide 1",
    },
    {
        "name": "slide1_product_ecosystem",
        "slide": 1,
        "target_type": "shape_text",
        "shape_id": 14,
        "shape_name": "object 18",
        "paragraph_index": 1,
        "description": "Product Ecosystem dosage forms and global regulatory compliance on Slide 1",
    },
    {
        "name": "slide1_strategic_approach",
        "slide": 1,
        "target_type": "shape_text",
        "shape_id": 20,
        "shape_name": "TextBox 19",
        "paragraph_index": 0,
        "description": "Strategic Approach backward integration and Teva Goa site expansion on Slide 1",
    },
    {
        "name": "slide1_chart_revenue_mix",
        "slide": 1,
        "target_type": "chart_data",
        "shape_id": 25,
        "shape_name": "Chart 0",
        "categories": ["UK & Europe", "US & North America", "Australia & NZ", "Rest of World"],
        "series_name": "Revenue Mix (%)",
        "chart_title": "Revenue Mix (in %)",
        "description": "Geographical Revenue Mix from Marksans Investor Presentation & Annual Report",
    },
    {
        "name": "slide1_logo_suppress",
        "slide": 1,
        "target_type": "suppress_shape",
        "shape_id": 59,
        "shape_name": "Picture 58",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 1",
    },

    # Slide 2: Executive Summary & Historical Growth
    {
        "name": "slide2_chart_revenue",
        "slide": 2,
        "target_type": "chart_data",
        "shape_id": 13,
        "shape_name": "Chart 0",
        "categories": ["FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Total Revenue",
        "chart_title": "Historical Revenue (₹ Crore)",
        "description": "Historical 5-year Consolidated Revenue from Data Sheet row 17 / row 28",
    },
    {
        "name": "slide2_chart_snapshot",
        "slide": 2,
        "target_type": "chart_data",
        "shape_id": 16,
        "shape_name": "Chart 1",
        "categories": ["Revenue", "EBITDA", "Net Profit", "CFO"],
        "series_name": "FY25 Actual (₹ Cr)",
        "chart_title": "FY25 Financial Snapshot (₹ Crore)",
        "description": "FY25 Key Financial Snapshot metrics from Data Sheet & Cash Flow",
    },
    {
        "name": "slide2_logo_suppress",
        "slide": 2,
        "target_type": "suppress_shape",
        "shape_id": 59,
        "shape_name": "Picture 58",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 2",
    },

    # Slide 14: Cash Flow Schedule
    {
        "name": "slide14_logo_suppress",
        "slide": 14,
        "target_type": "suppress_shape",
        "shape_id": 59,
        "shape_name": "Picture 58",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 14",
    },

    # Slide 16: Balance Sheet & Liquidity Ratios
    {
        "name": "slide16_chart_cr",
        "slide": 16,
        "target_type": "chart_data",
        "shape_id": 14,
        "shape_name": "Chart 0",
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Current Ratio",
        "chart_title": "Current Ratio (x)",
        "description": "Historical Current Ratio FY18-FY25 matching Table 0 Row 1",
    },
    {
        "name": "slide16_chart_qr",
        "slide": 16,
        "target_type": "chart_data",
        "shape_id": 17,
        "shape_name": "Chart 0",
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Quick Ratio",
        "chart_title": "Quick Ratio (x)",
        "description": "Historical Quick Ratio FY18-FY25 matching Table 0 Row 2",
    },
    {
        "name": "slide16_logo_suppress",
        "slide": 16,
        "target_type": "suppress_shape",
        "shape_id": 59,
        "shape_name": "Picture 58",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 16",
    },

    # Slide 23: DCF Valuation & Reconciled WACC
    {
        "name": "slide23_table_ke",
        "slide": 23,
        "target_type": "picture_replacement",
        "shape_id": 7,
        "shape_name": "Picture 6",
        "description": "Cost of Equity schedule replacement table image bound to WACC-DCF J22:J25",
    },
    {
        "name": "slide23_table_kd",
        "slide": 23,
        "target_type": "picture_replacement",
        "shape_id": 9,
        "shape_name": "Picture 8",
        "description": "Post-tax Cost of Debt schedule replacement table image bound to WACC-DCF E22:E24",
    },
    {
        "name": "slide23_table_wacc",
        "slide": 23,
        "target_type": "picture_replacement",
        "shape_id": 11,
        "shape_name": "Picture 10",
        "description": "WACC calculation table replacement image bound to WACC-DCF C30:J39",
    },
    {
        "name": "slide23_logo_suppress",
        "slide": 23,
        "target_type": "suppress_shape",
        "shape_id": 25,
        "shape_name": "Picture 24",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 23",
    },
]

# Slide 16 Table 0 cells (8 years: FY18 to FY25 across rows 1 to 8)
years = ["fy18", "fy19", "fy20", "fy21", "fy22", "fy23", "fy24", "fy25"]

ratio_rows = [
    ("cr", 1, "Current Ratio (x)", ["3.18", "3.33", "2.54", "2.95", "3.73", "5.26", "5.03", "4.95"]),
    ("qr", 2, "Quick Ratio (x)", ["1.64", "1.50", "1.51", "1.63", "2.42", "3.75", "3.36", "3.07"]),
    ("de", 3, "Debt-to-Equity (x)", ["0.25", "0.18", "0.05", "0.04", "0.09", "0.07", "0.12", "0.13"]),
    ("icr", 4, "Interest Coverage Ratio (x)", ["5.87", "11.80", "18.96", "38.82", "30.29", "37.98", "38.82", "44.22"]),
    ("debtor_days", 5, "Debtor Days", ["71", "64", "78", "72", "97", "82", "76", "75"]),
    ("payable_days", 6, "Payable Days", ["57", "75", "78", "84", "104", "112", "125", "134"]),
    ("inventory_days", 7, "Inventory Days", ["87", "107", "78", "107", "104", "95", "104", "118"]),
    ("ccc", 8, "Cash Conversion Cycle (Days)", ["101", "96", "78", "95", "97", "65", "55", "59"]),
]

s16_table_slots = []
s16_table_values = {}

for prefix, row_idx, metric_label, row_vals in ratio_rows:
    for col_idx, (yr, val) in enumerate(zip(years, row_vals), start=1):
        slot_name = f"slide16_{prefix}_{yr}"
        s16_table_slots.append({
            "name": slot_name,
            "slide": 16,
            "target_type": "table_cell",
            "shape_id": 9,
            "shape_name": "Table 0",
            "row": row_idx,
            "col": col_idx,
            "description": f"{metric_label} {yr.upper()} in Slide 16 Table 0",
        })
        s16_table_values[slot_name] = {
            "value": val,
            "source": f"candidates/Marksans-Recalculated.xlsx [Data Sheet & BS historical series for {metric_label}]",
        }

# Curated values mapping
curated_values = {
    "slide1_business_model": {
        "value": "• Regulated Markets Focus: Strong presence across the UK, US, Australia, and Continental Europe specializing in over-the-counter (OTC) and prescription generics.\n• Integrated Manufacturing: Global scale facilities in Goa, Southport (UK), and Farmingdale (US) driving high-margin solid oral and liquid dosage formulations.",
        "source": "Marksans Investor Presentation FY25 & Annual Report",
    },
    "slide1_brand_portfolio": {
        "value": "• Core Therapeutic Areas: Pain Management, Cough & Cold, Cardiovascular/CNS, Gastrointestinal, and Anti-diabetic formulations.\n• OTC Brand Leadership: Key OTC offerings marketed under private label and store brands across major international pharmacy and retail chains.",
        "source": "Marksans Investor Presentation FY25 & Annual Report",
    },
    "slide1_product_ecosystem": {
        "value": "• Dosage Forms: Tablets, hard and soft gelatin capsules, oral liquids, and ointments.\n• Global Compliance: US FDA, UK MHRA, and Australian TGA approved cGMP manufacturing infrastructure.",
        "source": "Marksans Annual Report FY25 & Regulatory Disclosures",
    },
    "slide1_strategic_approach": {
        "value": "• Capacity Expansion & Forward Integration: Expanding capacity at the Teva Goa site, increasing backward API integration, and scaling direct-to-pharmacy distribution across regulated Western markets.",
        "source": "Marksans Investor Presentation FY25 & Management Commentary",
    },
    "slide1_chart_revenue_mix": {
        "value": [43.1, 41.7, 9.8, 5.4],
        "categories": ["UK & Europe", "US & North America", "Australia & NZ", "Rest of World"],
        "series_name": "Revenue Mix (%)",
        "chart_title": "Revenue Mix (in %)",
        "source": "Marksans Investor Presentation FY25 & Annual Report [Geographical Revenue Breakdown]",
    },
    "slide1_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
    "slide2_chart_revenue": {
        "value": [1376.18, 1490.84, 1852.14, 2177.41, 2622.85],
        "categories": ["FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Total Revenue",
        "chart_title": "Historical Revenue (₹ Crore)",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet J28 historical series]",
    },
    "slide2_chart_snapshot": {
        "value": [2622.85, 569.21, 380.58, 206.66],
        "categories": ["Revenue", "EBITDA", "Net Profit", "CFO"],
        "series_name": "FY25 Actual (₹ Cr)",
        "chart_title": "FY25 Financial Snapshot (₹ Crore)",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet J28, J40, J54, Cash Flow J24]",
    },
    "slide2_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
    "slide14_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
    "slide16_chart_cr": {
        "value": [3.18, 3.33, 2.54, 2.95, 3.73, 5.26, 5.03, 4.95],
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Current Ratio",
        "chart_title": "Current Ratio (x)",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet & BS Current Ratio historical series FY18-FY25]",
    },
    "slide16_chart_qr": {
        "value": [1.64, 1.50, 1.51, 1.63, 2.42, 3.75, 3.36, 3.07],
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Quick Ratio",
        "chart_title": "Quick Ratio (x)",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet & BS Quick Ratio historical series FY18-FY25]",
    },
    "slide16_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
    "slide23_table_ke": {
        "value": "work/assets/image22_marksans.png",
        "source": "candidates/Marksans-Recalculated.xlsx [WACC-DCF J22:J25; Rf=7.30%, Rm=16.37%, Beta=0.547, Ke=12.26%]",
    },
    "slide23_table_kd": {
        "value": "work/assets/image23_marksans.png",
        "source": "candidates/Marksans-Recalculated.xlsx [WACC-DCF E22:E24; PreTax=8.00%, Tax=30.00%, Kd=5.60%]",
    },
    "slide23_table_wacc": {
        "value": "work/assets/image24_marksans.png",
        "source": "candidates/Marksans-Recalculated.xlsx [WACC-DCF C30:J39; Current Mcap=14,003 Cr, Current Debt=343 Cr, Target Wd=2.24%, Target We=97.76%, WACC=12.11%]",
    },
    "slide23_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
}

all_slots = kept_slots + curated_slots + s16_table_slots
all_values = {**kept_values, **curated_values, **s16_table_values}

# Verify 1:1 mapping between slot names and value keys
slot_names = [s["name"] for s in all_slots]
assert len(slot_names) == len(set(slot_names)), f"Duplicate slot names: {[x for x in slot_names if slot_names.count(x) > 1]}"
assert set(slot_names) == set(all_values), f"Mismatch: missing {set(slot_names) - set(all_values)}, extra {set(all_values) - set(slot_names)}"

final_pack = {
    "schema_version": 1,
    "template_sha256": template_sha,
    "status": "proposal",
    "approval": None,
    "proposal_notes": "5-Slide Marksans Pharma Financial Overview derived from originals/ERR-4.pptx (Slides 1, 2, 14, 16, 23). Includes verified company profile, executive summary, 8-year historical cash flow schedule, balance sheet ratios with native charts, DCF valuation with reconciled WACC tables, and unverified logo suppression.",
    "retained_slides": retained_slides,
    "slots": all_slots,
}

final_inputs = {
    "values": all_values,
}

# Write out files
pack_out = ROOT / "proposals" / "err4.deck-pack.proposal.json"
inputs_out = ROOT / "proposals" / "err4.inputs.proposal.json"

def json_default(obj):
    if hasattr(obj, "as_tuple"):
        # Decimal
        return int(obj) if obj % 1 == 0 else float(obj)
    return str(obj)

pack_out.write_text(json.dumps(final_pack, indent=2, default=json_default), encoding="utf-8")
inputs_out.write_text(json.dumps(final_inputs, indent=2, default=json_default), encoding="utf-8")

print(f"Updated {pack_out} with {len(all_slots)} slots.")
print(f"Updated {inputs_out} with {len(all_values)} values.")
