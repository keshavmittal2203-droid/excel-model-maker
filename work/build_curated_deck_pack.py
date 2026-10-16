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

# Read existing proposal slots from proposals/err4.deck-pack.proposal.json
existing_pack = read_json(str(ROOT / "proposals" / "err4.deck-pack.proposal.json"))
existing_inputs = read_json(str(ROOT / "proposals" / "err4.inputs.proposal.json"))

# Filter existing slots to keep only those on retained slides: 1, 2, 14, 16, 23
retained_slides = [1, 2, 14, 16, 23]
kept_slots = [s for s in existing_pack["slots"] if s["slide"] in retained_slides]
kept_values = {k: v for k, v in existing_inputs["values"].items() if any(s["name"] == k for s in kept_slots)}

# Update existing workbook_binding hashes to match current recalculated workbook
for slot in kept_slots:
    if "workbook_binding" in slot:
        slot["workbook_binding"]["workbook_sha256"] = wb_sha

# New Slots to add:
new_slots = [
    # Slide 1 additions
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
        "name": "slide1_logo_suppress",
        "slide": 1,
        "target_type": "suppress_shape",
        "shape_id": 59,
        "shape_name": "Picture 58",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 1",
    },
    # Slide 2 additions
    {
        "name": "slide2_chart_revenue",
        "slide": 2,
        "target_type": "chart_data",
        "shape_id": 13,
        "shape_name": "Chart 0",
        "categories": ["FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Revenue (₹ Cr)",
        "description": "Historical 5-year Revenue progression from Data Sheet row 17 / row 28",
    },
    {
        "name": "slide2_chart_snapshot",
        "slide": 2,
        "target_type": "chart_data",
        "shape_id": 16,
        "shape_name": "Chart 1",
        "categories": ["Revenue", "EBITDA", "Net Profit", "CFO"],
        "series_name": "FY25 Actual (₹ Cr)",
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
    # Slide 14 additions
    {
        "name": "slide14_logo_suppress",
        "slide": 14,
        "target_type": "suppress_shape",
        "shape_id": 59,
        "shape_name": "Picture 58",
        "reason": "Omit unverified Sun Pharma corporate logo pending verified Marksans vector asset",
        "description": "Suppress unverified corporate logo on Slide 14",
    },
    # Slide 16 additions
    {
        "name": "slide16_chart_cr",
        "slide": 16,
        "target_type": "chart_data",
        "shape_id": 14,
        "shape_name": "Chart 0",
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Current Ratio",
        "description": "Historical Current Ratio FY18-FY25 from Marksans Data Sheet",
    },
    {
        "name": "slide16_chart_qr",
        "slide": 16,
        "target_type": "chart_data",
        "shape_id": 17,
        "shape_name": "Chart 0",
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Quick Ratio",
        "description": "Historical Quick Ratio FY18-FY25 from Marksans Data Sheet",
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
    # Slide 23 additions
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

# New Input values:
new_values = {
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
    "slide1_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
    "slide2_chart_revenue": {
        "value": [1376.18, 1490.84, 1852.14, 2177.41, 2622.85],
        "categories": ["FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Revenue (₹ Cr)",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet J28 historical series]",
    },
    "slide2_chart_snapshot": {
        "value": [2622.85, 569.21, 380.58, 206.66],
        "categories": ["Revenue", "EBITDA", "Net Profit", "CFO"],
        "series_name": "FY25 Actual (₹ Cr)",
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
        "value": [3.06, 3.15, 2.45, 2.89, 3.62, 5.02, 4.69, 4.64],
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Current Ratio",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet Current Ratio historical series]",
    },
    "slide16_chart_qr": {
        "value": [1.53, 1.32, 1.43, 1.57, 2.30, 3.51, 3.03, 2.76],
        "categories": ["FY18", "FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25"],
        "series_name": "Quick Ratio",
        "source": "candidates/Marksans-Recalculated.xlsx [Data Sheet Quick Ratio historical series]",
    },
    "slide16_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
    "slide23_table_ke": {
        "value": str(ROOT / "work" / "assets" / "image22_marksans.png"),
        "source": "candidates/Marksans-Recalculated.xlsx [WACC-DCF J22:J25; Rf=7.30%, Rm=16.4%, Beta=0.547, Ke=12.26%]",
    },
    "slide23_table_kd": {
        "value": str(ROOT / "work" / "assets" / "image23_marksans.png"),
        "source": "candidates/Marksans-Recalculated.xlsx [WACC-DCF E22:E24; PreTax=8.00%, Tax=30%, Kd=5.60%]",
    },
    "slide23_table_wacc": {
        "value": str(ROOT / "work" / "assets" / "image24_marksans.png"),
        "source": "candidates/Marksans-Recalculated.xlsx [WACC-DCF C30:J39; Mcap=14,003 Cr, Debt=343 Cr, Wd=2.24%, We=97.76%, WACC=12.11%]",
    },
    "slide23_logo_suppress": {
        "value": "suppress",
        "source": "Repository Quality Policy: omit unverified corporate logo pending authentic asset",
    },
}

all_slots = kept_slots + new_slots
all_values = {**kept_values, **new_values}

# Verify 1:1 mapping between slot names and value keys
slot_names = [s["name"] for s in all_slots]
assert len(slot_names) == len(set(slot_names)), "Duplicate slot name"
assert set(slot_names) == set(all_values), f"Mismatch: missing {set(slot_names) - set(all_values)}, extra {set(all_values) - set(slot_names)}"

final_pack = {
    "schema_version": 1,
    "template_sha256": template_sha,
    "status": "proposal",
    "approval": None,
    "proposal_notes": "Curated 5-slide Marksans Pharma research package derived from originals/ERR-4.pptx (Slides 1, 2, 14, 16, 23). Includes complete profile, executive summary, 8-year historical cash flow schedule, balance sheet ratios with native charts, DCF valuation with high-resolution WACC tables, and unverified logo suppression.",
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
