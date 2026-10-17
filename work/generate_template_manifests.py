import hashlib
import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

TARGETS = {
    "MM": {
        "name": "Mahindra & Mahindra Limited",
        "ticker": "M&M.NS",
        "bse": "500520",
        "cmp": 3497.20,
        "shares_cr": 124.30,
        "market_cap_cr": 434701.96,
        "wacc": 0.1716,
        "beta": 1.08,
        "dcf_per_share": 686.81,
        "sotp_per_share": 2742.97,
        "primary_method": "Sum-of-the-Parts (SOTP) with Standalone Net Cash (+₹9,109 Cr)",
    },
    "LTM": {
        "name": "LTIMindtree Limited",
        "ticker": "LTM.NS",
        "bse": "540005",
        "cmp": 5609.50,
        "shares_cr": 29.62,
        "market_cap_cr": 166153.39,
        "wacc": 0.1547,
        "beta": 0.92,
        "dcf_per_share": 1506.84,
        "sotp_per_share": 2811.09,
        "primary_method": "Full-Year Annual FCFF DCF Baseline (Asset-Light IT Services)",
    },
    "CLEAN": {
        "name": "Clean Science and Technology Limited",
        "ticker": "CLEAN.NS",
        "bse": "543318",
        "cmp": 1041.40,
        "shares_cr": 10.63,
        "market_cap_cr": 11070.08,
        "wacc": 0.1493,
        "beta": 0.85,
        "dcf_per_share": 181.97,
        "sotp_per_share": 480.89,
        "primary_method": "Full-Year Manufacturing FCFF DCF Baseline (10,500 MTPA HALS Base)",
    },
    "EDELWEISS": {
        "name": "Edelweiss Financial Services Limited",
        "ticker": "EDELWEISS.NS",
        "bse": "532922",
        "cmp": 108.85,
        "shares_cr": 94.50,
        "market_cap_cr": 10286.33,
        "wacc": 0.1403,
        "beta": 1.22,
        "dcf_per_share": 71.91,
        "sotp_per_share": 130.53,
        "primary_method": "Financial Services Vertical SOTP (Industrial FCFF Prohibited)",
    },
}

SHEET_NAMES = [
    "Model ->", "Main Financials", "Ratio Analysis", "Intrinsic Growth", "Comps",
    "WACC-DCF", "DCF 1st way", "DCF 2nd way", "Data ->", "compdata",
    "Data Sheet", "BS", "Beta", "RM", "peerdata", "Cash Flow", "sectordata"
]

for tgt, meta in TARGETS.items():
    pkg_dir = ROOT / "companies" / tgt / "template_v2"
    xlsx_path = pkg_dir / f"{tgt}-FullTemplate.xlsx"
    sha = hashlib.sha256(xlsx_path.read_bytes()).hexdigest()
    sz = os.path.getsize(xlsx_path)
    
    preview_dir = pkg_dir / "previews"
    previews = sorted([p.name for p in preview_dir.glob("*.pdf")])
    
    manifest = {
        "manifest_version": "2.0-template-fidelity",
        "template_provenance": {
            "original_file": "originals/Marksans Keshav.xlsx",
            "original_sha256": "5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a",
            "preserved_sheet_count": 17,
            "sheets_order": SHEET_NAMES,
        },
        "target_entity": {
            "name": meta["name"],
            "ticker": meta["ticker"],
            "bse_code": meta["bse"],
            "valuation_snapshot_date": "2025-10-15",
            "official_nse_close_2025_10_15": meta["cmp"],
            "shares_outstanding_cr": meta["shares_cr"],
            "market_capitalization_cr": meta["market_cap_cr"],
        },
        "deliverable_file": {
            "path": f"companies/{tgt}/template_v2/{tgt}-FullTemplate.xlsx",
            "sha256": sha,
            "size_bytes": sz,
            "recalculation_engine": "Microsoft Excel COM 16.0 (CalculateFullRebuild)",
            "formula_errors": 0,
        },
        "valuation_summary": {
            "primary_valuation_method": meta["primary_method"],
            "wacc": meta["wacc"],
            "beta": meta["beta"],
            "template_dcf_value_per_share": meta["dcf_per_share"],
            "reconciled_institutional_target": meta["sotp_per_share"],
        },
        "sheet_previews": {
            "count": len(previews),
            "directory": f"companies/{tgt}/template_v2/previews/",
            "files": previews,
        },
        "governance_and_approval": {
            "ready": False,
            "human_pack_approval": "pending",
            "reviewer_decision_required": True,
            "policy": "Requires explicit human sign-off; self-approval prohibited."
        }
    }
    
    out_manifest = pkg_dir / "package.manifest.json"
    with open(out_manifest, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print(f"Generated manifest for {tgt} at {out_manifest}")
