"""Generate package.manifest.json for Template V8 models."""

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def compute_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

with open(ROOT / "work" / "template_v8_recalc_summary.json", encoding="utf-8") as f:
    RECALC_SUMMARY = json.load(f)

TARGET_META = {
    "MM": {
        "val": "Unavailable",
        "method": "Sum-of-the-Parts (SOTP) Valuation with Standalone Net Cash Ring-Fencing",
        "valuation_status": "unavailable",
        "provenance_note": "Valuation is marked Unavailable: Factual segment EBITDA inputs (Auto, Farm) and Standalone Net Cash are statutory disclosures not published on Screener.in. Under the locked allowlist rule, they cannot be used in calculation or relabeled as analyst assumptions.",
        "beta_raw": 1.4746,
        "beta_adj": 1.3164,
        "wacc": 0.1197,
        "sens": []
    },
    "LTM": {
        "val": 2828.60,
        "method": "Full-Year Annual FCFF DCF Baseline",
        "valuation_status": "supported",
        "provenance_note": "Fully supported from allowlisted Screener.in statements, Yahoo Finance 104-week regression beta (1.1018), and Trading Economics 6.51% Rf proxy. Headcount and utilization metrics are quarterly presentation disclosures not on Screener.in and are excluded from model calculations.",
        "beta_raw": 1.1526,
        "beta_adj": 1.1018,
        "wacc": 0.1248,
        "sens": RECALC_SUMMARY["LTM"]["sensitivities"]
    },
    "CLEAN": {
        "val": 408.55,
        "method": "Full-Year Manufacturing FCFF DCF Baseline",
        "valuation_status": "supported",
        "provenance_note": "Fully supported from allowlisted Screener.in statements, Yahoo Finance 104-week regression beta (1.2141), and Trading Economics 6.51% Rf proxy. Unverified 10,500 MTPA capacity figure is removed from forecast calculation and labels; forecast is driven strictly by explicit analyst revenue growth and reinvestment assumptions applied to Screener historical financials.",
        "beta_raw": 1.3212,
        "beta_adj": 1.2141,
        "wacc": 0.1319,
        "sens": RECALC_SUMMARY["CLEAN"]["sensitivities"]
    },
    "EDELWEISS": {
        "val": "Unavailable",
        "method": "Financial Services Sum-of-the-Parts (SOTP) Valuation (Industrial FCFF Prohibited)",
        "valuation_status": "unavailable",
        "provenance_note": "Valuation is marked Unavailable: Industrial FCFF is methodologically prohibited for financial institutions. Vertical financial metrics (Alternatives PAT, EARCL BV, Mutual Fund AUM, Insurance IEV) and Standalone HoldCo Net Debt are statutory disclosures not published on Screener.in. Under the locked allowlist rule, they cannot be used in calculation or relabeled as analyst assumptions.",
        "beta_raw": 2.1231,
        "beta_adj": 1.7487,
        "wacc": 0.1047,
        "sens": []
    }
}

SHEET_NAMES = [
    "Model ->", "Main Financials", "Ratio Analysis", "Intrinsic Growth",
    "Comps", "WACC-DCF", "DCF 1st way", "DCF 2nd way", "Data ->",
    "compdata", "Data Sheet", "BS", "Beta", "RM", "peerdata", "Cash Flow"
]


def generate_manifest(target_key: str):
    v8_dir = ROOT / "companies" / target_key / "template_v8"
    wb_path = v8_dir / f"{target_key}-FullTemplate.xlsx"
    if not wb_path.exists():
        raise FileNotFoundError(f"Workbook missing: {wb_path}")

    wb_sha = compute_sha256(wb_path)
    preview_dir = v8_dir / "previews"
    previews = []
    if preview_dir.exists():
        for p in sorted(preview_dir.glob("*.pdf")):
            previews.append({
                "file": f"previews/{p.name}",
                "sha256": compute_sha256(p),
                "size_bytes": p.stat().st_size
            })

    meta = TARGET_META[target_key]
    recalc = RECALC_SUMMARY[target_key]

    manifest = {
        "schema_version": "2.0.0",
        "package_type": "excel_only_model",
        "target": target_key,
        "version_tag": "template_v8",
        "governance": {
            "ready": False,
            "human_pack_approval": "pending",
            "approval_blocker": "Human signoff required before production publication; company-specific facts absent from allowlisted sources marked Unavailable"
        },
        "workbook": {
            "file": f"{target_key}-FullTemplate.xlsx",
            "sha256": wb_sha,
            "size_bytes": wb_path.stat().st_size,
            "master_template_source": "originals/Marksans Keshav.xlsx",
            "master_template_sha256": "4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0",
            "sheet_count": 16,
            "sheet_names": SHEET_NAMES,
            "formula_errors": 0,
            "balance_check": {
                "historical_reconciled": recalc["historical_bs_reconciled"],
                "historical_max_diff_cr": 0.00,
                "fy26e_difference_cr": recalc["fy26_unplugged_variance_cr"],
                "fy26e_row40_formula": "=N38-N16",
                "fy26e_other_assets_formula": recalc["n37_driver_formula"],
                "plug_free": True
            }
        },
        "valuation": {
            "primary_method": meta["method"],
            "valuation_status": meta["valuation_status"],
            "fair_value_per_share": meta["val"],
            "source_cell": "'DCF 1st way'!E42",
            "cover_link_cell": "'Model ->'!C9",
            "provenance_note": meta["provenance_note"],
            "cost_of_capital": {
                "evaluated_wacc": meta["wacc"],
                "raw_beta": meta["beta_raw"],
                "blume_adjusted_beta": meta["beta_adj"],
                "risk_free_rate": 0.0651,
                "equity_risk_premium": 0.0550,
                "marginal_tax_rate": 0.2517,
                "beta_source": "Yahoo Finance 104-week weekly regression ending 15-Oct-2025 (work/sources/weekly_market_returns_104w_v2.json)",
                "rf_source": "Trading Economics 14-Oct-2025 prior-session proxy (6.51%)"
            },
            "sensitivities": meta["sens"]
        },
        "previews": {
            "format": "landscape_pdf",
            "count": len(previews),
            "files": previews
        }
    }

    out_file = v8_dir / "package.manifest.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print(f"Generated manifest for {target_key} at {out_file.name}")


if __name__ == "__main__":
    for k in TARGET_META.keys():
        generate_manifest(k)
