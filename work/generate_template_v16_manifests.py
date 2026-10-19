"""Generate package.manifest.json for Template V16 models."""

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def compute_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

TARGET_META = {
    "MM": {
        "val": "N/A",
        "method": "Sum-of-the-Parts (SOTP) Valuation with Standalone Net Cash Ring-Fencing",
        "valuation_status": "unavailable",
        "forecast_status": "BLOCKED",
        "provenance_note": "Valuation is marked Unavailable: Factual segment EBITDA inputs (Auto, Farm) and Standalone Net Cash are statutory disclosures not published on Screener.in. Under the locked allowlist rule, they cannot be used in calculation or relabeled as analyst assumptions. Balance sheet forecast is marked BLOCKED: Screener consolidated financials aggregate NBFC loan book (>₹1.3L Cr) into Other Assets without sub-schedules; balance check BS!N40 exposes authentic +7,677.49 Cr gap without artificial plugs.",
        "beta_raw": 1.4746,
        "beta_adj": 1.3164,
        "wacc": 0.1197,
    },
    "LTM": {
        "val": 2620.24,
        "method": "Full-Year Annual FCFF DCF Baseline",
        "valuation_status": "supported",
        "forecast_status": "UNRECONCILED",
        "provenance_note": "Fully supported from allowlisted Screener.in statements, Yahoo Finance 104-week regression beta (1.1018), and Trading Economics 6.51% Rf proxy. Historical balance sheet reconciles across all 11 periods (FY15-FY25 = 0.00 Cr). FY26E balance check BS!N40 evaluates to +1.00 Cr, reflecting Screener's published sub-schedule integer rounding delta (Other Assets sub-schedule sums to ₹14,788 Cr with Other asset items = 4,558 Cr vs top-level ₹14,787 Cr) preserved without artificial balancing plugs.",
        "beta_raw": 1.1526,
        "beta_adj": 1.1018,
        "wacc": 0.1248,
    },
    "CLEAN": {
        "val": 393.03,
        "method": "Full-Year Manufacturing FCFF DCF Baseline",
        "valuation_status": "supported",
        "forecast_status": "UNRECONCILED",
        "provenance_note": "Fully supported from allowlisted Screener.in statements, Yahoo Finance 104-week regression beta (1.2141), and Trading Economics 6.51% Rf proxy. Historical balance sheet reconciles across all 11 periods (FY15-FY25 = 0.00 Cr). FY26E balance check BS!N40 evaluates to -1.13 Cr, reflecting Screener's published sub-schedule rounding delta (Borrowings schedule delta +0.13 Cr and Other Assets schedule delta -1.00 Cr) preserved without artificial balancing plugs.",
        "beta_raw": 1.3212,
        "beta_adj": 1.2141,
        "wacc": 0.1319,
    },
    "EDELWEISS": {
        "val": "N/A",
        "method": "Financial Services Sum-of-the-Parts (SOTP) Valuation (Industrial FCFF Prohibited)",
        "valuation_status": "unavailable",
        "forecast_status": "UNRECONCILED",
        "provenance_note": "Valuation is marked Unavailable: Industrial FCFF is methodologically prohibited for financial institutions. Vertical financial metrics (Alternatives PAT, EARCL BV, Mutual Fund AUM, Insurance IEV) and Standalone HoldCo Net Debt are statutory disclosures not published on Screener.in. Historical balance sheet reconciles across all 11 periods (FY15-FY25 = 0.00 Cr). FY26E balance check BS!N40 evaluates to +1.00 Cr, reflecting Screener's published sub-schedule integer rounding delta (Other Assets sub-schedule sums to ₹22,420 Cr with Other asset items = 15,548 Cr vs top-level ₹22,419 Cr) preserved without artificial balancing plugs.",
        "beta_raw": 2.1231,
        "beta_adj": 1.7487,
        "wacc": 0.1047,
    }
}

SHEET_NAMES = [
    "Model ->", "Main Financials", "Ratio Analysis", "Intrinsic Growth",
    "Comps", "WACC-DCF", "DCF 1st way", "DCF 2nd way", "Data ->",
    "compdata", "Data Sheet", "BS", "Beta", "RM", "peerdata", "Cash Flow"
]

def generate_manifests():
    summary_path = ROOT / "work" / "template_v16_recalc_summary.json"
    if not summary_path.exists():
        raise FileNotFoundError(f"Summary file not found: {summary_path}")

    with open(summary_path, encoding="utf-8") as f:
        recalc_summary = json.load(f)

    for target_key in TARGET_META.keys():
        v16_dir = ROOT / "companies" / target_key / "template_v16"
        wb_path = v16_dir / f"{target_key}-FullTemplate.xlsx"
        if not wb_path.exists():
            raise FileNotFoundError(f"Workbook missing: {wb_path}")

        wb_sha = compute_sha256(wb_path)
        preview_dir = v16_dir / "previews"
        previews = []
        if preview_dir.exists():
            for p in sorted(preview_dir.glob("*.pdf")):
                previews.append({
                    "file": f"previews/{p.name}",
                    "sha256": compute_sha256(p),
                    "size_bytes": p.stat().st_size
                })

        meta = TARGET_META[target_key]
        recalc = recalc_summary[target_key]

        manifest = {
            "schema_version": "2.0.0",
            "package_type": "excel_only_model",
            "target": target_key,
            "version_tag": "template_v16",
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
                "formula_errors": recalc["formula_errors"],
                "balance_check": {
                    "historical_reconciled": recalc["historical_bs_reconciled"],
                    "historical_max_diff_cr": 0.0,
                    "fy26e_difference_cr": recalc["fy26_balance_check_cr"],
                    "fy26e_row40_formula": "=N38-N16",
                    "forecast_status": meta["forecast_status"],
                    "plug_free": True
                }
            },
            "valuation": {
                "primary_method": meta["method"],
                "valuation_status": meta["valuation_status"],
                "fair_value_per_share": recalc["baseline_val"],
                "source_cell": "'DCF 1st way'!E42",
                "cover_link_cell": "'Model ->'!C9",
                "provenance_note": meta["provenance_note"],
                "cost_of_capital": {
                    "evaluated_wacc": recalc["wacc"],
                    "raw_beta": recalc["beta_raw"],
                    "blume_adjusted_beta": recalc["beta_adj"],
                    "risk_free_rate": 0.0651,
                    "equity_risk_premium": 0.055,
                    "marginal_tax_rate": 0.2517,
                    "beta_source": "Yahoo Finance 104-week weekly regression ending 15-Oct-2025 (work/sources/weekly_market_returns_104w_v2.json)",
                    "rf_source": "Trading Economics 14-Oct-2025 prior-session proxy (6.51%)"
                },
                "sensitivities": recalc["sensitivities"]
            },
            "previews": {
                "format": "landscape_pdf",
                "count": len(previews),
                "files": previews
            }
        }

        out_path = v16_dir / "package.manifest.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=2)
        print(f"[{target_key}] Wrote {out_path} ({len(previews)} previews)")

if __name__ == "__main__":
    generate_manifests()
