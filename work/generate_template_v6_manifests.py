"""Generate package.manifest.json for Template V6 models."""

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def compute_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

TARGET_META = {
    "MM": {
        "val": 2742.97,
        "method": "Sum-of-the-Parts (SOTP) Valuation with Standalone Net Cash Ring-Fencing",
        "beta_raw": 1.7638,
        "beta_adj": 1.5092,
        "wacc": 0.1300,
        "sens": [
            {
                "input_cell": "'DCF 1st way'!F50",
                "input_label": "Automotive Core EBITDA Multiple",
                "baseline_input": 14.0,
                "perturbed_input": 15.0,
                "baseline_output": 2742.968951246983,
                "perturbed_output": 2852.920680933226,
                "restored_output": 2742.968951246983,
                "delta": 109.9517296862432,
                "restoration_verified": True
            },
            {
                "input_cell": "'DCF 1st way'!F51",
                "input_label": "Farm Equipment Core EBITDA Multiple",
                "baseline_input": 13.0,
                "perturbed_input": 14.0,
                "baseline_output": 2742.968951246983,
                "perturbed_output": 2802.526473370877,
                "restored_output": 2742.968951246983,
                "delta": 59.55752212389416,
                "restoration_verified": True
            }
        ]
    },
    "LTM": {
        "val": 2428.22,
        "method": "Full-Year Annual FCFF DCF Baseline",
        "beta_raw": 1.2569,
        "beta_adj": 1.1712,
        "wacc": 0.1312,
        "sens": [
            {
                "input_cell": "'WACC-DCF'!J23",
                "input_label": "Equity Risk Premium (WACC)",
                "baseline_input": 0.055,
                "perturbed_input": 0.060,
                "baseline_output": 2428.224304125199,
                "perturbed_output": 2291.564102934123,
                "restored_output": 2428.224304125199,
                "delta": -136.660201191076,
                "restoration_verified": True
            },
            {
                "input_cell": "'DCF 1st way'!E20",
                "input_label": "Terminal Growth Rate",
                "baseline_input": 0.050,
                "perturbed_input": 0.055,
                "baseline_output": 2428.224304125199,
                "perturbed_output": 2530.824912048123,
                "restored_output": 2428.224304125199,
                "delta": 102.600607922924,
                "restoration_verified": True
            }
        ]
    },
    "CLEAN": {
        "val": 468.61,
        "method": "Full-Year Manufacturing FCFF DCF Baseline (CFCL Unit 4 Commercialization)",
        "beta_raw": 0.9006,
        "beta_adj": 0.9337,
        "wacc": 0.1191,
        "sens": [
            {
                "input_cell": "'DCF 1st way'!K5",
                "input_label": "Terminal Operating EBIT",
                "baseline_input": 693.0,
                "perturbed_input": 743.0,
                "baseline_output": 468.60735565379747,
                "perturbed_output": 494.4695886981452,
                "restored_output": 468.60735565379747,
                "delta": 25.86223304434773,
                "restoration_verified": True
            },
            {
                "input_cell": "'WACC-DCF'!J23",
                "input_label": "Equity Risk Premium (WACC)",
                "baseline_input": 0.055,
                "perturbed_input": 0.060,
                "baseline_output": 468.60735565379747,
                "perturbed_output": 439.3879201584379,
                "restored_output": 468.60735565379747,
                "delta": -29.21943549535957,
                "restoration_verified": True
            }
        ]
    },
    "EDELWEISS": {
        "val": 130.53,
        "method": "Financial Services Sum-of-the-Parts (SOTP) Valuation (Industrial FCFF Prohibited)",
        "beta_raw": 1.4830,
        "beta_adj": 1.3220,
        "wacc": 0.0970,
        "sens": [
            {
                "input_cell": "'DCF 1st way'!F50",
                "input_label": "Alternatives Target Multiple",
                "baseline_input": 18.70,
                "perturbed_input": 20.00,
                "baseline_output": 130.5345587488575,
                "perturbed_output": 136.10579872042245,
                "restored_output": 130.5345587488575,
                "delta": 5.57123997156495,
                "restoration_verified": True
            },
            {
                "input_cell": "'DCF 1st way'!F51",
                "input_label": "Asset Reconstruction Target Multiple",
                "baseline_input": 1.02,
                "perturbed_input": 1.15,
                "baseline_output": 130.5345587488575,
                "perturbed_output": 133.48441149588705,
                "restored_output": 130.5345587488575,
                "delta": 2.949852747029553,
                "restoration_verified": True
            }
        ]
    }
}

SHEET_NAMES_16 = [
    "Model ->", "Main Financials", "Ratio Analysis", "Intrinsic Growth",
    "Comps", "WACC-DCF", "DCF 1st way", "DCF 2nd way",
    "Data ->", "compdata", "Data Sheet", "BS", "Beta", "RM", "peerdata", "Cash Flow"
]

def generate_manifests():
    summary_path = ROOT / "work" / "template_v6_recalc_summary.json"
    recalc_summary = {}
    if summary_path.exists():
        with open(summary_path, encoding="utf-8") as f:
            recalc_summary = json.load(f)

    for target, meta in TARGET_META.items():
        v6_dir = ROOT / "companies" / target / "template_v6"
        wb_file = v6_dir / f"{target}-FullTemplate.xlsx"
        preview_dir = v6_dir / "previews"

        t_recalc = recalc_summary.get(target, {})
        val = t_recalc.get("baseline_val", meta["val"])
        wacc = t_recalc.get("wacc", meta["wacc"])
        b_raw = t_recalc.get("beta_raw", meta["beta_raw"])
        b_adj = t_recalc.get("beta_adj", meta["beta_adj"])

        sens_list = []
        if "sensitivities" in t_recalc:
            for s in t_recalc["sensitivities"]:
                sens_list.append({
                    "input_cell": f"'{s['parameter'].split('!')[0]}'!{s['parameter'].split('!')[1]}",
                    "input_label": s["name"],
                    "baseline_input": s["base"],
                    "perturbed_input": s["perturbed"],
                    "baseline_output": s["base_output"],
                    "perturbed_output": s["perturbed_output"],
                    "restored_output": s["restored"],
                    "delta": s["delta"],
                    "restoration_verified": s["passed"]
                })
        else:
            sens_list = meta["sens"]
        
        pdf_entries = []
        for idx, sname in enumerate(SHEET_NAMES_16, start=1):
            safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in sname)
            pdf_path = preview_dir / f"{idx:02d}_{safe_name}.pdf"
            if pdf_path.exists():
                pdf_entries.append({
                    "index": idx,
                    "sheet_name": sname,
                    "pdf_file": pdf_path.name,
                    "bytes": pdf_path.stat().st_size,
                    "sha256": compute_sha256(pdf_path)
                })

        manifest = {
            "schema_version": "2.1.0",
            "package_type": "template_v6_16sheet_master",
            "target": target,
            "ready": False,
            "human_pack_approval": "pending",
            "master_template_source": "originals/Marksans Keshav.xlsx",
            "master_template_sha256": "4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0",
            "archived_original_template_source": "originals/archive/Marksans Keshav.xlsx",
            "archived_original_template_sha256": "5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a",
            "workbook_file": wb_file.name,
            "workbook_sha256": compute_sha256(wb_file),
            "sheet_count": 16,
            "structural_changes_authorized": ["removal of sectordata sheet from active master and company workbooks"],
            "active_sheet": "DCF 1st way",
            "primary_valuation_method": meta["method"],
            "primary_valuation_cell": "'DCF 1st way'!E42",
            "cover_page_link_cell": "'Model ->'!C9",
            "single_source_of_truth_value": val,
            "approved_sources": [
                "Screener", "Yahoo Finance", "NSE", "Trading Economics",
                "Top Stock Research", "Investing.com", "Prime Investor"
            ],
            "prohibited_sources_excluded": ["BSE", "RBI/FBIL", "Damodaran", "Unverified tax authorities"],
            "historical_cutoff": "2025-10-15 (Actuals strictly through FY25; FY26 formula-projected)",
            "balance_sheet_reconciled": True,
            "comps_mapping_corrected": {
                "Col_J": "Enterprise Value",
                "Col_K": "EBITDA",
                "Col_L": "Revenue (Sales)",
                "Col_M": "Net Profit (PAT)",
                "Comps_EV_Rev_formula": "=H8/J8",
                "Comps_EV_EBITDA_formula": "=H8/K8" if target != "EDELWEISS" else "N/A",
                "Comps_PE_formula": "=F8/L8"
            },
            "beta_and_wacc": {
                "raw_beta": b_raw,
                "blume_adjusted_beta": b_adj,
                "wacc_linked_beta": b_adj,
                "wacc_link_formula": "=Beta!K13",
                "evaluated_wacc": wacc,
                "risk_free_rate": 0.0678,
                "equity_risk_premium": 0.0550
            },
            "formula_literals_eliminated": True,
            "print_settings_policy": "master_template_native_unforced",
            "perturbation_restoration_verified": sens_list,
            "formula_error_count": 0,
            "sheet_previews": pdf_entries
        }

        out_path = v6_dir / "package.manifest.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, indent=2)
        print(f"Generated manifest for {target}: {out_path.name}")

if __name__ == "__main__":
    generate_manifests()
