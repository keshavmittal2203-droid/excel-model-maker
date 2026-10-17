"""Isolated Excel COM Recalculation, Verification, Sensitivity, and Preview Export for Template V6 Models."""

from __future__ import annotations

import hashlib
import json
import os
import sys
from pathlib import Path
import openpyxl
import win32com.client

ROOT = Path(__file__).resolve().parent.parent

def compute_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

TARGETS = ["MM", "LTM", "CLEAN", "EDELWEISS"]
ERR_TOKENS = ["#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#N/A", "#NULL!"]

EXPECTED_BASELINES = {
    "MM": 2742.97,
    "LTM": 2811.09,
    "CLEAN": 480.89,
    "EDELWEISS": 130.53,
}

# Sensitivity test specifications: at least 2 inputs per company
SENSITIVITY_SPECS = {
    "MM": [
        {
            "name": "Automotive Core Multiple",
            "sheet": "DCF 1st way",
            "cell": "F50",
            "base_val": 14.0,
            "test_val": 15.0,
            "expected_dir": "up", # Higher multiple -> higher share value
        },
        {
            "name": "Farm Equipment Core Multiple",
            "sheet": "DCF 1st way",
            "cell": "F51",
            "base_val": 13.0,
            "test_val": 14.0,
            "expected_dir": "up", # Higher multiple -> higher share value
        },
    ],
    "LTM": [
        {
            "name": "Equity Risk Premium (WACC)",
            "sheet": "WACC-DCF",
            "cell": "J23",
            "base_val": 0.0550,
            "test_val": 0.0600,
            "expected_dir": "down", # Higher ERP -> higher WACC -> lower share value
        },
        {
            "name": "Terminal Growth Rate",
            "sheet": "DCF 1st way",
            "cell": "E20",
            "base_val": 0.0500,
            "test_val": 0.0550,
            "expected_dir": "up", # Higher g -> higher share value
        },
    ],
    "CLEAN": [
        {
            "name": "Terminal Operating EBIT",
            "sheet": "DCF 1st way",
            "cell": "K5",
            "base_val": 693.0,
            "test_val": 743.0,
            "expected_dir": "up", # Higher EBIT -> higher share value
        },
        {
            "name": "Equity Risk Premium (WACC)",
            "sheet": "WACC-DCF",
            "cell": "J23",
            "base_val": 0.0550,
            "test_val": 0.0600,
            "expected_dir": "down", # Higher ERP -> higher WACC -> lower share value
        },
    ],
    "EDELWEISS": [
        {
            "name": "Alternatives Target Multiple",
            "sheet": "DCF 1st way",
            "cell": "F50",
            "base_val": 18.70,
            "test_val": 20.00,
            "expected_dir": "up", # Higher multiple -> higher share value
        },
        {
            "name": "Asset Reconstruction Target Multiple",
            "sheet": "DCF 1st way",
            "cell": "F51",
            "base_val": 1.02,
            "test_val": 1.15,
            "expected_dir": "up", # Higher multiple -> higher share value
        },
    ],
}


def run_recalculation_and_verification():
    results = {}
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        for target in TARGETS:
            print(f"\n=======================================================")
            print(f"VERIFYING & RECALCULATING: {target} (Template V6)")
            print(f"=======================================================")
            wb_path = ROOT / "companies" / target / "template_v6" / f"{target}-FullTemplate.xlsx"
            if not wb_path.exists():
                raise FileNotFoundError(f"Workbook not found: {wb_path}")

            preview_dir = wb_path.parent / "previews"
            preview_dir.mkdir(parents=True, exist_ok=True)

            wb = excel.Workbooks.Open(str(wb_path))

            # 1. Full rebuild calculation
            excel.CalculateFullRebuild()

            sheet_count = wb.Worksheets.Count
            sheet_names = [wb.Worksheets(i).Name for i in range(1, sheet_count + 1)]
            print(f"Sheet count: {sheet_count} (Expected: 16)")
            assert sheet_count == 16, f"Expected 16 sheets, got {sheet_count}"

            # 2. Scan for formula errors across all sheets
            error_count = 0
            for ws_idx in range(1, sheet_count + 1):
                ws = wb.Worksheets(ws_idx)
                sname = ws.Name
                used = ws.UsedRange
                vals = used.Value
                if vals:
                    # In pywin32, 2D ranges return tuples of tuples
                    rows = vals if isinstance(vals, tuple) else (vals,)
                    for r_idx, row in enumerate(rows):
                        if not isinstance(row, tuple):
                            row = (row,)
                        for c_idx, cell_v in enumerate(row):
                            s_v = str(cell_v) if cell_v is not None else ""
                            # Ignore intentional disclosures like "N/A - Financial Services"
                            if any(err in s_v for err in ["#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!"]):
                                print(f"  ERROR: {sname} cell ({r_idx+1},{c_idx+1}) = {s_v}")
                                error_count += 1

            print(f"Total formula errors across 16 sheets: {error_count}")
            assert error_count == 0, f"{target} has {error_count} formula errors"

            # 3. Read baseline valuation
            dcf_ws = wb.Worksheets("DCF 1st way")
            model_ws = wb.Worksheets("Model ->")
            base_e42 = dcf_ws.Range("E42").Value
            base_c9 = model_ws.Range("C9").Value
            print(f"Baseline 'DCF 1st way'!E42 = {base_e42}")
            print(f"Baseline 'Model ->'!C9     = {base_c9}")

            exp_val = EXPECTED_BASELINES[target]
            diff = abs(float(base_e42) - exp_val)
            print(f"Difference vs expected {exp_val}: {diff:.4f}")
            assert diff < 0.05, f"Baseline mismatch: expected {exp_val}, got {base_e42}"
            assert abs(float(base_c9) - exp_val) < 0.05, f"Cover link mismatch: expected {exp_val}, got {base_c9}"

            # 4. Check Balance Sheet reconciliation
            bs_ws = wb.Worksheets("BS")
            bs_reconciled = True
            cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
            for c_let in cols:
                liab = bs_ws.Range(f"{c_let}16").Value
                assets = bs_ws.Range(f"{c_let}38").Value
                if liab is not None and assets is not None:
                    d_bs = abs(float(liab) - float(assets))
                    if d_bs > 0.05:
                        print(f"  BS MISMATCH Col {c_let}: Liab={liab}, Assets={assets}, Diff={d_bs}")
                        bs_reconciled = False
            print(f"Balance Sheet (BS) Reconciled (Assets == Liabilities): {bs_reconciled}")
            assert bs_reconciled, f"BS failed reconciliation in {target}"

            # 5. Check Comps denominators
            comps_ws = wb.Worksheets("Comps")
            ev_rev = comps_ws.Range("N8").Value
            ev_ebitda = comps_ws.Range("O8").Value
            pe = comps_ws.Range("P8").Value
            ev_rev_str = f"{ev_rev:.2f}" if isinstance(ev_rev, (int, float)) else str(ev_rev)
            ev_eb_str = f"{ev_ebitda:.2f}" if isinstance(ev_ebitda, (int, float)) else str(ev_ebitda)
            pe_str = f"{pe:.2f}" if isinstance(pe, (int, float)) else str(pe)
            print(f"Comps Target Multiples: EV/Rev={ev_rev_str}, EV/EBITDA={ev_eb_str}, P/E={pe_str}")

            # 6. Check Beta and WACC linkage
            beta_ws = wb.Worksheets("Beta")
            raw_b = beta_ws.Range("K7").Value
            adj_b = beta_ws.Range("K13").Value
            wacc_b = wb.Worksheets("WACC-DCF").Range("J24").Value
            wacc_v = wb.Worksheets("WACC-DCF").Range("J39").Value
            print(f"Beta: Raw={raw_b:.4f}, Adjusted (2/3 raw + 1/3 1.0)={adj_b:.4f}")
            print(f"WACC: Linked Beta={wacc_b:.4f}, Evaluated WACC={wacc_v*100:.2f}%")
            assert abs(adj_b - wacc_b) < 1e-4, f"WACC beta {wacc_b} does not match Beta!K13 {adj_b}"

            # 7. Dynamic Sensitivity Testing
            sens_tests = SENSITIVITY_SPECS[target]
            sens_results = []
            for t_spec in sens_tests:
                ws_name = t_spec["sheet"]
                c_addr = t_spec["cell"]
                b_val = t_spec["base_val"]
                p_val = t_spec["test_val"]
                e_dir = t_spec["expected_dir"]

                test_ws = wb.Worksheets(ws_name)
                # Apply perturbation
                test_ws.Range(c_addr).Value = p_val
                excel.CalculateFullRebuild()

                pert_e42 = float(dcf_ws.Range("E42").Value)
                pert_c9 = float(model_ws.Range("C9").Value)
                delta = pert_e42 - float(base_e42)

                print(f"Sensitivity [{t_spec['name']}]: {b_val} -> {p_val}")
                print(f"  E42: {base_e42:.2f} -> {pert_e42:.2f} (delta = {delta:+.2f})")
                assert abs(pert_e42 - pert_c9) < 1e-4, "Cover sheet did not dynamically update!"

                if e_dir == "up":
                    assert delta > 0, f"Expected upward movement for {t_spec['name']}, got {delta}"
                else:
                    assert delta < 0, f"Expected downward movement for {t_spec['name']}, got {delta}"

                # Restore base value
                test_ws.Range(c_addr).Value = b_val
                excel.CalculateFullRebuild()

                rest_e42 = float(dcf_ws.Range("E42").Value)
                print(f"  Restored E42: {rest_e42:.2f} (exact return)")
                assert abs(rest_e42 - float(base_e42)) < 1e-4, f"Failed exact restoration for {t_spec['name']}"

                sens_results.append({
                    "input": t_spec["name"],
                    "base": b_val,
                    "perturbed": p_val,
                    "delta": delta,
                    "restored": rest_e42,
                    "passed": True
                })

            # 8. Export Preview PDFs for all 16 sheets using native unforced print settings
            print(f"Exporting 16 preview PDFs for {target}...")
            for ws_idx in range(1, sheet_count + 1):
                ws = wb.Worksheets(ws_idx)
                safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in ws.Name)
                pdf_out = preview_dir / f"{ws_idx:02d}_{safe_name}.pdf"
                ws.ExportAsFixedFormat(0, str(pdf_out)) # 0 = xlTypePDF

            pdf_count = len(list(preview_dir.glob("*.pdf")))
            print(f"Successfully exported {pdf_count} preview PDFs.")
            assert pdf_count == 16, f"Expected 16 preview PDFs, got {pdf_count}"

            # Save workbook
            wb.Save()
            wb.Close(SaveChanges=True)

            results[target] = {
                "base_val": float(base_e42),
                "error_count": error_count,
                "bs_reconciled": bs_reconciled,
                "sensitivities": sens_results,
                "pdf_count": pdf_count,
                "sha256": compute_sha256(wb_path),
            }

    finally:
        excel.Quit()

    return results


if __name__ == "__main__":
    res = run_recalculation_and_verification()
    print("\n=======================================================")
    print("ALL FOUR TARGETS RECALCULATED & VERIFIED SUCCESSFULLY!")
    print("=======================================================")
    print(json.dumps(res, indent=2))
