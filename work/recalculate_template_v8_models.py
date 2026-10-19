"""Isolated Excel COM Recalculation, Verification, Sensitivity, and Preview Export for Template V8 Models."""

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
COM_ERRORS = {
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826288: "#NULL!"
}
ERR_TOKENS = ["#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"]

EXPECTED_BASELINES = {
    "MM": "Unavailable",
    "LTM": 2828.60,
    "CLEAN": 408.55,
    "EDELWEISS": "Unavailable",
}

EXPECTED_BETAS_AND_WACCS = {
    "MM": {"raw": 1.4746, "adj": 1.3164, "wacc": 0.1197},
    "LTM": {"raw": 1.1526, "adj": 1.1018, "wacc": 0.1248},
    "CLEAN": {"raw": 1.3212, "adj": 1.2141, "wacc": 0.1319},
    "EDELWEISS": {"raw": 2.1231, "adj": 1.7487, "wacc": 0.1047},
}

SENSITIVITY_SPECS = {
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
            "name": "Terminal Growth Rate",
            "sheet": "DCF 1st way",
            "cell": "E20",
            "base_val": 0.0500,
            "test_val": 0.0550,
            "expected_dir": "up", # Higher g -> higher share value
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
            print(f"VERIFYING & RECALCULATING: {target} (Template V8)")
            print(f"=======================================================")
            wb_path = ROOT / "companies" / target / "template_v8" / f"{target}-FullTemplate.xlsx"
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
                    rows = vals if isinstance(vals, tuple) else (vals,)
                    for r_idx, row in enumerate(rows):
                        if not isinstance(row, tuple):
                            row = (row,)
                        for c_idx, cell_v in enumerate(row):
                            s_v = str(cell_v) if cell_v is not None else ""
                            is_err = False
                            err_name = ""
                            if isinstance(cell_v, int) and cell_v in COM_ERRORS:
                                if COM_ERRORS[cell_v] != "#N/A": # Allow intentional #N/A if any, but flag others
                                    is_err = True
                                    err_name = COM_ERRORS[cell_v]
                            elif any(err in s_v for err in ERR_TOKENS):
                                is_err = True
                                err_name = s_v
                            if is_err:
                                print(f"  ERROR: {sname} cell ({r_idx+1},{c_idx+1}) = {err_name} ({cell_v})")
                                error_count += 1

            print(f"Total formula errors across 16 sheets: {error_count}")
            assert error_count == 0, f"{target} has {error_count} formula errors"

            # 3. Read baseline valuation
            dcf_ws = wb.Worksheets("DCF 1st way")
            model_ws = wb.Worksheets("Model ->")
            base_e42 = dcf_ws.Range("E42").Value
            base_c9 = model_ws.Range("C9").Value
            cov_c13 = model_ws.Range("C13").Value
            cov_c16 = model_ws.Range("C16").Value
            print(f"Baseline 'DCF 1st way'!E42 = {base_e42}")
            print(f"Baseline 'Model ->'!C9     = {base_c9}")
            print(f"Cover Upside/Downside C13  = {cov_c13}")

            exp_val = EXPECTED_BASELINES[target]
            if exp_val == "Unavailable":
                assert str(base_e42).lower() == "unavailable", f"Expected 'Unavailable' for {target}, got {base_e42}"
                assert str(base_c9).lower() == "unavailable", f"Expected 'Unavailable' for cover C9, got {base_c9}"
                assert str(cov_c13).startswith("N/A"), f"Expected N/A in cover C13, got {cov_c13}"
                print(f"  Confirmed: Valuation correctly marked Unavailable with N/A upside/downside.")
            else:
                diff = abs(float(base_e42) - exp_val)
                print(f"Difference vs expected {exp_val}: {diff:.4f}")
                assert diff < 0.05, f"Baseline mismatch: expected {exp_val}, got {base_e42}"
                assert abs(float(base_c9) - exp_val) < 0.05, f"Cover link mismatch: expected {exp_val}, got {base_c9}"

            # 4. Check Balance Sheet reconciliation
            bs_ws = wb.Worksheets("BS")
            bs_reconciled = True
            hist_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
            for c_let in hist_cols:
                liab = bs_ws.Range(f"{c_let}16").Value
                assets = bs_ws.Range(f"{c_let}38").Value
                if liab is not None and assets is not None:
                    d_bs = abs(float(liab) - float(assets))
                    if d_bs > 0.01:
                        print(f"  BS MISMATCH Historical Col {c_let}: Liab={liab}, Assets={assets}, Diff={d_bs}")
                        bs_reconciled = False
                else:
                    raise AssertionError(f"Historical BS totals are missing in column {c_let}: liabilities={liab!r}, assets={assets!r}")

            # Col N (FY26E) check: verify no plug in N37, verify Row 40 balance check == 0.00 Cr
            fy26_liab = bs_ws.Range("N16").Value
            fy26_assets = bs_ws.Range("N38").Value
            fy26_diff = round(float(fy26_assets) - float(fy26_liab), 2)
            n37_formula = str(bs_ws.Range("N37").Formula)
            row40_n = bs_ws.Range("N40").Value
            print(f"  Col N (FY26E) Balance Tie: Assets={fy26_assets:,.2f}, Liab={fy26_liab:,.2f}, Diff={fy26_diff:+.2f} Cr")
            print(f"  Col N (FY26E) N37 Driver Formula: {n37_formula}")
            print(f"  Col N (FY26E) Row 40 Balance Check (N38 - N16): {float(row40_n):+.2f} Cr")
            assert "16" not in n37_formula, f"Col N Row 37 contains plug reference to liabilities (Row 16): {n37_formula}"
            assert abs(float(row40_n)) <= 0.01, f"Col N Row 40 closing balance does not reconcile to zero: {row40_n}"
            assert abs(float(row40_n) - fy26_diff) < 0.01, f"N40 {row40_n} does not match N38-N16 {fy26_diff}"

            print(f"Historical Balance Sheet (BS) Reconciled (Columns C-M, absolute difference <= 0.01 Cr): {bs_reconciled}")
            assert bs_reconciled, f"Historical BS failed reconciliation in {target}"

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

            exp_b_meta = EXPECTED_BETAS_AND_WACCS[target]
            assert abs(float(raw_b) - exp_b_meta["raw"]) < 0.01, f"Raw beta mismatch: {raw_b} vs {exp_b_meta['raw']}"
            assert abs(float(adj_b) - exp_b_meta["adj"]) < 0.01, f"Adjusted beta mismatch: {adj_b} vs {exp_b_meta['adj']}"
            assert abs(float(wacc_v) - exp_b_meta["wacc"]) < 0.001, f"WACC mismatch: {wacc_v} vs {exp_b_meta['wacc']}"

            # 7. Dynamic Sensitivity Testing (LTM & CLEAN)
            sens_results = []
            if target in SENSITIVITY_SPECS:
                sens_tests = SENSITIVITY_SPECS[target]
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
                    elif e_dir == "down":
                        assert delta < 0, f"Expected downward movement for {t_spec['name']}, got {delta}"

                    # Restore baseline
                    test_ws.Range(c_addr).Value = b_val
                    excel.CalculateFullRebuild()
                    restored_e42 = float(dcf_ws.Range("E42").Value)
                    assert abs(restored_e42 - float(base_e42)) < 1e-4, f"Failed exact restoration: {restored_e42} vs {base_e42}"

                    sens_results.append({
                        "name": t_spec["name"],
                        "parameter": f"{ws_name}!{c_addr}",
                        "base": b_val,
                        "perturbed": p_val,
                        "base_output": base_e42,
                        "perturbed_output": pert_e42,
                        "delta": delta,
                        "direction": "up" if delta > 0 else "down",
                        "restored": restored_e42,
                        "passed": True,
                    })
            else:
                print(f"  Sensitivities: Inapplicable (Primary valuation is marked Unavailable on allowlisted sources).")

            # Save recalculated state
            wb.Save()

            # 8. Export 16 Native Previews per target (Landscape, Native scaling)
            print(f"Exporting 16 native sheet previews to {preview_dir.name}...")
            preview_files = []
            for ws_idx in range(1, sheet_count + 1):
                ws = wb.Worksheets(ws_idx)
                safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in ws.Name)
                pdf_path = preview_dir / f"{ws_idx:02d}_{safe_name}.pdf"
                try:
                    # 0 = xlTypePDF, Quality=0 (xlQualityStandard)
                    ws.ExportAsFixedFormat(0, str(pdf_path), 0, True, False)
                    preview_files.append(str(pdf_path.name))
                except Exception as e:
                    print(f"  Warning exporting PDF for {ws.Name}: {e}")

            print(f"Successfully exported {len(preview_files)}/16 PDFs")
            wb.Close(SaveChanges=True)

            final_hash = compute_sha256(wb_path)
            results[target] = {
                "sha256": final_hash,
                "baseline_val": float(base_e42) if isinstance(base_e42, (int, float)) else str(base_e42),
                "cover_val": float(base_c9) if isinstance(base_c9, (int, float)) else str(base_c9),
                "valuation_status": "unavailable" if exp_val == "Unavailable" else "supported",
                "formula_errors": error_count,
                "historical_bs_reconciled": bs_reconciled,
                "fy26_unplugged_variance_cr": fy26_diff,
                "n37_driver_formula": n37_formula,
                "wacc": float(wacc_v),
                "beta_raw": float(raw_b),
                "beta_adj": float(adj_b),
                "sensitivities": sens_results,
                "preview_count": len(preview_files),
            }

    finally:
        excel.Quit()

    summary_path = ROOT / "work" / "template_v8_recalc_summary.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)
    print(f"\nRecalculation summary saved to {summary_path}")
    return results


if __name__ == "__main__":
    res = run_recalculation_and_verification()
    print("\n=======================================================")
    print("FINAL SUMMARY MATRIX (Template V8 Recalculation):")
    print("=======================================================")
    for k, v in res.items():
        base_display = f"INR {v['baseline_val']:8.2f}" if isinstance(v['baseline_val'], (int, float)) else f"{v['baseline_val']:>12s}"
        cover_display = f"INR {v['cover_val']:8.2f}" if isinstance(v['cover_val'], (int, float)) else f"{v['cover_val']:>12s}"
        print(f"Target: {k:9s} | Base: {base_display} | Cover: {cover_display} | WACC: {v['wacc']*100:5.2f}% | BS Reconciled: {v['historical_bs_reconciled']} | FY26 Diff: {v['fy26_unplugged_variance_cr']:.2f} Cr | Errors: {v['formula_errors']} | Previews: {v['preview_count']}")
        print(f"  SHA-256: {v['sha256']}")
