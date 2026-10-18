"""Isolated Excel COM Recalculation, Verification, Sensitivity, and Preview Export for Template V9 Models."""

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
            print(f"VERIFYING & RECALCULATING: {target} (Template V9)")
            print(f"=======================================================")
            wb_path = ROOT / "companies" / target / "template_v9" / f"{target}-FullTemplate.xlsx"
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
            assert error_count == 0, f"Found {error_count} formula errors in {target}"

            # 3. Check Balance Sheet reconciliation across all 12 periods
            bs_ws = wb.Worksheets("BS")
            bs_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
            bs_reconciled = True
            for col_l in bs_cols:
                chk_v = bs_ws.Range(f"{col_l}40").Value
                try:
                    num_chk = float(chk_v)
                    if abs(num_chk) > 0.01:
                        print(f"  BS imbalance in {col_l}: {num_chk}")
                        bs_reconciled = False
                except (ValueError, TypeError):
                    print(f"  BS non-numeric check in {col_l}: {chk_v}")
                    bs_reconciled = False

            print(f"Balance Sheet Check (All 12 periods <= 0.01 Cr): {'PASS' if bs_reconciled else 'FAIL'}")
            assert bs_reconciled, f"Balance sheet does not balance to 0.00 Cr in {target}"

            # Check FY26E variance in Row 40
            fy26_diff = float(bs_ws.Range("N40").Value)
            print(f"FY26E Unplugged BS Variance: {fy26_diff:+.4f} Cr")
            assert abs(fy26_diff) <= 0.01, f"FY26E variance exceeds 0.01 Cr threshold: {fy26_diff}"

            # 4. Check Beta and WACC
            beta_ws = wb.Worksheets("Beta")
            raw_b = beta_ws.Range("K7").Value
            adj_b = beta_ws.Range("K13").Value
            wacc_ws = wb.Worksheets("WACC-DCF")
            wacc_v = wacc_ws.Range("J39").Value

            exp_bw = EXPECTED_BETAS_AND_WACCS[target]
            print(f"Beta (Raw / Blume Adj): {raw_b:.4f} / {adj_b:.4f} (Expected: {exp_bw['raw']:.4f} / {exp_bw['adj']:.4f})")
            print(f"WACC: {wacc_v*100:.2f}% (Expected: {exp_bw['wacc']*100:.2f}%)")
            assert abs(raw_b - exp_bw["raw"]) < 1e-3, f"Raw beta mismatch: {raw_b} vs {exp_bw['raw']}"
            assert abs(adj_b - exp_bw["adj"]) < 1e-3, f"Adjusted beta mismatch: {adj_b} vs {exp_bw['adj']}"
            assert abs(wacc_v - exp_bw["wacc"]) < 1e-3, f"WACC mismatch: {wacc_v} vs {exp_bw['wacc']}"

            # 5. Check Baseline Valuation
            dcf_ws = wb.Worksheets("DCF 1st way")
            cov_ws = wb.Worksheets("Model ->")
            base_e42 = dcf_ws.Range("E42").Value
            base_c9 = cov_ws.Range("C9").Value
            exp_val = EXPECTED_BASELINES[target]

            print(f"DCF 1st way E42: {base_e42} (Expected: {exp_val})")
            print(f"Cover Sheet C9: {base_c9} (Expected: {exp_val})")

            if exp_val == "Unavailable":
                assert str(base_e42).lower() == "unavailable", f"Expected Unavailable, got {base_e42}"
                assert str(base_c9).lower() == "unavailable", f"Expected Unavailable on Cover, got {base_c9}"
                comps_ws = wb.Worksheets("Comps")
                pe_implied = comps_ws.Range("P31").Value
                cross_check_lbl = cov_ws.Range("C18").Value
                cross_check_val = cov_ws.Range("D18").Value
                print(f"  Secondary Comps P/E Peer Cross-Check: {pe_implied} (Cover D18: {cross_check_val})")
                assert abs(float(cross_check_val) - float(pe_implied)) < 1e-2, "Cover secondary cross-check mismatch!"
            else:
                assert abs(float(base_e42) - exp_val) < 0.05, f"Baseline value mismatch: {base_e42} vs {exp_val}"
                assert abs(float(base_c9) - exp_val) < 0.05, f"Cover value mismatch: {base_c9} vs {exp_val}"

                # Verify DCF 1st way sensitivity table units: J27 base case must match E42 exactly!
                sens_base_j27 = float(dcf_ws.Range("J27").Value)
                print(f"  DCF 1st way Sensitivity Base Cell (J27): {sens_base_j27:.2f} (E42: {float(base_e42):.2f})")
                assert abs(sens_base_j27 - float(base_e42)) < 0.01, f"Sensitivity base case J27 ({sens_base_j27}) does not reconcile to primary output E42 ({base_e42})"

            # 6. Check Intrinsic Growth ROIC and Net Non-Current Assets
            ig_ws = wb.Worksheets("Intrinsic Growth")
            ig_nca = ig_ws.Range("I34").Value
            ig_roic = ig_ws.Range("I39").Value
            print(f"Intrinsic Growth FY26E Net Non-Current Assets (I34): {ig_nca}")
            print(f"Intrinsic Growth FY26E ROIC (I39): {ig_roic}")
            if target == "EDELWEISS":
                assert "N/A" in str(ig_roic), "EDELWEISS industrial ROIC must be marked N/A"
            elif target == "LTM":
                assert float(ig_nca) > 4000.0, f"Expected positive Net Non-Current Assets for LTM, got {ig_nca}"
                assert 0.30 <= float(ig_roic) <= 0.45, f"Expected realistic ROIC for LTM (30-45%), got {ig_roic}"
            elif target == "CLEAN":
                assert float(ig_nca) > 600.0, f"Expected positive Net Non-Current Assets for CLEAN, got {ig_nca}"
                assert 0.20 <= float(ig_roic) <= 0.35, f"Expected realistic ROIC for CLEAN (20-35%), got {ig_roic}"
            elif target == "MM":
                assert float(ig_nca) > 30000.0, f"Expected positive Net Non-Current Assets for MM, got {ig_nca}"

            # 7. Dynamic Sensitivity Perturbation & Restoration
            sens_results = []
            if target in SENSITIVITY_SPECS:
                for t_spec in SENSITIVITY_SPECS[target]:
                    ws_name = t_spec["sheet"]
                    c_addr = t_spec["cell"]
                    b_val = t_spec["base_val"]
                    p_val = t_spec["test_val"]
                    e_dir = t_spec["expected_dir"]

                    test_ws = wb.Worksheets(ws_name)
                    test_ws.Range(c_addr).Value = p_val
                    excel.CalculateFullRebuild()

                    pert_e42 = float(dcf_ws.Range("E42").Value)
                    pert_c9 = float(cov_ws.Range("C9").Value)
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

            # 8. Export 15 Native Previews per target (Landscape, FitToPagesWide = 1, FitToPagesTall = False; Exclude 'Data ->')
            print(f"Exporting 15 native sheet previews to {preview_dir.name}...")
            preview_files = []
            for ws_idx in range(1, sheet_count + 1):
                ws = wb.Worksheets(ws_idx)
                sname = ws.Name
                if sname.strip().lower() == "data ->":
                    print(f"  Skipping blank navigation sheet '{sname}' from PDF export.")
                    continue

                safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in sname)
                pdf_path = preview_dir / f"{ws_idx:02d}_{safe_name}.pdf"
                try:
                    ps = ws.PageSetup
                    ps.Orientation = 2 # 2 = xlLandscape
                    ps.Zoom = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False

                    # 0 = xlTypePDF, Quality=0 (xlQualityStandard)
                    ws.ExportAsFixedFormat(0, str(pdf_path), 0, True, False)
                    preview_files.append(str(pdf_path.name))
                except Exception as e:
                    print(f"  Warning exporting PDF for {sname}: {e}")

            print(f"Successfully exported {len(preview_files)}/15 PDFs")
            assert len(preview_files) == 15, f"Expected 15 previews, got {len(preview_files)}"
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
                "wacc": float(wacc_v),
                "beta_raw": float(raw_b),
                "beta_adj": float(adj_b),
                "sensitivities": sens_results,
                "preview_count": len(preview_files),
            }

    finally:
        excel.Quit()

    summary_path = ROOT / "work" / "template_v9_recalc_summary.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)
    print(f"\nRecalculation summary saved to {summary_path}")
    return results


if __name__ == "__main__":
    res = run_recalculation_and_verification()
    print("\n=======================================================")
    print("FINAL SUMMARY MATRIX (Template V9 Recalculation):")
    print("=======================================================")
    for k, v in res.items():
        base_display = f"INR {v['baseline_val']:8.2f}" if isinstance(v['baseline_val'], (int, float)) else f"{v['baseline_val']:>12s}"
        cover_display = f"INR {v['cover_val']:8.2f}" if isinstance(v['cover_val'], (int, float)) else f"{v['cover_val']:>12s}"
        print(f"Target: {k:9s} | Base: {base_display} | Cover: {cover_display} | WACC: {v['wacc']*100:5.2f}% | BS Reconciled: {v['historical_bs_reconciled']} | FY26 Diff: {v['fy26_unplugged_variance_cr']:.2f} Cr | Errors: {v['formula_errors']} | Previews: {v['preview_count']}")
        print(f"  SHA-256: {v['sha256']}")
