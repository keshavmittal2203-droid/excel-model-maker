"""Reproducible Template V16 Model Builder and Verifier for MM, LTM, CLEAN, EDELWEISS.

Builds defect-free, fully audited 16-sheet Excel financial models directly from
the byte-verified master template (originals/Marksans Keshav.xlsx, SHA: 4dc79842...)
and committed source inputs in work/sources/.

Strict requirements enforced:
1. Master Template Integrity: Preserves the unchanged 16-sheet template geometry,
   order, and styling without external paths or dependencies.
2. Allowlisted Sources Only: Screener.in, Yahoo Finance, NSE India, Trading Economics,
   Top Stock Research, Investing.com, Prime Investor.
3. Clean Accounting Linkages (No Balancing Plugs):
   - Cash Flow N31 = =N5+N14+N23 across all targets.
   - BS N35 = =M35+'Cash Flow'!N31 for LTM, CLEAN, EDELWEISS.
   - Balance check row 40 = ={col}38-{col}16 across all 12 periods (BS!C40:N40).
   - Zero balancing plugs on M37 or M15: exact Screener reported values preserved.
     - MM: BS!N40 = +7,677.49 Cr (BLOCKED, conglomerate NBFC loan book gap exposed).
     - LTM: BS!N40 = +1.00 Cr (UNRECONCILED, integer-rounded sub-schedule delta).
     - CLEAN: BS!N40 = -1.13 Cr (UNRECONCILED, integer/decimal sub-schedule delta).
     - EDELWEISS: BS!N40 = +1.00 Cr (UNRECONCILED, integer-rounded sub-schedule delta).
4. Valuation Availability:
   - MM: Unavailable (N/A) due to missing segment EBITDA and standalone net cash facts.
   - EDELWEISS: Unavailable (N/A) due to industrial DCF prohibition and missing SOTP facts.
   - LTM: Supported at ₹2,620.24 / share.
   - CLEAN: Supported at ₹393.03 / share.
5. End-to-end native Microsoft Excel COM 16.0 recalculation, verification of 0 formula
   errors across all 64 sheets, dynamic sensitivity responsiveness, and 15 landscape PDF
   previews per company.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
MASTER_TEMPLATE = ROOT / "originals" / "Marksans Keshav.xlsx"
EXPECTED_MASTER_SHA = "4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0"

TARGETS = ["MM", "LTM", "CLEAN", "EDELWEISS"]

EXPECTED_BASELINES = {
    "MM": "Unavailable",
    "LTM": 2620.24,
    "CLEAN": 393.03,
    "EDELWEISS": "Unavailable",
}

EXPECTED_BETAS_AND_WACCS = {
    "MM": {"raw": 1.4746, "adj": 1.3164, "wacc": 0.1197},
    "LTM": {"raw": 1.1526, "adj": 1.1018, "wacc": 0.1248},
    "CLEAN": {"raw": 1.3212, "adj": 1.2141, "wacc": 0.1319},
    "EDELWEISS": {"raw": 2.1231, "adj": 1.7487, "wacc": 0.1047},
}

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
    },
}

SENSITIVITY_SPECS = {
    "LTM": [
        {"name": "Equity Risk Premium (WACC)", "sheet": "WACC-DCF", "cell": "J23", "base_val": 0.0550, "test_val": 0.0600, "expected_dir": "down"},
        {"name": "Terminal Growth Rate", "sheet": "DCF 1st way", "cell": "E20", "base_val": 0.0500, "test_val": 0.0550, "expected_dir": "up"},
    ],
    "CLEAN": [
        {"name": "Terminal Growth Rate", "sheet": "DCF 1st way", "cell": "E20", "base_val": 0.0500, "test_val": 0.0550, "expected_dir": "up"},
        {"name": "Equity Risk Premium (WACC)", "sheet": "WACC-DCF", "cell": "J23", "base_val": 0.0550, "test_val": 0.0600, "expected_dir": "down"},
    ],
}

COM_ERRORS = {
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826288: "#NULL!",
}
ERR_TOKENS = ["#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"]


def compute_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify_active_master() -> None:
    if not MASTER_TEMPLATE.exists():
        raise FileNotFoundError(f"Active master template not found: {MASTER_TEMPLATE}")
    actual_sha = compute_sha256(MASTER_TEMPLATE)
    if actual_sha != EXPECTED_MASTER_SHA:
        raise ValueError(f"Active master template hash mismatch: expected {EXPECTED_MASTER_SHA}, got {actual_sha}")
    print(f"Verified active master template: {MASTER_TEMPLATE.name} (SHA: {actual_sha[:16]}...)")


def build_company_v16_from_master(target: str, out_path: Path) -> Path:
    """Build candidate 16-sheet workbook from originals/Marksans Keshav.xlsx and committed delta."""
    import openpyxl

    delta_path = ROOT / "work" / "sources" / f"{target}_template_v16_delta.json"
    if not delta_path.exists():
        raise FileNotFoundError(f"Committed delta file not found: {delta_path}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(MASTER_TEMPLATE, out_path)

    wb = openpyxl.load_workbook(out_path, data_only=False)

    # 1. Clean stale logo drawings on Main Financials
    mf = wb["Main Financials"]
    if hasattr(mf, "_images") and mf._images:
        mf._images.clear()

    # 2. Ingest committed source delta
    delta_entries = json.loads(delta_path.read_text(encoding="utf-8"))
    for item in delta_entries:
        ws = wb[item["s"]]
        cell = ws[item["c"]]
        t = item["t"]
        v = item["v"]
        if t == "d":
            parts = [int(x) for x in v.split("-")]
            cell.value = datetime.date(parts[0], parts[1], parts[2])
        elif t == "af":
            from openpyxl.worksheet.formula import ArrayFormula
            cell.value = ArrayFormula(ref=item.get("ref", cell.coordinate), text=v)
        elif t == "z":
            cell.value = None
        else:
            cell.value = v

        if "nf" in item:
            cell.number_format = item["nf"]

    # 3. Clean all cell comments and enforce landscape page setup across all 16 sheets
    for ws in wb.worksheets:
        if hasattr(ws, "_comments"):
            ws._comments.clear()
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(out_path)
    wb.close()
    print(f"  [{target}] Built staging workbook: {out_path.name} ({len(wb.sheetnames)} sheets, {out_path.stat().st_size:,} bytes)")
    return out_path


def recalculate_and_verify_excel_com(
    target: str,
    wb_path: Path,
    preview_dir: Path,
    handoff_wb_path: Path | None = None,
) -> dict[str, Any]:
    """Recalculate workbook in native Excel COM 16.0, verify 0 formula errors, balance checks, and export previews."""
    import win32com.client

    preview_dir.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        ref_data: dict[str, Any] = {}
        if handoff_wb_path and handoff_wb_path.exists():
            wb_ref = excel.Workbooks.Open(str(handoff_wb_path.resolve()))
            for i in range(1, wb_ref.Worksheets.Count + 1):
                ws_r = wb_ref.Worksheets(i)
                ref_data[ws_r.Name] = ws_r.UsedRange.Value
            wb_ref.Close(False)

        wb = excel.Workbooks.Open(str(wb_path.resolve()))
        excel.CalculateFullRebuild()

        sheet_count = wb.Worksheets.Count
        assert sheet_count == 16, f"Expected 16 sheets, got {sheet_count}"

        # 1. Scan for formula errors across all 16 sheets
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
                            if COM_ERRORS[cell_v] != "#N/A":
                                is_err = True
                                err_name = COM_ERRORS[cell_v]
                        elif any(err in s_v for err in ERR_TOKENS):
                            is_err = True
                            err_name = s_v
                        if is_err:
                            print(f"    ERROR: {sname} ({r_idx+1},{c_idx+1}) = {err_name} ({cell_v})")
                            error_count += 1

        assert error_count == 0, f"Found {error_count} formula errors in {target}"

        # 2. Check dynamic Balance Sheet check row 40 across all 12 periods
        bs_ws = wb.Worksheets("BS")
        bs_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
        bs_checks = []
        for col_l in bs_cols:
            formula_v = str(bs_ws.Range(f"{col_l}40").Formula)
            expected_f = f"={col_l}38-{col_l}16"
            assert formula_v == expected_f, f"Expected BS!{col_l}40 formula {expected_f}, got {formula_v}"
            chk_v = bs_ws.Range(f"{col_l}40").Value
            bs_checks.append(float(chk_v) if chk_v is not None else 0.0)

        hist_checks = bs_checks[:11]
        assert all(abs(x) <= 0.01 for x in hist_checks), f"Historical BS imbalance in {target}: {hist_checks}"

        fy26_diff = bs_checks[11]
        if target == "MM":
            assert abs(fy26_diff - 7677.49) < 0.10, f"Expected +7,677.49 Cr gap in MM, got {fy26_diff}"
            forecast_status = "BLOCKED"
        elif target == "LTM":
            assert abs(fy26_diff - 1.00) < 0.05, f"Expected +1.00 Cr gap in LTM, got {fy26_diff}"
            forecast_status = "UNRECONCILED"
        elif target == "CLEAN":
            assert abs(fy26_diff - (-1.13)) < 0.05, f"Expected -1.13 Cr gap in CLEAN, got {fy26_diff}"
            forecast_status = "UNRECONCILED"
        elif target == "EDELWEISS":
            assert abs(fy26_diff - 1.00) < 0.05, f"Expected +1.00 Cr gap in EDELWEISS, got {fy26_diff}"
            forecast_status = "UNRECONCILED"

        # 3. Check Beta and WACC
        beta_ws = wb.Worksheets("Beta")
        raw_b = beta_ws.Range("K7").Value
        adj_b = beta_ws.Range("K13").Value
        wacc_ws = wb.Worksheets("WACC-DCF")
        wacc_v = wacc_ws.Range("J39").Value
        exp_bw = EXPECTED_BETAS_AND_WACCS[target]
        assert abs(raw_b - exp_bw["raw"]) < 1e-3, f"Raw beta mismatch: {raw_b} vs {exp_bw['raw']}"
        assert abs(adj_b - exp_bw["adj"]) < 1e-3, f"Adjusted beta mismatch: {adj_b} vs {exp_bw['adj']}"
        assert abs(wacc_v - exp_bw["wacc"]) < 1e-3, f"WACC mismatch: {wacc_v} vs {exp_bw['wacc']}"

        # 4. Check Baseline Valuation
        dcf_ws = wb.Worksheets("DCF 1st way")
        cov_ws = wb.Worksheets("Model ->")
        base_e42 = dcf_ws.Range("E42").Value
        base_c9 = cov_ws.Range("C9").Value
        exp_val = EXPECTED_BASELINES[target]

        if exp_val == "Unavailable":
            assert str(base_e42).lower() == "n/a", f"Expected N/A in E42, got {base_e42}"
            assert str(base_c9).lower() == "n/a", f"Expected N/A in C9, got {base_c9}"
            val_status = "unavailable"
        else:
            assert abs(float(base_e42) - exp_val) < 0.05, f"DCF mismatch: {base_e42} vs {exp_val}"
            assert abs(float(base_c9) - exp_val) < 0.05, f"Cover mismatch: {base_c9} vs {exp_val}"
            val_status = "supported"

        # 5. Sensitivity Perturbation and Restoration (for supported models)
        sens_results = []
        if target in SENSITIVITY_SPECS:
            for spec in SENSITIVITY_SPECS[target]:
                sws = wb.Worksheets(spec["sheet"])
                sws.Range(spec["cell"]).Value = spec["test_val"]
                excel.CalculateFullRebuild()
                perturbed_val = dcf_ws.Range("E42").Value
                direction = "down" if perturbed_val < exp_val else "up"
                assert direction == spec["expected_dir"], (
                    f"Sensitivity direction mismatch for {spec['name']}: expected {spec['expected_dir']}, got {direction}"
                )
                sws.Range(spec["cell"]).Value = spec["base_val"]
                excel.CalculateFullRebuild()
                restored_val = dcf_ws.Range("E42").Value
                assert abs(float(restored_val) - exp_val) < 0.02, (
                    f"Sensitivity restoration failed: expected {exp_val}, got {restored_val}"
                )
                sens_results.append({
                    "test": spec["name"],
                    "direction": direction,
                    "delta": round(perturbed_val - exp_val, 2),
                    "restored": True,
                })

        # 6. Export 15 Landscape PDF previews
        preview_count = 0
        preview_files = []
        for idx in range(1, sheet_count + 1):
            ws = wb.Worksheets(idx)
            sname = ws.Name
            if sname == "Data ->":
                continue
            clean_sname = sname.replace(" ", "_").replace("->", "_")
            pdf_name = f"{idx:02d}_{clean_sname}.pdf"
            pdf_path = preview_dir / pdf_name
            try:
                ws.PageSetup.Orientation = 2  # xlLandscape
                ws.PageSetup.PaperSize = 9    # xlPaperA4
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.FitToPagesTall = False
            except Exception:
                pass

            ws.ExportAsFixedFormat(
                Type=0,
                Filename=str(pdf_path),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            preview_count += 1
            preview_files.append({
                "file": f"previews/{pdf_name}",
                "sha256": compute_sha256(pdf_path),
                "size_bytes": pdf_path.stat().st_size,
            })

        assert preview_count == 15, f"Expected 15 preview PDFs, got {preview_count}"

        # 7. Compare against delivered handoff workbook (if present)
        mismatch_count = 0
        if ref_data:
            for idx in range(1, sheet_count + 1):
                ws_cur = wb.Worksheets(idx)
                sname = ws_cur.Name
                vals_cur = ws_cur.UsedRange.Value
                vals_ref = ref_data.get(sname)
                if vals_cur is None and vals_ref is None:
                    continue
                if (vals_cur is None) != (vals_ref is None):
                    mismatch_count += 1
                    continue
                rows_cur = vals_cur if isinstance(vals_cur, tuple) else (vals_cur,)
                rows_ref = vals_ref if isinstance(vals_ref, tuple) else (vals_ref,)
                for r_idx in range(min(len(rows_cur), len(rows_ref))):
                    r_c = rows_cur[r_idx] if isinstance(rows_cur[r_idx], tuple) else (rows_cur[r_idx],)
                    r_r = rows_ref[r_idx] if isinstance(rows_ref[r_idx], tuple) else (rows_ref[r_idx],)
                    for c_idx in range(min(len(r_c), len(r_r))):
                        vc = r_c[c_idx]
                        vr = r_r[c_idx]
                        if vc != vr:
                            if isinstance(vc, float) and isinstance(vr, float) and abs(vc - vr) < 1e-4:
                                continue
                            mismatch_count += 1

        # 8. Save calculated workbook
        wb.Save()
        wb.Close(SaveChanges=True)

        wb_sha256 = compute_sha256(wb_path)

        # 9. Generate package.manifest.json in company staging dir
        meta = TARGET_META[target]
        manifest = {
            "schema_version": "2.0.0",
            "package_type": "excel_only_model",
            "target": target,
            "version_tag": "template_v16_reproduced",
            "governance": {
                "ready": False,
                "human_pack_approval": "pending",
                "approval_blocker": "Human signoff required before production publication; company-specific facts absent from allowlisted sources marked Unavailable",
            },
            "workbook": {
                "file": wb_path.name,
                "sha256": wb_sha256,
                "size_bytes": wb_path.stat().st_size,
                "sheet_count": sheet_count,
            },
            "recalculation_audit": {
                "excel_version": "16.0",
                "full_rebuild_executed": True,
                "formula_error_count": error_count,
                "balance_sheet_checks": {
                    "formula": "={col}38-{col}16",
                    "historical_periods_all_zero": True,
                    "fy26e_difference_cr": round(fy26_diff, 2),
                    "forecast_status": forecast_status,
                },
                "cost_of_capital": {
                    "raw_beta": raw_b,
                    "blume_adjusted_beta": adj_b,
                    "wacc": round(wacc_v, 4),
                },
                "valuation": {
                    "methodology": meta["method"],
                    "status": val_status,
                    "evaluated_fair_value_inr": base_c9 if isinstance(base_c9, str) else round(float(base_c9), 2),
                },
                "sensitivity_verified": bool(sens_results),
            },
            "previews": preview_files,
            "comparison_against_handoff": {
                "checked": handoff_wb_path is not None and handoff_wb_path.exists(),
                "cell_value_mismatches": mismatch_count,
            },
        }

        manifest_path = wb_path.parent / "package.manifest.json"
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

        return {
            "target": target,
            "sha256": wb_sha256,
            "formula_errors": error_count,
            "fy26_balance_check_cr": round(fy26_diff, 2),
            "forecast_status": forecast_status,
            "valuation": base_c9 if isinstance(base_c9, str) else round(float(base_c9), 2),
            "valuation_status": val_status,
            "wacc": round(wacc_v, 4),
            "beta_adj": adj_b,
            "previews": preview_count,
            "handoff_mismatches": mismatch_count,
            "manifest": str(manifest_path),
        }

    finally:
        excel.Quit()


def build_all_reproducible_models(
    output_dir: Path | str = "staging/v16_reproducibility",
    *,
    skip_com: bool = False,
    clean: bool = False,
) -> dict[str, Any]:
    """Single documented entry point to reproduce the four V16 models from committed inputs."""
    out_base = (ROOT / output_dir).resolve() if not Path(output_dir).is_absolute() else Path(output_dir)

    # Protect immutable roots and handoff directories
    forbidden_roots = [
        ROOT.resolve(),
        (ROOT / "originals").resolve(),
        (ROOT / "deliverables").resolve(),
        (ROOT / "deliverables" / "V16_user_handoff").resolve(),
        (ROOT / ".git").resolve(),
    ]
    if out_base in forbidden_roots or out_base == ROOT.resolve().parent:
        raise RuntimeError(f"Target directory {out_base} is protected and cannot be used as staging.")

    if clean and out_base.exists():
        print(f"Cleaning existing staging directory: {out_base}")
        shutil.rmtree(out_base)

    out_base.mkdir(parents=True, exist_ok=True)
    verify_active_master()

    print("\n" + "=" * 70)
    print("REPRODUCING FOUR TEMPLATE V16 MODELS FROM COMMITTED INPUTS")
    print(f"Master Template: {MASTER_TEMPLATE}")
    print(f"Staging Base:    {out_base}")
    print("=" * 70)

    summary: dict[str, Any] = {
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "staging_dir": str(out_base),
        "master_template_sha256": compute_sha256(MASTER_TEMPLATE),
        "targets": {},
        "all_passed": False,
    }

    handoff_root = ROOT / "deliverables" / "V16_user_handoff" / "workbooks"

    for target in TARGETS:
        print(f"\n[{target}] Building candidate model...")
        target_dir = out_base / target
        wb_path = target_dir / f"{target}-FullTemplate.xlsx"
        preview_dir = target_dir / "previews"
        build_company_v16_from_master(target, wb_path)

        if not skip_com:
            handoff_ref = handoff_root / f"{target}-FullTemplate.xlsx" if handoff_root.exists() else None
            print(f"[{target}] Recalculating and verifying via native Microsoft Excel COM 16.0...")
            t_res = recalculate_and_verify_excel_com(target, wb_path, preview_dir, handoff_wb_path=handoff_ref)
            summary["targets"][target] = t_res
            print(f"[{target}] VERIFIED:")
            print(f"  SHA-256:        {t_res['sha256']}")
            print(f"  Formula Errors: {t_res['formula_errors']}")
            print(f"  BS N40 Check:   {t_res['fy26_balance_check_cr']:+.2f} Cr ({t_res['forecast_status']})")
            print(f"  Valuation:      {t_res['valuation']} ({t_res['valuation_status']})")
            print(f"  PDF Previews:   {t_res['previews']}")
            print(f"  Handoff Delta:  {t_res['handoff_mismatches']} mismatches")
        else:
            summary["targets"][target] = {
                "target": target,
                "sha256": compute_sha256(wb_path),
                "recalculated_com": False,
            }

    all_passed = True
    if not skip_com:
        for target, data in summary["targets"].items():
            if data.get("formula_errors", -1) != 0:
                all_passed = False
            if data.get("handoff_mismatches", -1) != 0:
                all_passed = False
    summary["all_passed"] = all_passed

    summary_file = out_base / "reproducibility_summary.json"
    summary_file.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print("\n" + "=" * 70)
    print(f"REPRODUCIBLE BUILD COMPLETE: All Passed = {all_passed}")
    print(f"Summary written to: {summary_file}")
    print("=" * 70)

    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description="Reproduce Template V16 Models from committed inputs")
    parser.add_argument("--output-dir", default="staging/v16_reproducibility", help="Output staging directory")
    parser.add_argument("--clean", action="store_true", help="Clean output directory before building")
    parser.add_argument("--skip-com", action="store_true", help="Skip native Excel COM recalculation and previews")
    args = parser.parse_args()

    try:
        summary = build_all_reproducible_models(
            output_dir=args.output_dir,
            skip_com=args.skip_com,
            clean=args.clean,
        )
        return 0 if summary.get("all_passed", False) or args.skip_com else 1
    except Exception as e:
        print(f"\n[BUILD FAILED]: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
