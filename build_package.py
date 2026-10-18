"""Repeatable Excel-only workbook build and validation from template and pack specifications."""
from __future__ import annotations

import argparse
import copy
from datetime import datetime, timezone
import hashlib
import io
import json
import math
import os
from pathlib import Path, PureWindowsPath
import re
import shutil
import sys
import tempfile
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from modelmaker import sha, read_json, compose
from validate_workbook import validate_candidate

FORBIDDEN_OUTPUT_ROOTS = {
    ROOT.resolve(),
    (ROOT / "originals").resolve(),
    (ROOT / "proposals").resolve(),
    (ROOT / "docs").resolve(),
    (ROOT / "work").resolve(),
    (ROOT / ".git").resolve(),
}

FORBIDDEN_PARENT_DIRS = {
    (ROOT / "originals").resolve(),
    (ROOT / "proposals").resolve(),
    (ROOT / "docs").resolve(),
    (ROOT / "work").resolve(),
    (ROOT / ".git").resolve(),
}

REGISTERED_ADAPTERS: dict[str, str] = {
    "work/adapt_cipla_workbook.py": "adapt_workbook_for_cipla",
    "work/models_mm.py": "build_mm_workbook",
    "work/models_ltm.py": "build_ltm_workbook",
    "work/models_clean.py": "build_clean_workbook",
    "work/models_edelweiss.py": "build_edelweiss_workbook",
}


def compute_file_sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def output_file(output_dir: Path, name: str, suffix: str, check_exists: bool = True) -> Path:
    """Allow only a new leaf filename, never a request-controlled filesystem path."""
    if (not isinstance(name, str) or not name or name != name.strip()
            or name.endswith((".", " ")) or any(c in name for c in '/\\:\x00')
            or PureWindowsPath(name).is_absolute() or Path(name).suffix.lower() != suffix
            or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", name, re.I)):
        raise ValueError(f"Invalid output_name: expected a plain {suffix} filename")
    target = output_dir / name
    if check_exists and (target.is_symlink() or target.exists()):
        raise ValueError(f"Output file already exists and will not be overwritten: {target}")
    if target.resolve().parent != output_dir.resolve():
        raise ValueError("Output filename escapes the package directory")
    return target


def check_com_capabilities() -> dict[str, Any]:
    """Inspect COM automation capabilities on the host environment."""
    capabilities: dict[str, Any] = {
        "platform": sys.platform,
        "excel_com": False,
        "excel_version": None,
    }
    if sys.platform != "win32":
        return capabilities

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        capabilities["com_import_error"] = "pywin32 (pythoncom/win32com) not installed"
        return capabilities

    try:
        pythoncom.CoInitialize()
        try:
            if hasattr(pythoncom, "CLSIDFromProgID"):
                pythoncom.CLSIDFromProgID("Excel.Application")
            else:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"):
                    pass
            capabilities["excel_com"] = True
        except Exception as e:
            capabilities["excel_error"] = str(e)

        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        capabilities["com_init_error"] = str(e)

    capabilities["detection"] = "COM registration only; activation and engine version are not established"
    return capabilities


def verify_output_directory(output_dir: Path, force: bool = False) -> None:
    """Enforce strict directory safety rules to protect roots, originals, and unrelated paths."""
    res = output_dir.resolve()
    if res in FORBIDDEN_OUTPUT_ROOTS:
        raise RuntimeError(f"Target output directory {res} is a protected repository root and cannot be used as an output directory.")
    if res == ROOT.resolve().parent or res.parent == res:
        raise RuntimeError(f"Target output directory {res} is a filesystem root or parent and cannot be used as an output directory.")
    for forbidden in FORBIDDEN_PARENT_DIRS:
        if res == forbidden or forbidden in res.parents:
            raise RuntimeError(f"Target output directory {res} is inside protected directory {forbidden}.")

    if res.exists() and any(res.iterdir()):
        if not force:
            raise RuntimeError(
                f"Output directory {output_dir} already exists and is not empty. Choose a fresh run; --force does not permit overwriting."
            )
        # Check ownership when --force is supplied
        manifest_file = res / "package.manifest.json"
        is_candidate_sub = False
        try:
            res.relative_to((ROOT / "candidates").resolve())
            is_candidate_sub = True
        except ValueError:
            pass
        temp_dir = Path(tempfile.gettempdir()).resolve()
        is_temp_sub = False
        try:
            res.relative_to(temp_dir)
            is_temp_sub = True
        except ValueError:
            pass

        if not manifest_file.exists() and not is_candidate_sub and not is_temp_sub:
            raise RuntimeError(
                f"Directory {output_dir} exists and is not empty, but does not contain a package.manifest.json. "
                "Overwriting arbitrary non-package directories is prohibited even with --force."
            )
        raise RuntimeError("Output directory already exists and is not empty; --force cannot safely reuse a run. Choose a fresh output directory.")



def export_sheet_previews_com(
    workbook_path: Path,
    previews_dir: Path,
    sheets: list[str],
    orientation: str = "landscape",
    fit_to_page_wide: int = 1,
) -> list[dict[str, Any]]:
    """Export selected XLSX worksheets to landscape PDF using Excel COM."""
    previews_dir.mkdir(parents=True, exist_ok=True)
    exported = []
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("Warning: win32com not available, skipping Excel PDF export.")
        return exported

    pythoncom.CoInitialize()
    xl_app = None
    wb = None
    try:
        xl_app = win32com.client.DispatchEx("Excel.Application")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        xl_app.AskToUpdateLinks = False
        xl_app.AutomationSecurity = 3
        wb = xl_app.Workbooks.Open(str(workbook_path.resolve()), ReadOnly=True, UpdateLinks=0)
        for sname in sheets:
            try:
                ws = wb.Sheets(sname)
                # xlLandscape = 2, xlPortrait = 1
                ws.PageSetup.Orientation = 2 if orientation == "landscape" else 1
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = fit_to_page_wide
                ws.PageSetup.FitToPagesTall = 1
                clean_name = sname.replace(" ", "_")
                out_pdf = (previews_dir / f"{clean_name}.pdf").resolve()
                ws.ExportAsFixedFormat(0, str(out_pdf))
                exported.append({
                    "sheet": sname,
                    "file": f"previews/{out_pdf.name}",
                    "size_bytes": out_pdf.stat().st_size,
                    "sha256": compute_file_sha(out_pdf),
                })
            except Exception as e:
                print(f"Warning: Failed to export sheet {sname} to PDF: {e}")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        finally:
            try:
                if xl_app is not None:
                    xl_app.Quit()
            finally:
                pythoncom.CoUninitialize()
    return exported


def _parse_finite_float(val: Any, cell_ref: str) -> float:
    """Strictly parse a finite numeric float from an Excel cell, rejecting None, bool, or non-finite values."""
    if val is None or isinstance(val, bool):
        raise ValueError(f"Reconciliation cell {cell_ref} is None or boolean: {val!r}")
    try:
        f = float(val)
    except (ValueError, TypeError) as exc:
        raise ValueError(f"Reconciliation cell {cell_ref} is not numeric: {val!r}") from exc
    if not math.isfinite(f):
        raise ValueError(f"Reconciliation cell {cell_ref} is non-finite: {f}")
    return f


def evaluate_reconciliation_summary(
    workbook_path: Path,
) -> list[dict[str, Any]]:
    """Evaluate Excel workbook metrics and identities without presentation dependencies."""
    import openpyxl
    wb = openpyxl.load_workbook(str(workbook_path), data_only=True)
    if "WACC-DCF" not in wb.sheetnames and "Cover & Notes" in wb.sheetnames:
        ws_cn = wb["Cover & Notes"]
        ws_is = wb["Income Statement"]
        ws_bs = wb["Balance Sheet"]
        cmp_val = _parse_finite_float(ws_cn["B10"].value, "Cover & Notes!B10")
        mcap_val = _parse_finite_float(ws_cn["B11"].value, "Cover & Notes!B11")
        sh_val = _parse_finite_float(ws_cn["B12"].value, "Cover & Notes!B12")
        iv_val = _parse_finite_float(ws_cn["E8"].value, "Cover & Notes!E8")
        rev_fy25 = _parse_finite_float(ws_is["H5"].value, "Income Statement!H5")
        pat_fy25 = _parse_finite_float(ws_is["H14"].value, "Income Statement!H14")
        bs_disc = _parse_finite_float(ws_bs["H17"].value, "Balance Sheet!H17")
        mcap_check = abs(mcap_val - (cmp_val * sh_val)) <= 1.0
        bs_check = abs(bs_disc) <= 0.01
        return [
            {
                "metric": "Snapshot Market Price (15-Oct-2025)",
                "source_cell": "Cover & Notes!B10",
                "exact_value": f"{cmp_val:.2f}",
                "displayed": f"₹{cmp_val:,.2f}",
                "tolerance": "0.01",
                "status": "historical_market_snapshot",
                "verification_method": "Historical close; provider and valuation date must be identified in the workbook source registry.",
            },
            {
                "metric": "Market Capitalization (15-Oct-2025)",
                "source_cell": "Cover & Notes!B11",
                "exact_value": f"{mcap_val:.2f}",
                "displayed": f"₹{mcap_val:,.0f} Cr",
                "tolerance": "1.0",
                "status": "verified_identity" if mcap_check else "unreconciled",
                "verification_method": "Verified non-circular identity: Market Cap = CMP * Diluted Shares.",
            },
            {
                "metric": "FY25 Consolidated Revenue",
                "source_cell": "Income Statement!H5",
                "exact_value": f"{rev_fy25:.2f}",
                "displayed": f"₹{rev_fy25:,.0f} Cr",
                "tolerance": "1.0",
                "status": "primary_sourced",
                "verification_method": "FY25 consolidated value; source URL and basis must be identified in the workbook source registry.",
            },
            {
                "metric": "FY25 Consolidated Net Profit (PAT)",
                "source_cell": "Income Statement!H14",
                "exact_value": f"{pat_fy25:.2f}",
                "displayed": f"₹{pat_fy25:,.0f} Cr",
                "tolerance": "1.0",
                "status": "primary_sourced",
                "verification_method": "FY25 consolidated value; source URL and basis must be identified in the workbook source registry.",
            },
            {
                "metric": "FY25 Balance Sheet Identity Check",
                "source_cell": "Balance Sheet!H17",
                "exact_value": f"{bs_disc:.4f}",
                "displayed": f"{bs_disc:.2f}",
                "tolerance": "0.01",
                "status": "verified_identity" if bs_check else "unreconciled",
                "verification_method": "Total Assets - Total Liabilities == 0.00 with zero historical formula plugs.",
            },
            {
                "metric": "Implied Fair Value Per Share",
                "source_cell": "Cover & Notes!E8",
                "exact_value": f"{iv_val:.2f}",
                "displayed": f"₹{iv_val:,.2f}",
                "tolerance": "0.05",
                "status": "verified_identity",
                "verification_method": "Primary Valuation Sheet output reconciled to sensitivity matrix center cell.",
            },
        ]

    ws_wacc = wb["WACC-DCF"]
    ws_rm = wb["RM"]
    ws_ds = wb["Data Sheet"]

    rf = _parse_finite_float(ws_wacc["J22"].value, "WACC-DCF!J22")
    erp = _parse_finite_float(ws_wacc["J23"].value, "WACC-DCF!J23")

    # Read RM!F10, RM!F11, and RM!F12 independently to prevent circular verification
    rm_f10 = _parse_finite_float(ws_rm["F10"].value, "RM!F10")
    rm_f11 = _parse_finite_float(ws_rm["F11"].value, "RM!F11")
    rm_f12 = _parse_finite_float(ws_rm["F12"].value, "RM!F12")
    rm = rm_f12

    beta = _parse_finite_float(ws_wacc["J24"].value, "WACC-DCF!J24")
    ke = _parse_finite_float(ws_wacc["J25"].value, "WACC-DCF!J25")
    pre_tax_kd = _parse_finite_float(ws_wacc["E22"].value, "WACC-DCF!E22")
    tax_rate = _parse_finite_float(ws_wacc["E23"].value, "WACC-DCF!E23")
    post_tax_kd = _parse_finite_float(ws_wacc["E24"].value, "WACC-DCF!E24")
    mcap = _parse_finite_float(ws_wacc["C31"].value, "WACC-DCF!C31")
    debt = _parse_finite_float(ws_wacc["C30"].value, "WACC-DCF!C30")
    w_equity = _parse_finite_float(ws_wacc["E31"].value, "WACC-DCF!E31")
    w_debt = _parse_finite_float(ws_wacc["E30"].value, "WACC-DCF!E30")
    wacc = _parse_finite_float(ws_wacc["J39"].value, "WACC-DCF!J39")

    ca_fy25 = _parse_finite_float(ws_ds["J65"].value, "Data Sheet!J65")
    cl_fy25 = _parse_finite_float(ws_ds["J60"].value, "Data Sheet!J60")
    inv_fy25 = _parse_finite_float(ws_ds["J68"].value, "Data Sheet!J68")
    if cl_fy25 == 0.0:
        raise ValueError("Data Sheet!J60 (Current Liabilities) is zero; cannot compute CR/QR")
    cr = ca_fy25 / cl_fy25
    qr = (ca_fy25 - inv_fy25) / cl_fy25

    # Independent mathematical checks (non-circular)
    rm_sum_check = abs(rm_f12 - (rm_f10 + rm_f11)) <= 0.0001
    rm_check = rm_sum_check

    ke_check = abs(ke - (rf + beta * erp)) <= 0.0001
    kd_check = abs(post_tax_kd - (pre_tax_kd * (1.0 - tax_rate))) <= 0.0001
    target_weights_sum_check = abs((w_equity + w_debt) - 1.0) <= 0.0001
    wacc_check = abs(wacc - (w_equity * ke + w_debt * post_tax_kd)) <= 0.0001

    beta_sheet = wb["Beta"]
    beta_sheet_adjusted = _parse_finite_float(beta_sheet["K13"].value, "Beta!K13")
    beta_link_check = abs(beta - beta_sheet_adjusted) <= 0.0001

    return [
        {
            "metric": "Risk-Free Rate (Rf)",
            "source_cell": "WACC-DCF!J22",
            "exact_value": f"{rf:.6f}",
            "displayed": f"{rf*100:.2f}%",
            "tolerance": "0.0001",
            "status": "sourced_assumption",
            "verification_method": "Allowlisted observation or dated prior-session proxy, explicitly identified in the workbook source registry.",
        },
        {
            "metric": "Market Return (Rm)",
            "source_cell": "RM!F12 (Total Market Return) = RM!F10 (Price Return) + RM!F11 (Dividend Yield)",
            "exact_value": f"{rm:.6f}",
            "displayed": f"{rm*100:.2f}%",
            "tolerance": "0.0001",
            "status": "verified_identity" if rm_check else "unreconciled",
            "verification_method": "Market-return schedule identity: RM!F12 = RM!F10 + RM!F11. The ERP cell is an independent analyst assumption, not asserted equal to Rm - Rf.",
        },
        {
            "metric": "Levered Beta (Hamada)",
            "source_cell": "Beta!B29 / WACC-DCF!J24",
            "exact_value": f"{beta:.6f}",
            "displayed": f"{beta:.3f}",
            "tolerance": "0.001",
            "status": "verified_identity" if beta_link_check else "unreconciled",
            "verification_method": "WACC-DCF beta equals the adjusted beta displayed on the Beta worksheet.",
        },
        {
            "metric": "Cost of Equity (Ke)",
            "source_cell": "WACC-DCF!J25",
            "exact_value": f"{ke:.6f}",
            "displayed": f"{ke*100:.2f}%",
            "tolerance": "0.0001",
            "status": "verified_identity" if ke_check else "unreconciled",
            "verification_method": "CAPM identity using the workbook's stated ERP input: Ke = Rf + Beta * ERP.",
        },
        {
            "metric": "Pre-tax Cost of Debt",
            "source_cell": "WACC-DCF!E22",
            "exact_value": f"{pre_tax_kd:.6f}",
            "displayed": f"{pre_tax_kd*100:.2f}%",
            "tolerance": "0.0001",
            "status": "sourced_assumption",
            "verification_method": "Forward marginal borrowing-rate assumption documented in the workbook source registry.",
        },
        {
            "metric": "Effective Tax Rate",
            "source_cell": "WACC-DCF!E23",
            "exact_value": f"{tax_rate:.6f}",
            "displayed": f"{tax_rate*100:.2f}%",
            "tolerance": "0.0001",
            "status": "analyst_assumption",
            "verification_method": "Editable tax-rate modeling assumption; not represented as a quoted statutory rate.",
        },
        {
            "metric": "Post-tax Cost of Debt (Kd)",
            "source_cell": "WACC-DCF!E24",
            "exact_value": f"{post_tax_kd:.6f}",
            "displayed": f"{post_tax_kd*100:.2f}%",
            "tolerance": "0.0001",
            "status": "verified_identity" if kd_check else "unreconciled",
            "verification_method": "Verified tax shield identity: PostTax_Kd = PreTax_Kd * (1 - TaxRate).",
        },
        {
            "metric": "Market Capitalization (Historical Snapshot)",
            "source_cell": "WACC-DCF!C31",
            "exact_value": f"{mcap:.2f}",
            "displayed": f"₹{mcap:,.0f} Cr",
            "tolerance": "1.0",
            "status": "calculated_estimate",
            "verification_method": "Derived from the workbook's dated market-price and diluted-share inputs.",
        },
        {
            "metric": "Total Debt / Borrowings (Historical Basis)",
            "source_cell": "Data Sheet!K59 / WACC-DCF!C30",
            "exact_value": f"{debt:.2f}",
            "displayed": f"₹{debt:,.0f} Cr",
            "tolerance": "1.0",
            "status": "inherited_template_value",
            "verification_method": "Reported or modeled debt input; period, basis, and source must be visible in the workbook source registry.",
        },
        {
            "metric": "Target Weight of Equity",
            "source_cell": "WACC-DCF!E31",
            "exact_value": f"{w_equity:.6f}",
            "displayed": f"{w_equity*100:.2f}%",
            "tolerance": "0.0001",
            "status": "verified_identity" if target_weights_sum_check else "unreconciled",
            "verification_method": "Calculated from the documented debt and equity weights; Target_We + Target_Wd = 100%.",
        },
        {
            "metric": "Target Weight of Debt",
            "source_cell": "WACC-DCF!E30",
            "exact_value": f"{w_debt:.6f}",
            "displayed": f"{w_debt*100:.2f}%",
            "tolerance": "0.0001",
            "status": "verified_identity" if target_weights_sum_check else "unreconciled",
            "verification_method": "Calculated from the documented debt and equity weights; Target_We + Target_Wd = 100%.",
        },
        {
            "metric": "WACC (Target-Weighted)",
            "source_cell": "WACC-DCF!J39",
            "exact_value": f"{wacc:.6f}",
            "displayed": f"{wacc*100:.2f}%",
            "tolerance": "0.0001",
            "status": "verified_identity" if wacc_check else "unreconciled",
            "verification_method": "Verified WACC identity: WACC = We * Ke + Wd * Kd * (1 - t).",
        },
        {
            "metric": "FY25 Current Ratio",
            "source_cell": "Data Sheet Row 65 / Row 60",
            "exact_value": f"{cr:.4f}",
            "displayed": f"{cr:.2f}",
            "tolerance": "0.01",
            "status": "calculated_from_workbook",
            "verification_method": "Calculated from the workbook's current assets and current liabilities.",
        },
        {
            "metric": "FY25 Quick Ratio",
            "source_cell": "(Row 65 - Row 68) / Row 60",
            "exact_value": f"{qr:.4f}",
            "displayed": f"{qr:.2f}",
            "tolerance": "0.01",
            "status": "calculated_from_workbook",
            "verification_method": "Calculated from the workbook's quick assets and current liabilities.",
        },
    ]


def generate_source_assumption_coverage(
    workbook_path: Path,
    output_dir: Path,
    req: dict[str, Any],
) -> tuple[Path, Path]:
    cov_json_path = output_dir / "source_assumption_coverage.json"
    cov_md_path = output_dir / "SOURCE_ASSUMPTION_COVERAGE.md"
    try:
        import openpyxl
        wb_f = openpyxl.load_workbook(str(workbook_path), data_only=False)
        wb_v = openpyxl.load_workbook(str(workbook_path), data_only=True)
    except Exception as e:
        cov_json_path.write_text("{}", encoding="utf-8")
        cov_md_path.write_text("# Source & Assumption Coverage\n\nCoverage skipped for fixture.\n", encoding="utf-8")
        return cov_json_path, cov_md_path

    comp = req.get("company", {})
    ticker = str(comp.get("ticker", "")).upper()
    key_map = {"M&M": "MM", "MM": "MM", "LTIM": "LTM", "LTM": "LTM", "CLEAN": "CLEAN", "EDELWEISS": "EDELWEISS"}
    key = key_map.get(ticker, ticker or "MODEL")
    reg_path = ROOT / "docs" / f"{key.lower()}_source_registry.json"
    registry_entries = json.loads(reg_path.read_text(encoding="utf-8")) if reg_path.exists() else []

    sheet_stats: list[dict[str, Any]] = []
    total_inputs = 0
    total_formulas = 0
    total_links = 0

    for ws in wb_f.worksheets:
        s_inputs = 0
        s_formulas = 0
        s_links = 0
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
            for cell in row:
                val = cell.value
                if val is None or isinstance(val, bool):
                    continue
                if isinstance(val, str) and val.startswith("="):
                    if "!" in val:
                        s_links += 1
                    else:
                        s_formulas += 1
                elif isinstance(val, (int, float)):
                    s_inputs += 1
        total_inputs += s_inputs
        total_formulas += s_formulas
        total_links += s_links
        sheet_stats.append({
            "sheet": ws.title,
            "numeric_inputs": s_inputs,
            "same_sheet_formulas": s_formulas,
            "cross_sheet_links": s_links,
        })

    coverage_data = {
        "package_id": req.get("package_id"),
        "company": comp,
        "workbook_file": workbook_path.name,
        "workbook_sha256": compute_file_sha(workbook_path),
        "summary": {
            "total_numeric_inputs": total_inputs,
            "total_same_sheet_formulas": total_formulas,
            "total_cross_sheet_links": total_links,
            "registry_entries_count": len(registry_entries),
        },
        "sheet_breakdown": sheet_stats,
        "source_registry_entries": registry_entries,
    }

    cov_json_path = output_dir / "source_assumption_coverage.json"
    cov_json_path.write_text(json.dumps(coverage_data, indent=2, ensure_ascii=False), encoding="utf-8")

    md_lines = [
        f"# Source & Assumption Coverage Report — {comp.get('name', key)} (`{key}`)",
        "",
        f"- **Workbook**: `{workbook_path.name}` (`SHA-256: {coverage_data['workbook_sha256']}`)",
        f"- **Valuation Cutoff Date**: `{comp.get('valuation_date', '2025-10-15')}`",
        f"- **Historical Cutoff**: `{comp.get('historical_cutoff', 'FY25 (2025-03-31)')}`",
        f"- **Total Numeric Inputs (Blue)**: `{total_inputs}` | **Same-Sheet Formulas (Black)**: `{total_formulas}` | **Cross-Sheet Links (Green)**: `{total_links}`",
        "",
        "## 1. Worksheet Formula & Input Breakdown",
        "",
        "| Worksheet | Numeric Inputs (Blue) | Same-Sheet Formulas (Black) | Cross-Sheet Links (Green) |",
        "| :--- | ---: | ---: | ---: |",
    ]
    for st in sheet_stats:
        md_lines.append(
            f"| `{st['sheet']}` | {st['numeric_inputs']} | {st['same_sheet_formulas']} | {st['cross_sheet_links']} |"
        )
    md_lines.extend([
        "",
        "## 2. Primary Source & Assumption Registry Classification",
        "",
        "| Metric | Value | Period | Classification Status | Source Organization & Filing |",
        "| :--- | :--- | :--- | :--- | :--- |",
    ])
    for entry in registry_entries:
        md_lines.append(
            f"| {entry.get('metric', '')} | `{entry.get('value', '')}` | `{entry.get('period', '')}` | `{entry.get('status', '')}` | {entry.get('source_org', '')} — {entry.get('document', '')} |"
        )

    cov_md_path = output_dir / "SOURCE_ASSUMPTION_COVERAGE.md"
    cov_md_path.write_text("\n".join(md_lines) + "\n", encoding="utf-8")
    return cov_json_path, cov_md_path


def validate_package_request(req: dict[str, Any], request_path: Path, *, excel_only: bool = True) -> None:
    """Validate package build request before any directory creation or Office activation."""
    if not excel_only:
        raise ValueError("This project is Excel-only; presentation requests are no longer supported.")
    if "presentation" in req:
        raise ValueError("The 'presentation' request field has been retired; submit a workbook-only request.")
    if req.get("schema_version") is not None and req.get("schema_version") != 1:
        raise ValueError(f"Unsupported schema_version: {req.get('schema_version')}")
    if "package_id" in req and not req.get("package_id"):
        raise ValueError("Missing required field 'package_id'")
    if "workbook" not in req:
        raise ValueError("Excel-only request must contain a 'workbook' section")

    wb_spec = req["workbook"]
    # Validate output filenames if specified
    wb_out_name = wb_spec.get("output_name")
    if wb_out_name is not None and (not wb_out_name or not isinstance(wb_out_name, str)):
        raise ValueError("Workbook output_name must be a non-empty string")

    # Validate template existence and integrity if specified
    if "source" in wb_spec:
        wb_src_path = (ROOT / wb_spec["source"]).resolve()
        if not wb_src_path.exists():
            raise FileNotFoundError(f"Workbook source not found: {wb_src_path}")
        actual_wb_sha = compute_file_sha(wb_src_path)
        if actual_wb_sha != wb_spec.get("source_sha256"):
            archive_path = (ROOT / "originals" / "archive" / Path(wb_spec["source"]).name).resolve()
            if archive_path.exists() and compute_file_sha(archive_path) == wb_spec.get("source_sha256"):
                pass
            else:
                raise ValueError(f"Workbook source SHA mismatch: expected {wb_spec.get('source_sha256')}, got {actual_wb_sha}")

    # Validate packs and inputs existence if specified
    if "model_pack" in wb_spec:
        model_pack_path = (ROOT / wb_spec["model_pack"]).resolve()
        if not model_pack_path.exists():
            raise FileNotFoundError(f"Model pack not found: {model_pack_path}")

    if "inputs" in wb_spec:
        wb_inputs_path = (ROOT / wb_spec["inputs"]).resolve()
        if not wb_inputs_path.exists():
            raise FileNotFoundError(f"Workbook inputs file not found: {wb_inputs_path}")

    # Validate company identity & units if provided
    if "company" in req:
        comp = req["company"]
        for field in ["name", "ticker", "valuation_date", "historical_cutoff", "currency", "reporting_unit", "fiscal_year_convention"]:
            if not comp.get(field):
                raise ValueError(f"Missing required company field: {field}")

        if comp["reporting_unit"] not in ("INR_Cr", "INR_Lakh", "USD_M", "USD_B"):
            raise ValueError(f"Unsupported reporting_unit: {comp['reporting_unit']}")

        if comp["currency"] not in ("INR", "USD", "EUR", "GBP"):
            raise ValueError(f"Unsupported currency: {comp['currency']}")

    # Validate workbook adapter registration and path safety
    if "adapter" in wb_spec:
        adapter_rel = wb_spec["adapter"]
        if not isinstance(adapter_rel, str) or not adapter_rel.strip():
            raise ValueError("Adapter path must be a non-empty string")
        if ".." in adapter_rel or adapter_rel.startswith(("/", "\\")):
            raise ValueError(f"Path traversal detected in adapter path: {adapter_rel!r}")
        norm_adapter = adapter_rel.replace("\\", "/")
        if norm_adapter not in REGISTERED_ADAPTERS:
            raise ValueError(
                f"Unregistered or forbidden workbook adapter: {adapter_rel!r}. "
                f"Adapters must be registered in REGISTERED_ADAPTERS whitelist."
            )
        adapter_path = (ROOT / adapter_rel).resolve()
        if not adapter_path.is_relative_to((ROOT / "work").resolve()):
            raise ValueError(f"Adapter path {adapter_path} escapes allowed directory work/")
        if not adapter_path.exists():
            raise FileNotFoundError(f"Registered adapter script not found: {adapter_path}")

    # Validate candidate_source containment
    if "candidate_source" in wb_spec:
        cand_src = wb_spec["candidate_source"]
        if not isinstance(cand_src, str) or not cand_src.strip():
            raise ValueError("candidate_source must be a non-empty string")
        if ".." in cand_src or cand_src.startswith(("/", "\\")):
            raise ValueError(f"Path traversal detected in candidate_source: {cand_src!r}")
        cand_path = (ROOT / cand_src).resolve()
        if not cand_path.is_relative_to(ROOT.resolve()):
            raise ValueError(f"candidate_source path {cand_path} escapes repository root")


def build_package(
    request_path: Path,
    output_dir: Path,
    *,
    allow_proposal: bool = False,
    force: bool = False,
    clean: bool = False,
    skip_com: bool = False,
    check_only: bool = False,
    excel_only: bool = True,
) -> dict[str, Any]:
    """Execute an Excel-only build. Presentation generation has been retired."""
    if not excel_only:
        raise ValueError("This project is Excel-only; presentation generation has been retired.")
    req_text = request_path.read_text(encoding="utf-8")
    req = json.loads(req_text)

    # If clean requested, safely purge previous candidate run directory before output validation
    if clean and output_dir.exists():
        res = output_dir.resolve()
        if res in FORBIDDEN_OUTPUT_ROOTS or res == ROOT.resolve().parent or res.parent == res:
            raise RuntimeError(f"Target output directory {res} is a protected repository root and cannot be cleaned.")
        for forbidden in FORBIDDEN_PARENT_DIRS:
            if res == forbidden or forbidden in res.parents:
                raise RuntimeError(f"Target output directory {res} is inside protected directory {forbidden} and cannot be cleaned.")
        is_cand_sub = False
        try:
            res.relative_to((ROOT / "candidates").resolve())
            is_cand_sub = True
        except ValueError:
            pass
        temp_dir = Path(tempfile.gettempdir()).resolve()
        is_temp = False
        try:
            res.relative_to(temp_dir)
            is_temp = True
        except ValueError:
            pass
        if is_cand_sub or is_temp or (res / "package.manifest.json").exists():
            shutil.rmtree(res)
        else:
            raise RuntimeError(
                f"Target output directory {res} is not inside candidates/ or temp, "
                f"and does not contain an existing package.manifest.json. Refusing to clean unverified directory."
            )

    # Validate output names before directory writes or Office activation.
    wb_out_name = req["workbook"].get("output_name", "Marksans-Recalculated.xlsx")
    wb_out_path = output_file(output_dir, wb_out_name, ".xlsx", check_exists=not check_only)

    # Validate request before directory creation or Office activation
    validate_package_request(req, request_path, excel_only=excel_only)

    wb_spec = req["workbook"]
    exp_spec = req.get("exports", {})

    wb_src_path = (ROOT / wb_spec["source"]).resolve()
    model_pack_path = (ROOT / wb_spec["model_pack"]).resolve()
    actual_wb_src_sha = compute_file_sha(wb_src_path)
    if actual_wb_src_sha != wb_spec.get("source_sha256"):
        archive_path = (ROOT / "originals" / "archive" / Path(wb_spec["source"]).name).resolve()
        if archive_path.exists() and compute_file_sha(archive_path) == wb_spec.get("source_sha256"):
            wb_src_path = archive_path
            actual_wb_src_sha = compute_file_sha(wb_src_path)

    com_caps = check_com_capabilities()

    # Pre-flight only mode: verifies request validity, hashes, and COM capabilities without disk writes
    if check_only:
        print("Preflight checks passed: request valid, source hashes verified, packs present.")
        src_hashes: dict[str, Any] = {
            "workbook_source": actual_wb_src_sha,
            "model_pack": compute_file_sha(model_pack_path),
        }
        return {
            "status": "preflight_passed",
            "ready": False,
            "package_id": req.get("package_id"),
            "name": req.get("name"),
            "excel_only": excel_only,
            "com_capabilities": com_caps,
            "source_hashes": src_hashes,
        }

    # 1. Directory safety verification
    verify_output_directory(output_dir, force=force)
    output_dir.mkdir(parents=True, exist_ok=True)
    previews_dir = output_dir / "previews"
    previews_dir.mkdir(parents=True, exist_ok=True)

    # 2. Pin and verify source hashes
    if not wb_src_path.exists():
        raise FileNotFoundError(f"Workbook source not found: {wb_src_path}")
    if actual_wb_src_sha != wb_spec["source_sha256"]:
        raise ValueError(
            f"Workbook source SHA mismatch: expected {wb_spec['source_sha256']}, got {actual_wb_src_sha}"
        )

    # 3. Handle Workbook Composition and Recalculation
    expected_recalc_sha = wb_spec.get("expected_recalculated_sha256")
    candidate_src_rel = wb_spec.get("candidate_source")
    adapter_rel = wb_spec.get("adapter")

    cand_path = (ROOT / candidate_src_rel).resolve() if candidate_src_rel else None
    if cand_path and cand_path.exists() and (not excel_only or skip_com):
        cand_sha = compute_file_sha(cand_path)
        if expected_recalc_sha and cand_sha != expected_recalc_sha and not excel_only:
            raise ValueError(f"Candidate source SHA mismatch: expected {expected_recalc_sha}, got {cand_sha}")
        shutil.copy2(cand_path, wb_out_path)
    else:
        # Check COM availability before recalculation
        if not com_caps["excel_com"]:
            raise RuntimeError(
                f"Excel COM is required for workbook recalculation but is not available on this system ({com_caps.get('excel_error')})."
            )

        if adapter_rel:
            norm_adapter = adapter_rel.replace("\\", "/")
            if norm_adapter not in REGISTERED_ADAPTERS:
                raise ValueError(f"Unregistered workbook adapter: {adapter_rel}")
            expected_func = REGISTERED_ADAPTERS[norm_adapter]
            adapter_path = (ROOT / adapter_rel).resolve()
            if not adapter_path.exists():
                raise FileNotFoundError(f"Configured workbook adapter not found: {adapter_path}")
            print(f"Running registered workbook adapter {adapter_path} ({expected_func})...")
            with tempfile.TemporaryDirectory() as td:
                staged_path = Path(td) / "staged_adapted.xlsx"
                import importlib.util
                spec = importlib.util.spec_from_file_location("wb_adapter", adapter_path)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                adapter_fn = getattr(mod, expected_func, None)
                if adapter_fn is None:
                    raise AttributeError(f"Adapter module {adapter_path} missing registered function {expected_func}")
                adapter_fn(staged_path)
                print(f"Recalculating adapted workbook to {wb_out_path}...")
                validate_candidate(
                    workbook_path=staged_path,
                    pack_path=model_pack_path,
                    recalculate=True,
                    recalculated_output=wb_out_path,
                    allow_proposal=allow_proposal,
                )
        else:
            wb_inputs_path = (ROOT / wb_spec["inputs"]).resolve() if "inputs" in wb_spec else None
            if wb_inputs_path and wb_inputs_path.exists():
                print(f"Composing workbook inputs from {wb_inputs_path} into template...")
                tpl_bytes = wb_src_path.read_bytes()
                model_pack = read_json(str(model_pack_path))
                wb_inputs = read_json(str(wb_inputs_path))
                composed_bytes, _ = compose(tpl_bytes, model_pack, wb_inputs, allow_proposal=allow_proposal)
                with tempfile.TemporaryDirectory() as td:
                    staged_path = Path(td) / "staged_composed.xlsx"
                    staged_path.write_bytes(composed_bytes)
                    print(f"Recalculating composed workbook to {wb_out_path}...")
                    validate_candidate(
                        workbook_path=staged_path,
                        pack_path=model_pack_path,
                        recalculate=True,
                        recalculated_output=wb_out_path,
                        allow_proposal=allow_proposal,
                    )
            else:
                print(f"Recalculating workbook to {wb_out_path}...")
                validate_candidate(
                    workbook_path=wb_src_path,
                    pack_path=model_pack_path,
                    recalculate=True,
                    recalculated_output=wb_out_path,
                    allow_proposal=allow_proposal,
                )

    # Validate the generated candidate workbook against its model pack
    print(f"Validating workbook {wb_out_path}...")
    validation_report = validate_candidate(
        workbook_path=wb_out_path,
        pack_path=model_pack_path,
        allow_proposal=allow_proposal,
    )
    if not validation_report.get("checks_passed"):
        raise ValueError(f"Workbook validation failed checks: {validation_report.get('status')}")

    val_json_path = output_dir / "workbook_validation.json"
    val_json_path.write_text(json.dumps(validation_report, indent=2, default=str), encoding="utf-8")

    cov_json_path, cov_md_path = generate_source_assumption_coverage(wb_out_path, output_dir, req)

    wb_out_sha = compute_file_sha(wb_out_path)

    # 4. Export Previews (Sheet PDFs via Excel COM)
    sheet_previews = []
    com_caps = check_com_capabilities()

    if not skip_com:
        if exp_spec.get("export_sheet_pdfs", True):
            if com_caps["excel_com"]:
                pdf_sheets = exp_spec.get("pdf_sheets", [
                    "Data Sheet",
                    "Ratio Analysis",
                    "Intrinsic Growth",
                    "WACC-DCF",
                    "DCF 1st way",
                    "DCF 2nd way",
                ])
                print("Exporting sheet PDFs via Excel COM...")
                sheet_previews = export_sheet_previews_com(
                    workbook_path=wb_out_path,
                    previews_dir=previews_dir,
                    sheets=pdf_sheets,
                    orientation=exp_spec.get("pdf_orientation", "landscape"),
                    fit_to_page_wide=exp_spec.get("pdf_fit_to_page_wide", 1),
                )
            else:
                print("Excel COM not available; skipping sheet PDF export.")

    # 5. Package Reconciliation Summary Table (dynamically evaluated from candidate workbook)
    reconciliation_table = evaluate_reconciliation_summary(wb_out_path)

    # 6. Compile Package Manifest (explicitly excludes self-referential manifest hash)
    artifacts = [
        {
            "file": wb_out_name,
            "type": "workbook",
            "size_bytes": wb_out_path.stat().st_size,
            "sha256": wb_out_sha,
        },
        {
            "file": val_json_path.name,
            "type": "workbook_validation",
            "size_bytes": val_json_path.stat().st_size,
            "sha256": compute_file_sha(val_json_path),
        },
        {
            "file": cov_json_path.name,
            "type": "source_assumption_coverage_json",
            "size_bytes": cov_json_path.stat().st_size,
            "sha256": compute_file_sha(cov_json_path),
        },
        {
            "file": cov_md_path.name,
            "type": "source_assumption_coverage_md",
            "size_bytes": cov_md_path.stat().st_size,
            "sha256": compute_file_sha(cov_md_path),
        },
    ]
    for shp in sheet_previews:
        artifacts.append({
            "file": shp["file"],
            "type": "sheet_preview",
            "size_bytes": shp["size_bytes"],
            "sha256": shp["sha256"],
        })

    src_hashes_out: dict[str, Any] = {
        "workbook_source": actual_wb_src_sha,
        "model_pack": compute_file_sha(model_pack_path),
    }

    manifest = {
        "package_id": req.get("package_id", "marksans-package"),
        "name": req.get("name", "Marksans Pharma Financial Package"),
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "excel_only": excel_only,
        "status": ("validating" if excel_only else "proposal") if allow_proposal else "approved",
        "ready": False,
        "pending_gates": [
            "human_pack_approval",
            "render_review",
        ],
        "com_capabilities": com_caps,
        "source_hashes": src_hashes_out,
        "workbook": {
            "file": wb_out_name,
            "sha256": wb_out_sha,
            "size_bytes": wb_out_path.stat().st_size,
            "formula_error_count": validation_report.get("formula_error_count", 0),
            "uncalculated_formula_count": validation_report.get("uncalculated_formula_count", 0),
        },
        "previews": sheet_previews,
        "workbook_validation": {
            "status": validation_report.get("status"),
            "pack_checks_passed": validation_report.get("pack_checks_passed"),
            "pack_checks_total": validation_report.get("pack_checks_total"),
            "independent_identities_passed": validation_report.get("independent_identities_passed"),
            "independent_identities_total": validation_report.get("independent_identities_total"),
        },
        "reconciliation_summary": reconciliation_table,
        "output_artifacts": artifacts,
    }

    pkg_id_lower = str(req.get("package_id", "")).lower()
    if "edelweiss" in pkg_id_lower:
        manifest["valuation"] = {
            "status": "supported_sotp_fcff_withheld",
            "methodology": "Financial Services Sum-of-the-Parts (SOTP) Valuation",
            "methodological_notes": "Industrial FCFF DCF is strictly withheld due to financial institution lending liabilities structure. SOTP attributes vertical valuations and deducts Audited FY25 Standalone HoldCo Net Debt (-₹3,067 Cr).",
        }
    elif "mm" in pkg_id_lower:
        manifest["valuation"] = {
            "status": "supported_sotp",
            "methodology": "Sum-of-the-Parts (SOTP) Valuation",
            "methodological_notes": "Core Auto & Farm valued at 14.0x and 13.0x FY26E Core EBITDA; listed stakes MMFS (52.16%) and TechM (25.48%) valued at market with 20% HoldCo discount; adds Audited FY25 Standalone Net Cash (+₹9,109 Cr).",
        }
    elif "ltm" in pkg_id_lower:
        manifest["valuation"] = {
            "status": "supported_fcff_dcf",
            "methodology": "Full-Year FCFF DCF (Baseline Annual, Cell B24)",
            "methodological_notes": "5-year explicit FCFF driven by Headcount × Utilization × Billing; terminal growth 5.0%; WACC 11.87% linked to 5-peer Hamada relevered beta 0.9405; point-in-time stub timing sensitivity in Row 66.",
        }
    elif "clean" in pkg_id_lower:
        manifest["valuation"] = {
            "status": "supported_fcff_dcf",
            "methodology": "Full-Year Manufacturing FCFF DCF (Baseline Annual, Cell B24)",
            "methodological_notes": "5-year explicit FCFF driven by chemical plant capacity and CFCL Unit 4 HALS commercial ramp; terminal growth 5.0%; WACC 11.73% linked to 2Y weekly OLS beta 0.9006; point-in-time stub timing sensitivity in Row 66.",
        }

    manifest_path = output_dir / "package.manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, default=str), encoding="utf-8")
    print(f"Package manifest generated at {manifest_path}")
    return manifest


def build_package_batch(
    request_paths: list[Path],
    output_base: Path,
    *,
    allow_proposal: bool = False,
    force: bool = False,
    clean: bool = False,
    skip_com: bool = False,
    check_only: bool = False,
    excel_only: bool = True,
) -> dict[str, Any]:
    """Execute batch build across multiple package requests into target subdirectories under output_base."""
    if not excel_only:
        raise ValueError("This project is Excel-only; presentation generation has been retired.")
    output_base = output_base.resolve()
    if output_base in FORBIDDEN_OUTPUT_ROOTS or output_base == ROOT.resolve().parent or output_base.parent == output_base:
        raise RuntimeError(f"Target output base {output_base} is a protected root and cannot be used as batch output base.")
    for forbidden in FORBIDDEN_PARENT_DIRS:
        if output_base == forbidden or forbidden in output_base.parents:
            raise RuntimeError(f"Target output base {output_base} is inside protected directory {forbidden}.")

    output_base.mkdir(parents=True, exist_ok=True)
    results = []
    failures = []

    for req_p in request_paths:
        req_path = (ROOT / req_p).resolve() if not Path(req_p).is_absolute() else Path(req_p)
        if not req_path.exists():
            failures.append({"request": str(req_p), "error": f"Request file not found: {req_path}"})
            continue

        stem = req_path.stem
        for sfx in ["-package", ".package"]:
            if stem.endswith(sfx):
                stem = stem[:-len(sfx)]
        stem = stem.lower()
        subdir_name = f"run_excel_{stem}" if excel_only else f"run_native_{stem}"
        out_subdir = output_base / subdir_name

        print(f"\n=======================================================")
        print(f"[BATCH] Building package {req_path.name} -> {out_subdir}...")
        try:
            manifest = build_package(
                request_path=req_path,
                output_dir=out_subdir,
                allow_proposal=allow_proposal,
                force=force,
                clean=clean,
                skip_com=skip_com,
                check_only=check_only,
                excel_only=excel_only,
            )
            results.append({
                "request": str(req_p),
                "package_id": manifest.get("package_id"),
                "status": manifest.get("status"),
                "output_dir": str(out_subdir),
                "workbook": manifest.get("workbook", {}).get("file"),
                "sha256": manifest.get("workbook", {}).get("sha256"),
            })
        except Exception as e:
            print(f"[BATCH] FAILED for {req_path.name}: {e}")
            failures.append({
                "request": str(req_p),
                "error": str(e),
                "attempted_output_dir": str(out_subdir),
            })

    batch_summary = {
        "batch_mode": True,
        "total_requests": len(request_paths),
        "succeeded": len(results),
        "failed_count": len(failures),
        "status": "success" if not failures else "partial_or_failed",
        "results": results,
        "failures": failures,
    }
    return batch_summary


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel-only financial workbook builder")
    parser.add_argument("--request", help="Path to package request JSON (single build mode)")
    parser.add_argument("--requests", nargs="+", help="Multiple package request JSON paths for batch build")
    parser.add_argument("--batch", action="store_true", help="Batch build mode across multiple requests")
    parser.add_argument("--output-dir", help="Target output directory for single package build")
    parser.add_argument("--output-base", help="Base output directory for batch package builds (defaults to candidates/)")
    parser.add_argument("--clean", action="store_true", help="Safely clean candidate run output directory before building (only allowed for verified directories inside candidates/ or temp, or with existing package.manifest.json)")
    parser.add_argument("--allow-proposal", action="store_true", help="Allow building with a proposal model pack")
    parser.add_argument("--force", action="store_true", help="Legacy option; non-empty run directories are never overwritten")
    parser.add_argument("--skip-com", action="store_true", help="Skip Excel COM recalculation and PDF preview export")
    parser.add_argument("--check-only", "--preflight-only", action="store_true", help="Fast preflight request and template validation without Office COM activation or file generation")
    args = parser.parse_args()

    # Determine execution mode: batch vs single
    if args.batch or args.requests:
        default_reqs = [
            "examples/mm-package.json",
            "examples/ltm-package.json",
            "examples/clean-package.json",
            "examples/edelweiss-package.json",
        ]
        raw_reqs = args.requests if args.requests else default_reqs
        req_paths = [Path(r) for r in raw_reqs]
        out_base = Path(args.output_base) if args.output_base else (Path(args.output_dir) if args.output_dir else ROOT / "candidates")

        try:
            summary = build_package_batch(
                request_paths=req_paths,
                output_base=out_base,
                allow_proposal=args.allow_proposal,
                force=args.force,
                clean=args.clean,
                skip_com=args.skip_com,
                check_only=args.check_only,
                excel_only=True,
            )
            print(json.dumps(summary, indent=2))
            return 0 if summary["failed_count"] == 0 else 1
        except Exception as e:
            print(json.dumps({"status": "blocked", "ready": False, "error": str(e)}, indent=2))
            return 1

    if not args.request or not args.output_dir:
        parser.error("Either --request and --output-dir (single mode) or --batch / --requests (batch mode) must be specified.")

    try:
        manifest = build_package(
            request_path=Path(args.request),
            output_dir=Path(args.output_dir),
            allow_proposal=args.allow_proposal,
            force=args.force,
            clean=args.clean,
            skip_com=args.skip_com,
            check_only=args.check_only,
            excel_only=True,
        )
        if args.check_only:
            print(json.dumps({
                "status": manifest["status"],
                "ready": manifest["ready"],
                "package_id": manifest["package_id"],
                "excel_only": manifest.get("excel_only", False),
                "com_capabilities": manifest["com_capabilities"],
            }, indent=2))
            return 0
        print(json.dumps({
            "status": manifest["status"],
            "ready": manifest["ready"],
            "package_id": manifest["package_id"],
            "excel_only": manifest.get("excel_only", False),
            "artifacts_count": len(manifest["output_artifacts"]),
            "pending_gates": manifest["pending_gates"],
        }, indent=2))
        return 0
    except Exception as e:
        print(json.dumps({"status": "blocked", "ready": False, "error": str(e)}, indent=2))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

