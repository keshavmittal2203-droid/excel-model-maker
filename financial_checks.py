"""Independent financial checks against a RECALCULATED workbook.

Every check recomputes an accounting/valuation identity in Python from the
workbook's cached values and compares it with the workbook's own cached
result. The recomputation is independent of modelmaker's edit machinery;
it trusts only cell values. Files that were never recalculated by a real
spreadsheet engine must not be fed to this tool (see QUALITY.md).

Exit code 0 means every check passed within tolerance; 1 means at least one
check failed or a cached error cell was found.
"""
from __future__ import annotations

import argparse
import json
from decimal import Decimal
import sys
from pathlib import Path

from modelmaker import ModelError, NS, Package

TOL = Decimal("1e-9")


def value(package: Package, sheet: str, ref: str):
    tree = package.sheets[sheet][1]
    for cell in tree.iter(f"{{{NS}}}c"):
        if cell.get("r") != ref:
            continue
        if cell.get("t") == "e":
            raise ModelError(f"cached error at {sheet}!{ref}: {cell.findtext(f'{{{NS}}}v')}")
        cached = cell.findtext(f"{{{NS}}}v")
        if cached is None:
            return None
        return Decimal(cached)
    return None


def check(name, computed, expected, tolerance=TOL):
    if computed is None or expected is None:
        return {"name": name, "pass": False, "detail": "missing value"}
    diff = abs(computed - expected)
    rel = diff / max(abs(expected), Decimal("1e-12"))
    return {"name": name, "pass": diff <= tolerance or rel <= 1e-9,
            "computed": str(computed), "expected": str(expected), "abs_diff": str(diff)}


def cached_errors(package: Package):
    found = []
    for name, (_, tree) in package.sheets.items():
        for cell in tree.iter(f"{{{NS}}}c"):
            if cell.get("t") == "e":
                found.append(f"{name}!{cell.get('r')}={cell.findtext(f'{{{NS}}}v')}")
    return found


def run_checks(package: Package) -> dict:
    results = []
    v = lambda sheet, ref: value(package, sheet, ref)  # noqa: E731

    # ---- WACC chain (WACC-DCF) ----
    results.append(check("wacc_supproduct_identity",
                         v("WACC-DCF", "J39"),
                         v("WACC-DCF", "I37") * v("WACC-DCF", "J37") + v("WACC-DCF", "I38") * v("WACC-DCF", "J38")))
    results.append(check("wacc_weights_sum_to_one",
                         v("WACC-DCF", "E30") + v("WACC-DCF", "E31"), Decimal(1)))
    results.append(check("capm_identity",
                         v("WACC-DCF", "J25"),
                         v("WACC-DCF", "J22") + v("WACC-DCF", "J24") * v("WACC-DCF", "J23")))
    results.append(check("levered_beta_hamada",
                         v("WACC-DCF", "J24"),
                         v("WACC-DCF", "J29") * (1 + (1 - v("WACC-DCF", "J31")) * v("WACC-DCF", "J30"))))

    # ---- DCF 2nd way ----
    wacc2 = v("DCF 2nd way", "D9")
    tgr = v("DCF 2nd way", "D10")
    results.append(check("dcf2_terminal_value",
                         v("DCF 2nd way", "N56"), v("DCF 2nd way", "N53") * (1 + tgr) / (wacc2 - tgr), Decimal("1e-6")))
    results.append(check("dcf2_tv_discount",
                         v("DCF 2nd way", "N57"), v("DCF 2nd way", "N56") / (1 + wacc2) ** v("DCF 2nd way", "I11"), Decimal("1e-6")))
    pv_sum = v("DCF 2nd way", "N54") + v("DCF 2nd way", "M54") + v("DCF 2nd way", "L54") + v("DCF 2nd way", "K54") + v("DCF 2nd way", "J54")
    results.append(check("dcf2_explicit_pv_sum",
                         v("DCF 2nd way", "N58"), pv_sum + v("DCF 2nd way", "N57"), Decimal("1e-6")))
    for ref, exponent in (("J54", "J31"), ("K54", "K31"), ("L54", "L31"), ("M54", "M31"), ("N54", "N31")):
        results.append(check(f"dcf2_pv_factor_{ref}",
                             v("DCF 2nd way", ref), v("DCF 2nd way", ref[:-1] + "3") / (1 + wacc2) ** v("DCF 2nd way", exponent), Decimal("1e-6")))
    results.append(check("dcf2_equity_bridge",
                         v("DCF 2nd way", "N61"),
                         v("DCF 2nd way", "N58") + v("DCF 2nd way", "N59") - v("DCF 2nd way", "N60"), Decimal("1e-6")))
    results.append(check("dcf2_per_share",
                         v("DCF 2nd way", "N63"), v("DCF 2nd way", "N61") / v("DCF 2nd way", "N62"), Decimal("1e-6")))

    # ---- DCF 1st way ----
    results.append(check("dcf1_tv_numerator",
                         v("DCF 1st way", "E25"), v("DCF 1st way", "K13") * (1 + v("DCF 1st way", "E27")), Decimal("1e-6")))
    results.append(check("dcf1_terminal_value",
                         v("DCF 1st way", "E29"), v("DCF 1st way", "E25") / (v("DCF 1st way", "E26") - v("DCF 1st way", "E27")), Decimal("1e-6")))
    explicit = sum(v("DCF 1st way", f"{col}17") for col in "GHIJK")
    results.append(check("dcf1_explicit_pv_sum",
                         v("DCF 1st way", "E34"), explicit, Decimal("1e-6")))
    results.append(check("dcf1_enterprise_value",
                         v("DCF 1st way", "E36"), v("DCF 1st way", "E34") + v("DCF 1st way", "E35"), Decimal("1e-6")))
    results.append(check("dcf1_tv_pv",
                         v("DCF 1st way", "E35"), v("DCF 1st way", "E29") * v("DCF 1st way", "K15"), Decimal("1e-6")))
    results.append(check("dcf1_equity_bridge",
                         v("DCF 1st way", "E40"),
                         v("DCF 1st way", "E36") + v("DCF 1st way", "E38") - v("DCF 1st way", "E39"), Decimal("1e-6")))
    results.append(check("dcf1_per_share",
                         v("DCF 1st way", "E42"), v("DCF 1st way", "E40") / v("DCF 1st way", "E41"), Decimal("1e-6")))
    # data table: cell at current WACC row and 5% growth column equals computed EV
    results.append(check("dcf1_datatable_matches_ev",
                         v("DCF 1st way", "J27"), v("DCF 1st way", "E36"), Decimal("1e-4")))

    # ---- cross-sheet reconciliations ----
    data_rev = [v("Data Sheet", f"{col}17") for col in "GHIJK"]
    mf_rev = [v("Main Financials", f"{col}11") for col in "FGHIJ"]
    results.append(check("revenue_reconciliation_data_sheet_to_main_financials",
                         sum((a - b) ** 2 for a, b in zip(data_rev, mf_rev)), Decimal(0), Decimal("1e-6")))
    ig_reinvest = [v("Intrinsic Growth", f"{col}43") for col in "EFGHI"]
    cf_reinvest = [-v("Cash Flow", f"{c}15") - v("Cash Flow", f"{c}16") - v("Main Financials", f"{m}29")
                   for c, m in zip("JKLMN", "FGHIJ")]
    results.append(check("reinvestment_reconciliation_intrinsic_growth",
                         sum((a - b) ** 2 for a, b in zip(ig_reinvest, cf_reinvest)), Decimal(0), Decimal("1e-6")))

    errors = cached_errors(package)
    passed = all(item["pass"] for item in results) and not errors
    return {"passed": passed, "checks": results, "cached_error_cells": errors}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", help="recalculated XLSX path")
    parser.add_argument("--report", help="optional JSON report output path")
    args = parser.parse_args()
    try:
        package = Package(Path(args.workbook).read_bytes())
        report = run_checks(package)
        report["workbook_sha256"] = package.summary()["sha256"]
        text = json.dumps(report, indent=2)
        print(text)
        if args.report:
            Path(args.report).write_text(text, encoding="utf-8")
        return 0 if report["passed"] else 1
    except (ModelError, OSError, KeyError, TypeError) as exc:
        print(json.dumps({"passed": False, "error": str(exc)}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
