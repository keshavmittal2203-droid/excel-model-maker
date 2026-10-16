"""Recalculation and independent financial validation runner for XLSX candidates.

Never self-asserts ready=True: human pack approval and visual render review remain mandatory release gates.
"""
from __future__ import annotations

import argparse
import ast
from copy import copy, deepcopy
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
import io
import json
from pathlib import Path
import re
import subprocess
import tempfile
from typing import Any
import uuid
import xml.etree.ElementTree as ET
import zipfile

from modelmaker import (
    ModelError,
    NS,
    Package,
    cell_range,
    compose,
    numeric,
    read_json,
    sha,
    validate_pack,
)

EXCEL_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#GETTING_DATA"}
DEFAULT_TOL = Decimal("1e-6")
CELL_REF_RE = re.compile(r"((?:'[^']+'|[A-Za-z0-9_.-]+)!)?(\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})\b")
FUNC_CALL_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


def get_cell_decimal(package: Package, sheet: str, ref: str) -> Decimal | None:
    if sheet not in package.sheets:
        raise ModelError(f"Missing worksheet: {sheet}")
    tree = package.sheets[sheet][1]
    for cell in tree.iter(f"{{{NS}}}c"):
        if cell.get("r") != ref:
            continue
        if cell.get("t") == "e":
            raise ModelError(f"Cached error at {sheet}!{ref}: {cell.findtext(f'{{{NS}}}v')}")
        raw = cell.findtext(f"{{{NS}}}v")
        if raw is None or raw.strip() == "":
            return None
        if raw.strip().upper() in EXCEL_ERRORS:
            raise ModelError(f"Error string at {sheet}!{ref}: {raw}")
        try:
            return Decimal(raw.strip())
        except InvalidOperation:
            return None
    return None


def scan_formula_errors_and_missing_caches(
    package: Package,
    excluded_sheets: set[str] | None = None,
) -> tuple[list[dict[str, str]], list[str]]:
    """Return (error_cells, uncalculated_formula_cells) across all worksheets and formula ranges."""
    errors: list[dict[str, str]] = []
    missing_caches: list[str] = []
    seen_missing: set[str] = set()
    excluded = excluded_sheets or set()

    for sheet_name, (_, tree) in package.sheets.items():
        if sheet_name in excluded:
            continue
        followers = package.formula_followers.get(sheet_name, set())
        present_cells: dict[str, ET.Element] = {}
        for cell in tree.iter(f"{{{NS}}}c"):
            ref = cell.get("r") or ""
            present_cells[ref] = cell
            t = cell.get("t")
            v_elem = cell.find(f"{{{NS}}}v")
            val = (v_elem.text or "").strip() if v_elem is not None else ""
            has_formula = cell.find(f"{{{NS}}}f") is not None or ref in followers
            if t == "e" or (v_elem is not None and val.upper() in EXCEL_ERRORS):
                errors.append({"sheet": sheet_name, "cell": ref, "error": val or "error"})
            elif has_formula and v_elem is None:
                coord = f"{sheet_name}!{ref}"
                if coord not in seen_missing:
                    seen_missing.add(coord)
                    missing_caches.append(coord)

        # Array/dataTable follower cells must physically exist with a cached value.
        # Shared-formula ref ranges may legitimately contain blank, non-formula gap cells.
        for f_ref in sorted(followers):
            if f_ref not in present_cells:
                coord = f"{sheet_name}!{f_ref}"
                if coord not in seen_missing:
                    seen_missing.add(coord)
                    missing_caches.append(coord)

    return errors, missing_caches


def _compare_metric(
    name: str,
    computed: Decimal | None,
    expected: Decimal | None,
    abs_tolerance: Decimal = DEFAULT_TOL,
    rel_tolerance: Decimal | None = None,
) -> dict[str, Any]:
    if computed is None or expected is None:
        return {
            "name": name,
            "passed": False,
            "computed": str(computed) if computed is not None else None,
            "expected": str(expected) if expected is not None else None,
            "abs_diff": None,
            "abs_tolerance": str(abs_tolerance),
            "detail": "missing cached value",
        }
    diff = abs(computed - expected)
    passed = bool(diff <= abs_tolerance)
    result: dict[str, Any] = {
        "name": name,
        "passed": passed,
        "computed": str(computed),
        "expected": str(expected),
        "abs_diff": str(diff),
        "abs_tolerance": str(abs_tolerance),
    }
    if rel_tolerance is not None:
        denom = max(abs(expected), Decimal("1e-12"))
        rel = diff / denom
        # Explicit rel_tolerance is an additional constraint, never a fallback that bypasses abs_tolerance.
        result["passed"] = bool(passed and rel <= rel_tolerance)
        result["rel_diff"] = str(rel)
        result["rel_tolerance"] = str(rel_tolerance)
    return result


def evaluate_pack_checks(package: Package, pack: dict) -> list[dict[str, Any]]:
    results = []
    for chk in pack.get("checks", []):
        name = chk["name"]
        sheet = chk["sheet"]
        cell = chk["cell"]
        expected = numeric(chk["expected"])
        tol = numeric(chk["abs_tolerance"])
        rel_tol = numeric(chk["rel_tolerance"]) if "rel_tolerance" in chk else None
        try:
            actual = get_cell_decimal(package, sheet, cell)
        except ModelError as exc:
            results.append({
                "name": name,
                "sheet": sheet,
                "cell": cell,
                "passed": False,
                "computed": None,
                "expected": str(expected),
                "abs_diff": None,
                "abs_tolerance": str(tol),
                "detail": str(exc),
            })
            continue
        item = _compare_metric(name, actual, expected, abs_tolerance=tol, rel_tolerance=rel_tol)
        item["sheet"] = sheet
        item["cell"] = cell
        results.append(item)
    return results


def is_equity_research_workbook(package: Package) -> bool:
    required = {
        "WACC-DCF",
        "Intrinsic Growth",
        "DCF 1st way",
        "DCF 2nd way",
        "Data Sheet",
        "BS",
        "Cash Flow",
        "Main Financials",
        "compdata",
    }
    return required.issubset(set(package.sheets))


def run_equity_research_identities(package: Package, check_datatable: bool = True) -> list[dict[str, Any]]:
    """Evaluate 29 independent Python-recomputed accounting and valuation identities."""
    results: list[dict[str, Any]] = []
    v = lambda s, r: get_cell_decimal(package, s, r)  # noqa: E731

    try:
        # 1. WACC-DCF identities
        results.append(_compare_metric(
            "wacc_after_tax_cost_of_debt",
            v("WACC-DCF", "E24"),
            v("WACC-DCF", "E22") * (Decimal(1) - v("WACC-DCF", "E23")),
            Decimal("1e-9"),
        ))
        results.append(_compare_metric(
            "wacc_weights_sum_to_one",
            v("WACC-DCF", "E30") + v("WACC-DCF", "E31"),
            Decimal(1),
            Decimal("1e-9"),
        ))
        results.append(_compare_metric(
            "wacc_levered_beta_hamada",
            v("WACC-DCF", "J24"),
            v("WACC-DCF", "J29") * (Decimal(1) + (Decimal(1) - v("WACC-DCF", "J31")) * v("WACC-DCF", "J30")),
            Decimal("1e-9"),
        ))
        results.append(_compare_metric(
            "wacc_capm_cost_of_equity",
            v("WACC-DCF", "J25"),
            v("WACC-DCF", "J22") + v("WACC-DCF", "J24") * v("WACC-DCF", "J23"),
            Decimal("1e-9"),
        ))
        results.append(_compare_metric(
            "wacc_sumproduct_identity",
            v("WACC-DCF", "J39"),
            v("WACC-DCF", "I37") * v("WACC-DCF", "J37") + v("WACC-DCF", "I38") * v("WACC-DCF", "J38"),
            Decimal("1e-9"),
        ))

        # 2. DCF 1st way identities
        results.append(_compare_metric(
            "dcf1_tv_numerator",
            v("DCF 1st way", "E25"),
            v("DCF 1st way", "K13") * (Decimal(1) + v("DCF 1st way", "E27")),
            Decimal("1e-6"),
        ))
        results.append(_compare_metric(
            "dcf1_terminal_value",
            v("DCF 1st way", "E29"),
            v("DCF 1st way", "E25") / (v("DCF 1st way", "E26") - v("DCF 1st way", "E27")),
            Decimal("1e-5"),
        ))
        explicit_pv1 = sum(v("DCF 1st way", f"{c}17") for c in "GHIJK")
        results.append(_compare_metric(
            "dcf1_explicit_pv_sum",
            v("DCF 1st way", "E34"),
            explicit_pv1,
            Decimal("1e-6"),
        ))
        results.append(_compare_metric(
            "dcf1_tv_pv",
            v("DCF 1st way", "E35"),
            v("DCF 1st way", "E29") * v("DCF 1st way", "K15"),
            Decimal("1e-5"),
        ))
        results.append(_compare_metric(
            "dcf1_enterprise_value",
            v("DCF 1st way", "E36"),
            v("DCF 1st way", "E34") + v("DCF 1st way", "E35"),
            Decimal("1e-6"),
        ))
        results.append(_compare_metric(
            "dcf1_equity_bridge",
            v("DCF 1st way", "E40"),
            v("DCF 1st way", "E36") + v("DCF 1st way", "E38") - v("DCF 1st way", "E39"),
            Decimal("1e-6"),
        ))
        results.append(_compare_metric(
            "dcf1_value_per_share",
            v("DCF 1st way", "E42"),
            v("DCF 1st way", "E40") / v("DCF 1st way", "E41"),
            Decimal("1e-6"),
        ))
        if check_datatable:
            if abs(v("DCF 1st way", "G27") - v("DCF 1st way", "E26")) <= Decimal("1e-6") and abs(
                v("DCF 1st way", "J24") - v("DCF 1st way", "E27")
            ) <= Decimal("1e-6"):
                results.append(_compare_metric(
                    "dcf1_datatable_center_matches_ev",
                    v("DCF 1st way", "J27"),
                    v("DCF 1st way", "E36"),
                    Decimal("1e-4"),
                ))

        # 3. DCF 2nd way identities
        wacc2 = v("DCF 2nd way", "D9")
        tgr2 = v("DCF 2nd way", "D10")
        wacc2_f = float(wacc2)
        results.append(_compare_metric(
            "dcf2_terminal_value",
            v("DCF 2nd way", "N56"),
            v("DCF 2nd way", "N53") * (Decimal(1) + tgr2) / (wacc2 - tgr2),
            Decimal("1e-5"),
        ))
        exp_tv = float(v("DCF 2nd way", "I11"))
        results.append(_compare_metric(
            "dcf2_tv_pv",
            v("DCF 2nd way", "N57"),
            Decimal(str(float(v("DCF 2nd way", "N56")) / ((1.0 + wacc2_f) ** exp_tv))),
            Decimal("1e-5"),
        ))
        for col in "JKLMN":
            fcf_c = float(v("DCF 2nd way", f"{col}53"))
            disc_p = float(v("DCF 2nd way", f"{col}31"))
            results.append(_compare_metric(
                f"dcf2_pv_fcf_{col}54",
                v("DCF 2nd way", f"{col}54"),
                Decimal(str(fcf_c / ((1.0 + wacc2_f) ** disc_p))),
                Decimal("1e-5"),
            ))
        pv_sum2 = sum(v("DCF 2nd way", f"{col}54") for col in "JKLMN")
        results.append(_compare_metric(
            "dcf2_enterprise_value",
            v("DCF 2nd way", "N58"),
            pv_sum2 + v("DCF 2nd way", "N57"),
            Decimal("1e-5"),
        ))
        results.append(_compare_metric(
            "dcf2_equity_bridge",
            v("DCF 2nd way", "N61"),
            v("DCF 2nd way", "N58") + v("DCF 2nd way", "N59") - v("DCF 2nd way", "N60"),
            Decimal("1e-5"),
        ))
        results.append(_compare_metric(
            "dcf2_value_per_share",
            v("DCF 2nd way", "N63"),
            v("DCF 2nd way", "N61") / v("DCF 2nd way", "N62"),
            Decimal("1e-5"),
        ))

        # 4. Cross-sheet accounting & reconciliation checks
        ds_bs_diff = sum(
            abs(v("Data Sheet", f"{col}61") - v("Data Sheet", f"{col}66"))
            for col in "BCDEFGHIJK"
        )
        results.append(_compare_metric(
            "datasheet_bs_total_assets_eq_liabilities",
            ds_bs_diff,
            Decimal(0),
            Decimal("1e-5"),
        ))
        mf_bs_diff = sum(
            abs(v("Main Financials", f"{col}70") - v("Main Financials", f"{col}57"))
            for col in "DEFGHIJ"
        )
        results.append(_compare_metric(
            "main_financials_bs_balance_check",
            mf_bs_diff,
            Decimal(0),
            Decimal("1e-5"),
        ))
        rev_diff = sum(
            abs(v("Data Sheet", f"{dcol}17") - v("Main Financials", f"{mcol}11"))
            for dcol, mcol in zip("GHIJK", "FGHIJ")
        )
        results.append(_compare_metric(
            "revenue_tie_datasheet_to_main_financials",
            rev_diff,
            Decimal(0),
            Decimal("1e-6"),
        ))
        reinv_diff = sum(
            abs(
                v("Intrinsic Growth", f"{icol}43")
                - (-v("Cash Flow", f"{ccol}15") - (v("Cash Flow", f"{ccol}16") or Decimal(0)) - v("Main Financials", f"{mcol}29"))
            )
            for icol, ccol, mcol in zip("EFGHI", "JKLMN", "FGHIJ")
        )
        results.append(_compare_metric(
            "reinvestment_tie_cash_flow_to_intrinsic_growth",
            reinv_diff,
            Decimal(0),
            Decimal("1e-5"),
        ))
        ig_growth_diff = sum(
            abs(v("Intrinsic Growth", f"{col}62") - (v("Intrinsic Growth", f"{col}59") * v("Intrinsic Growth", f"{col}60")))
            for col in "FGHI"
        )
        results.append(_compare_metric(
            "intrinsic_growth_roic_reinvestment_identity",
            ig_growth_diff,
            Decimal(0),
            Decimal("1e-6"),
        ))
        comps_ev_diff = sum(
            abs(v("compdata", f"J{row}") - (v("compdata", f"F{row}") + v("compdata", f"G{row}") - v("compdata", f"H{row}")))
            for row in range(3, 11)
        )
        results.append(_compare_metric(
            "compdata_enterprise_value_identity",
            comps_ev_diff,
            Decimal(0),
            Decimal("1e-5"),
        ))
    except (ModelError, TypeError, ZeroDivisionError) as exc:
        results.append({
            "name": "equity_research_identities_execution",
            "passed": False,
            "computed": None,
            "expected": "all identities evaluable",
            "abs_diff": None,
            "abs_tolerance": str(DEFAULT_TOL),
            "detail": str(exc),
        })
    return results


def recalculate_excel_com(input_xlsx: Path, output_xlsx: Path, timeout_sec: int = 120) -> dict[str, Any]:
    """Recalculate an XLSX workbook using installed Microsoft Excel COM on Windows.

    Refuses to overwrite input_xlsx or any pre-existing output_xlsx file.
    """
    in_resolved = input_xlsx.resolve()
    out_resolved = output_xlsx.resolve()
    if in_resolved == out_resolved:
        raise ModelError("Recalculated output path must differ from input path")
    if out_resolved.exists():
        raise ModelError(f"Recalculated output file already exists and will not be overwritten: {out_resolved}")
    out_resolved.parent.mkdir(parents=True, exist_ok=True)

    in_escaped = str(in_resolved).replace("'", "''")
    out_escaped = str(out_resolved).replace("'", "''")
    ps_script = (
        "$ErrorActionPreference = 'Stop'\n"
        "$excel = $null\n"
        "$wb = $null\n"
        "try {\n"
        "  $excel = New-Object -ComObject Excel.Application\n"
        "  $excel.Visible = $false\n"
        "  $excel.DisplayAlerts = $false\n"
        "  $excel.AskToUpdateLinks = $false\n"
        "  $excel.AutomationSecurity = 3\n"
        f"  $wb = $excel.Workbooks.Open('{in_escaped}', 0, $true)\n"
        "  $excel.CalculateFullRebuild()\n"
        f"  $wb.SaveAs('{out_escaped}', 51)\n"
        "  $wb.Close($false)\n"
        "  $wb = $null\n"
        "  Write-Output (\"EXCEL_COM_OK version=\" + $excel.Version)\n"
        "} finally {\n"
        "  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }\n"
        "  if ($excel -ne $null) { try { $excel.Quit() } catch {} }\n"
        "}\n"
    )
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        capture_output=True,
        text=True,
        timeout=timeout_sec,
        check=False,
    )
    if proc.returncode != 0 or not out_resolved.exists():
        raise ModelError(f"Excel COM recalculation failed: {(proc.stderr or proc.stdout).strip()}")
    version_match = re.search(r"EXCEL_COM_OK version=(\S+)", proc.stdout)
    excel_version = version_match.group(1) if version_match else "unknown"
    return {
        "engine": "Microsoft Excel COM",
        "engine_version": excel_version,
        "synthetic_only": False,
        "input_sha256": sha(in_resolved.read_bytes()),
        "output_sha256": sha(out_resolved.read_bytes()),
        "output_path": str(out_resolved),
    }


def _eval_decimal_ast(node: ast.AST) -> Decimal:
    if isinstance(node, ast.Expression):
        return _eval_decimal_ast(node.body)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return Decimal(str(node.value))
    if isinstance(node, ast.UnaryOp):
        operand = _eval_decimal_ast(node.operand)
        if isinstance(node.op, ast.UAdd):
            return +operand
        if isinstance(node.op, ast.USub):
            return -operand
    if isinstance(node, ast.BinOp):
        left = _eval_decimal_ast(node.left)
        right = _eval_decimal_ast(node.right)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            if right == 0:
                raise ModelError("Division by zero in Python formula evaluator")
            return left / right
        if isinstance(node.op, ast.Pow):
            return Decimal(str(float(left) ** float(right)))
    raise ModelError(f"Unsupported syntax in Python formula evaluator: {ast.dump(node)}")


def recalculate_package_python(data: bytes) -> bytes:
    """Strict, partial Python formula evaluator for basic synthetic test fixtures only.

    Explicitly rejects special formulas (shared, array, dataTable), workbook blockers,
    unsupported functions, and missing/uncalculated cell references. Never substitutes 0
    for missing cells and never certifies a real workbook.
    """
    pkg = Package(data)
    if pkg.blockers():
        raise ModelError(f"Workbook has blockers unsupported by Python evaluator: {'; '.join(pkg.blockers())}")
    if pkg.special_formulas:
        raise ModelError("Python evaluator rejects workbooks with shared, array, or dataTable formulas; use Excel COM")

    values: dict[tuple[str, str], Decimal] = {}
    formulas: dict[tuple[str, str], str] = {}

    for sname, (_, tree) in pkg.sheets.items():
        for c in tree.iter(f"{{{NS}}}c"):
            ref = c.get("r") or ""
            f = c.find(f"{{{NS}}}f")
            if f is not None:
                if f.get("t", "normal") != "normal" or not (f.text or "").strip():
                    raise ModelError(f"Unsupported or empty formula in Python evaluator at {sname}!{ref}")
                formulas[(sname, ref)] = (f.text or "").strip()
            else:
                if c.get("t") == "e":
                    raise ModelError(f"Error cell encountered in Python evaluator at {sname}!{ref}")
                v = c.findtext(f"{{{NS}}}v")
                if v is not None and v.strip() != "" and c.get("t", "n") == "n":
                    try:
                        values[(sname, ref)] = Decimal(v.strip())
                    except InvalidOperation as exc:
                        raise ModelError(f"Non-numeric value at {sname}!{ref}") from exc

    supported_funcs = {"SUM", "AVERAGE", "MIN", "MAX"}

    def eval_expr(sheet: str, expr: str) -> Decimal | None:
        expr = expr.strip()
        for fn_match in FUNC_CALL_RE.finditer(expr):
            fn_name = fn_match.group(1).upper()
            if fn_name not in supported_funcs:
                raise ModelError(f"Unsupported function in Python evaluator: {fn_name} in {expr!r}")

        m_func = re.fullmatch(r"(SUM|AVERAGE|MIN|MAX)\(([^()]+)\)", expr, re.I)
        if m_func:
            fn = m_func.group(1).upper()
            arg_str = m_func.group(2)
            items: list[Decimal] = []
            for token in arg_str.split(","):
                token = token.strip()
                if "!" in token:
                    sp, rp = token.split("!", 1)
                    sp = sp.strip("'")
                else:
                    sp, rp = sheet, token
                rp = rp.replace("$", "")
                for cr in sorted(cell_range(rp)):
                    if (sp, cr) not in values:
                        if (sp, cr) in formulas:
                            return None
                        raise ModelError(f"Python evaluator cannot resolve missing/unpopulated cell {sp}!{cr}")
                    items.append(values[(sp, cr)])
            if not items:
                raise ModelError(f"Empty argument list in {expr!r}")
            if fn == "SUM":
                return sum(items, Decimal(0))
            if fn == "AVERAGE":
                return sum(items, Decimal(0)) / Decimal(len(items))
            if fn == "MIN":
                return min(items)
            if fn == "MAX":
                return max(items)

        if FUNC_CALL_RE.search(expr):
            raise ModelError(f"Nested or complex function call unsupported by Python evaluator: {expr!r}")

        missing_dep = False

        def repl_ref(m: re.Match) -> str:
            nonlocal missing_dep
            s_part = m.group(1)
            c_ref = m.group(2).replace("$", "")
            target_sheet = s_part[:-1].strip("'") if s_part else sheet
            if (target_sheet, c_ref) not in values:
                if (target_sheet, c_ref) in formulas:
                    missing_dep = True
                    return "0"
                raise ModelError(f"Python evaluator cannot resolve missing/unpopulated cell {target_sheet}!{c_ref}")
            val = values[(target_sheet, c_ref)]
            return f"({val})"

        py_expr = CELL_REF_RE.sub(repl_ref, expr).replace("^", "**")
        if missing_dep:
            return None
        if not re.fullmatch(r"[0-9().+\-*/\sE]+", py_expr, re.I):
            raise ModelError(f"Formula requires desktop Excel recalculation engine: {expr}")
        try:
            tree_ast = ast.parse(py_expr, mode="eval")
        except SyntaxError as exc:
            raise ModelError(f"Invalid formula syntax in Python evaluator: {expr}") from exc
        return _eval_decimal_ast(tree_ast)

    unresolved = set(formulas)
    for _ in range(len(formulas) + 2):
        if not unresolved:
            break
        progressed = False
        for key in list(unresolved):
            sname, ref = key
            new_val = eval_expr(sname, formulas[key])
            if new_val is not None:
                values[key] = new_val
                unresolved.remove(key)
                progressed = True
        if not progressed:
            break

    if unresolved:
        sample = ", ".join(f"{s}!{r}" for s, r in sorted(unresolved)[:5])
        raise ModelError(f"Unresolved or circular formula dependencies in Python evaluator: {sample}")

    changed = dict(pkg.parts)
    for sname, (part, tree) in pkg.sheets.items():
        root = deepcopy(tree)
        for c in root.iter(f"{{{NS}}}c"):
            ref = c.get("r")
            if (sname, ref) in formulas and (sname, ref) in values:
                c.attrib.pop("t", None)
                for old_v in list(c.findall(f"{{{NS}}}v")):
                    c.remove(old_v)
                v_el = ET.SubElement(c, f"{{{NS}}}v")
                v_el.text = str(values[(sname, ref)])
        ET.register_namespace("", NS)
        changed[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for entry in pkg.entries:
            if entry.filename in changed:
                archive.writestr(copy(entry), changed[entry.filename])
    return stream.getvalue()


def validate_candidate(
    workbook_path: Path,
    pack_path: Path,
    recalculate: bool = False,
    recalculated_output: Path | None = None,
    perturb_check: bool = False,
    template_path: Path | None = None,
    inputs_path: Path | None = None,
    use_python_recalc_fallback: bool = False,
    allow_proposal: bool = False,
) -> dict[str, Any]:
    run_id = uuid.uuid4().hex
    timestamp_utc = datetime.now(timezone.utc).isoformat()
    pack_bytes = pack_path.read_bytes()
    pack = read_json(str(pack_path))
    is_proposal = pack.get("status") == "proposal" or pack.get("approval") is None
    if is_proposal and not allow_proposal:
        raise ModelError("Pack is an unapproved proposal; pass --allow-proposal to run validation on a proposal pack")
    if template_path is not None:
        validate_pack(pack, Package(template_path.read_bytes()), allow_proposal=allow_proposal)

    input_workbook_bytes = workbook_path.read_bytes()
    input_workbook_sha256 = sha(input_workbook_bytes)
    target_path = workbook_path
    recalc_info: dict[str, Any] | None = None
    temp_dir: tempfile.TemporaryDirectory | None = None

    try:
        if recalculate:
            if recalculated_output is None:
                temp_dir = tempfile.TemporaryDirectory()
                recalculated_output = Path(temp_dir.name) / f"recalc-{workbook_path.name}"
            else:
                if recalculated_output.resolve() == workbook_path.resolve():
                    raise ModelError("Recalculated output path must differ from input workbook path")
                if template_path is not None and recalculated_output.resolve() == template_path.resolve():
                    raise ModelError("Recalculated output path must differ from template path")
                if recalculated_output.exists():
                    raise ModelError(f"Recalculated output already exists and will not be overwritten: {recalculated_output}")
            if use_python_recalc_fallback:
                recalc_bytes = recalculate_package_python(input_workbook_bytes)
                with recalculated_output.open("xb") as f_out:
                    f_out.write(recalc_bytes)
                recalc_info = {
                    "engine": "Python Synthetic Fixture Evaluator (partial, non-certifying)",
                    "engine_version": "python-ast-1.0",
                    "synthetic_only": True,
                    "input_sha256": input_workbook_sha256,
                    "output_sha256": sha(recalc_bytes),
                    "output_path": str(recalculated_output),
                }
            else:
                recalc_info = recalculate_excel_com(workbook_path, recalculated_output)
            target_path = recalculated_output

        evaluated_bytes = target_path.read_bytes()
        evaluated_sha256 = sha(evaluated_bytes)
        pkg = Package(evaluated_bytes)
        if pkg.blockers():
            raise ModelError(f"Evaluated workbook contains formula/package blockers: {'; '.join(pkg.blockers())}")
        pack_excluded = set(pack.get("excluded_schedules", []))
        errors, missing_caches = scan_formula_errors_and_missing_caches(pkg, excluded_sheets=pack_excluded)
        pack_checks = evaluate_pack_checks(pkg, pack)
        identities = run_equity_research_identities(pkg) if is_equity_research_workbook(pkg) else []

        perturbation_report: dict[str, Any] | None = None
        if perturb_check:
            if template_path is None or inputs_path is None:
                raise ModelError("--perturb-check requires both --template and --inputs")
            template_bytes = template_path.read_bytes()
            supplied = read_json(str(inputs_path))
            base_bytes, _ = compose(template_bytes, pack, supplied, allow_proposal=allow_proposal)

            perturb_field = next(
                (f for f in pack["inputs"] if f["name"] == "dcf2_terminal_growth_rate"),
                pack["inputs"][0],
            )
            p_name = perturb_field["name"]
            orig_val = numeric(supplied["values"][p_name]["value"])
            delta = Decimal("0.005") if orig_val < Decimal("1") else Decimal("1")
            new_val = orig_val + delta
            if "maximum" in perturb_field and new_val > numeric(perturb_field["maximum"]):
                new_val = orig_val - delta

            perturbed_supplied = deepcopy(supplied)
            perturbed_supplied["values"][p_name]["value"] = new_val
            pert_bytes, _ = compose(template_bytes, pack, perturbed_supplied, allow_proposal=allow_proposal)

            restored_supplied = deepcopy(perturbed_supplied)
            restored_supplied["values"][p_name]["value"] = orig_val
            restored_bytes, _ = compose(template_bytes, pack, restored_supplied, allow_proposal=allow_proposal)
            restoration_identical = sha(restored_bytes) == sha(base_bytes)

            with tempfile.TemporaryDirectory() as pdir:
                pert_in = Path(pdir) / "perturbed_in.xlsx"
                pert_out = Path(pdir) / "perturbed_recalc.xlsx"
                pert_in.write_bytes(pert_bytes)
                if use_python_recalc_fallback:
                    pert_out.write_bytes(recalculate_package_python(pert_bytes))
                else:
                    recalculate_excel_com(pert_in, pert_out)
                pert_recalc_bytes = pert_out.read_bytes()
                pert_pkg = Package(pert_recalc_bytes)
                pert_errors, pert_missing = scan_formula_errors_and_missing_caches(pert_pkg, excluded_sheets=pack_excluded)
                pert_identities = run_equity_research_identities(pert_pkg) if is_equity_research_workbook(pert_pkg) else []
                moved_checks = []
                for chk in pack["checks"]:
                    b_v = get_cell_decimal(pkg, chk["sheet"], chk["cell"])
                    p_v = get_cell_decimal(pert_pkg, chk["sheet"], chk["cell"])
                    if b_v is not None and p_v is not None and b_v != p_v:
                        moved_checks.append({
                            "name": chk["name"],
                            "baseline": str(b_v),
                            "perturbed": str(p_v),
                        })
                perturbation_passed = (
                    not pert_errors
                    and not pert_missing
                    and all(i["passed"] for i in pert_identities)
                    and len(moved_checks) > 0
                    and restoration_identical
                )
                perturbation_report = {
                    "passed": perturbation_passed,
                    "perturbed_input": p_name,
                    "original_value": str(orig_val),
                    "perturbed_value": str(new_val),
                    "baseline_candidate_sha256": sha(base_bytes),
                    "perturbed_candidate_sha256": sha(pert_bytes),
                    "perturbed_recalculated_sha256": sha(pert_recalc_bytes),
                    "restored_candidate_sha256": sha(restored_bytes),
                    "moved_check_cells": moved_checks,
                    "perturbed_error_count": len(pert_errors),
                    "perturbed_uncalculated_count": len(pert_missing),
                    "perturbed_identities_passed": all(i["passed"] for i in pert_identities),
                    "restoration_byte_identical": restoration_identical,
                }

        all_pack_passed = all(c["passed"] for c in pack_checks)
        all_identities_passed = all(i["passed"] for i in identities)
        pert_ok = True if perturbation_report is None else perturbation_report["passed"]
        fully_recalculated = len(missing_caches) == 0 and len(errors) == 0

        pending_gates: list[str] = []
        if is_proposal:
            pending_gates.append("human_pack_approval")
        if not fully_recalculated:
            pending_gates.append("engine_recalculation")
        elif use_python_recalc_fallback:
            pending_gates.append("full_spreadsheet_engine_recalculation")
        if errors:
            pending_gates.append("formula_error_scan")
        if not (all_pack_passed and all_identities_passed):
            pending_gates.append("independent_model_checks")
        if perturbation_report is None or not pert_ok:
            pending_gates.append("input_perturbation_and_restoration")
        pending_gates.append("render_review")

        checks_passed = bool(fully_recalculated and all_pack_passed and all_identities_passed and pert_ok)
        if missing_caches:
            status = "needs_recalculation"
        elif checks_passed:
            status = "recalculated_checks_passed"
        else:
            status = "blocked"

        return {
            "run_id": run_id,
            "timestamp_utc": timestamp_utc,
            "status": status,
            "ready": False,
            "checks_passed": checks_passed,
            "pack_status": "proposal" if is_proposal else "approved",
            "calculation_engine": recalc_info["engine"] if recalc_info else "none (cached values inspected)",
            "calculation_engine_version": recalc_info["engine_version"] if recalc_info else None,
            "template_sha256": sha(template_path.read_bytes()) if template_path is not None else pack.get("template_sha256"),
            "pack_sha256": sha(pack_bytes),
            "inputs_sha256": sha(inputs_path.read_bytes()) if inputs_path is not None else None,
            "candidate_input_workbook_sha256": input_workbook_sha256,
            "workbook_sha256": evaluated_sha256,
            "recalculation": recalc_info,
            "excluded_schedules": sorted(pack_excluded),
            "formula_error_count": len(errors),
            "formula_errors": errors,
            "uncalculated_formula_count": len(missing_caches),
            "uncalculated_formula_sample": missing_caches[:20],
            "pack_checks_passed": sum(1 for c in pack_checks if c["passed"]),
            "pack_checks_total": len(pack_checks),
            "pack_checks": pack_checks,
            "independent_identities_passed": sum(1 for i in identities if i["passed"]),
            "independent_identities_total": len(identities),
            "independent_identities": identities,
            "perturbation_and_restoration": perturbation_report,
            "pending_gates": pending_gates,
        }
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", help="Candidate or recalculated XLSX path")
    parser.add_argument("--pack", required=True, help="Approved or proposal model-pack JSON path")
    parser.add_argument("--recalculate", action="store_true", help="Recalculate via Excel COM before validating")
    parser.add_argument("--recalculated-output", help="Path to write the recalculated XLSX file (must not exist)")
    parser.add_argument("--perturb-check", action="store_true", help="Run input perturbation and restoration verification")
    parser.add_argument("--template", help="Original template XLSX path (required for --perturb-check)")
    parser.add_argument("--inputs", help="Baseline inputs JSON path (required for --perturb-check)")
    parser.add_argument("--report", help="Optional path to write the JSON validation report (must not exist)")
    parser.add_argument("--allow-proposal", action="store_true", help="Allow validating against an unapproved proposal pack")
    args = parser.parse_args()
    try:
        if args.report and Path(args.report).exists():
            raise ModelError(f"Report output already exists and will not be overwritten: {args.report}")
        report = validate_candidate(
            workbook_path=Path(args.workbook),
            pack_path=Path(args.pack),
            recalculate=args.recalculate,
            recalculated_output=Path(args.recalculated_output) if args.recalculated_output else None,
            perturb_check=args.perturb_check,
            template_path=Path(args.template) if args.template else None,
            inputs_path=Path(args.inputs) if args.inputs else None,
            allow_proposal=args.allow_proposal,
        )
        text = json.dumps(report, indent=2)
        if args.report:
            with Path(args.report).open("x", encoding="utf-8") as report_file:
                report_file.write(text + "\n")
        print(text)
        return 0 if report.get("checks_passed") else 1
    except (ModelError, OSError, KeyError, TypeError) as exc:
        err_json = json.dumps({"status": "blocked", "ready": False, "error": str(exc)}, indent=2)
        print(err_json)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
