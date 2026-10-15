"""Compare cached cell values between two XLSX files (recalc evidence helper).

Usage: python tools/compare_values.py <file_a> <file_b> --cells "Sheet!C1" ...
Prints a table and exits 0 when every listed value matches within tolerance
(or differs, when --expect-diff is given for those cells).
"""
from __future__ import annotations

import argparse
from decimal import Decimal
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from modelmaker import ModelError, NS, Package  # noqa: E402


def value(package: Package, sheet: str, ref: str):
    tree = package.sheets[sheet][1]
    for cell in tree.iter(f"{{{NS}}}c"):
        if cell.get("r") != ref:
            continue
        if cell.get("t") == "e":
            raise ModelError(f"cached error at {sheet}!{ref}")
        cached = cell.findtext(f"{{{NS}}}v")
        return None if cached is None else Decimal(cached)
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file_a")
    parser.add_argument("file_b")
    parser.add_argument("--cells", nargs="+", required=True, help='"Sheet!C1" entries')
    parser.add_argument("--expect-diff", nargs="*", default=[], help="cells expected to differ (by exact label)")
    parser.add_argument("--tolerance", type=str, default="1e-9")
    args = parser.parse_args()
    tolerance = Decimal(args.tolerance)
    expect_diff = {entry.strip("'\"") for entry in (args.expect_diff or [])}
    try:
        a = Package(__import__("pathlib").Path(args.file_a).read_bytes())
        b = Package(__import__("pathlib").Path(args.file_b).read_bytes())
    except (ModelError, OSError) as exc:
        print(json.dumps({"error": str(exc)}))
        return 2
    rows = []
    for raw in args.cells:
        label = raw.strip("'\"")
        sheet, _, ref = label.partition("!")
        va, vb = value(a, sheet, ref), value(b, sheet, ref)
        diff = None if va is None or vb is None else vb - va
        ok = (diff is not None and abs(diff) > tolerance) if label in expect_diff else (
            diff is not None and abs(diff) <= tolerance)
        rows.append({"cell": label, "a": str(va), "b": str(vb), "diff": str(diff), "ok": ok})
    print(json.dumps(rows, indent=2))
    return 0 if all(row["ok"] for row in rows) else 1


if __name__ == "__main__":
    raise SystemExit(main())
