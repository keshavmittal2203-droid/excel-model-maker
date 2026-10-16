"""Command-line interface for the styled report builder."""
from __future__ import annotations

import argparse
import json

from . import __version__
from .builder import write_report
from .errors import ReportError
from .spec import spec_from_file


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="cli.py",
        description="Build a styled Excel report from a JSON spec. Prints a JSON manifest on success; exit code 2 on failure.",
    )
    parser.add_argument("--input", "-i", required=True, metavar="data.json", help="report specification JSON file")
    parser.add_argument("--output", "-o", required=True, metavar="report.xlsx", help="workbook to create; refuses to overwrite without --force")
    parser.add_argument("--force", action="store_true", help="overwrite the output file if it already exists")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)
    try:
        spec = spec_from_file(args.input)
        manifest = write_report(spec, args.output, overwrite=args.force)
    except ReportError as exc:
        print(json.dumps({"status": "blocked", "ready": False, "error": str(exc)}, indent=2))
        return 2
    print(json.dumps(manifest, indent=2))
    return 0
