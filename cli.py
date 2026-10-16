"""Entry point: python cli.py --input data.json --output report.xlsx

Builds a styled presentation workbook from a JSON report specification.
See AGENTS.md for the spec format, rules, and design decisions.
"""
from reportmaker.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
