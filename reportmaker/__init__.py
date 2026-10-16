"""Styled JSON-to-XLSX report builder (the cli.py workstream).

Builds a new presentation workbook from a JSON spec: professional styling,
number/currency/percent/date formats, SUM totals as real Excel formulas,
conditional formatting, and an embedded chart. It does not edit templates,
does not evaluate formulas, and never certifies financial correctness.
"""
from .builder import build_workbook, number_format, write_report
from .errors import ReportError
from .spec import ChartSpec, Column, ReportSpec, parse_spec, spec_from_file

__version__ = "0.2.0"

__all__ = [
    "ChartSpec",
    "Column",
    "ReportError",
    "ReportSpec",
    "build_workbook",
    "number_format",
    "parse_spec",
    "spec_from_file",
    "write_report",
    "__version__",
]
