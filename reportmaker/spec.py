"""Strict parsing and validation of a styled-report specification.

The spec is a JSON object. Unknown fields, duplicate JSON keys, NaN/Infinity
constants, booleans in numeric positions, and mismatched row keys are all
rejected with a message that names the exact location of the problem.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import json
from pathlib import Path
import re
from typing import Any, Mapping

from .errors import ReportError

TYPES = ("text", "integer", "number", "currency", "percent", "date")
CONDITIONALS = ("color_scale", "data_bar")
CHART_TYPES = ("bar", "line", "pie")

MAX_COLUMNS = 256
MAX_ROWS = 100_000

_DEFAULT_DECIMALS = {"integer": 0, "number": 2, "currency": 2, "percent": 1}
_SHEET_FORBIDDEN = set(':\\/?*[]')
_CELL_RE = re.compile(r"[A-Z]{1,3}[1-9][0-9]{0,6}\Z")
_DATE_RE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}\Z")

_ALLOWED_TOP = ("title", "sheet_name", "columns", "rows", "chart")
_ALLOWED_COLUMN = ("key", "label", "type", "decimals", "currency_symbol", "total", "conditional")
_ALLOWED_CONDITIONAL = ("type",)
_ALLOWED_CHART = ("type", "title", "categories", "series", "position")


@dataclass(frozen=True)
class Column:
    key: str
    label: str
    type: str
    decimals: int
    currency_symbol: str = "$"
    total: bool = False
    conditional: str | None = None


@dataclass(frozen=True)
class ChartSpec:
    type: str
    title: str | None
    categories: str
    series: tuple[str, ...]
    position: str | None = None


@dataclass(frozen=True)
class ReportSpec:
    title: str | None
    sheet_name: str
    columns: tuple[Column, ...]
    rows: tuple[Mapping[str, Any], ...]
    chart: ChartSpec | None = None


def _require_mapping(value: Any, where: str) -> dict:
    if not isinstance(value, dict):
        raise ReportError(f"{where} must be a JSON object")
    return value


def _check_keys(obj: Mapping, allowed: tuple[str, ...], required: tuple[str, ...], where: str) -> None:
    unknown = sorted(str(key) for key in obj if key not in allowed)
    if unknown:
        raise ReportError(f"{where} has unknown field(s): {', '.join(unknown)}")
    missing = [name for name in required if name not in obj]
    if missing:
        raise ReportError(f"{where} is missing required field(s): {', '.join(missing)}")


def _string(value: Any, where: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ReportError(f"{where} must be a non-empty string; got {value!r}")
    return value


def _number(value: Any, where: str) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ReportError(f"{where} must be a JSON number; got {value!r}")
    if isinstance(value, float) and (value != value or value in (float("inf"), float("-inf"))):
        raise ReportError(f"{where} must be a finite number")
    return value


def _parse_column(raw: Any, index: int) -> Column:
    where = f"columns[{index}]"
    _require_mapping(raw, where)
    _check_keys(raw, _ALLOWED_COLUMN, ("key",), where)
    key = _string(raw["key"], f"{where}.key")
    label = _string(raw.get("label", key), f"{where}.label")
    ctype = raw.get("type", "text")
    if ctype not in TYPES:
        raise ReportError(f"{where}.type must be one of {', '.join(TYPES)}; got {ctype!r}")
    decimals = raw.get("decimals", _DEFAULT_DECIMALS.get(ctype, 0))
    if isinstance(decimals, bool) or not isinstance(decimals, int) or not 0 <= decimals <= 10:
        raise ReportError(f"{where}.decimals must be an integer between 0 and 10")
    if ctype != "currency" and "currency_symbol" in raw:
        raise ReportError(f"{where}.currency_symbol is only valid for currency columns")
    symbol = _string(raw.get("currency_symbol", "$"), f"{where}.currency_symbol") if ctype == "currency" else "$"
    total = raw.get("total", False)
    if not isinstance(total, bool):
        raise ReportError(f"{where}.total must be true or false")
    if total and ctype not in ("integer", "number", "currency"):
        raise ReportError(f"{where}.total is supported only for integer, number, or currency columns")
    conditional = raw.get("conditional")
    if conditional is not None:
        if isinstance(conditional, str):
            conditional = {"type": conditional}
        _require_mapping(conditional, f"{where}.conditional")
        _check_keys(conditional, _ALLOWED_CONDITIONAL, ("type",), f"{where}.conditional")
        conditional = conditional["type"]
        if conditional not in CONDITIONALS:
            raise ReportError(f"{where}.conditional.type must be one of {', '.join(CONDITIONALS)}")
        if ctype not in ("integer", "number", "currency", "percent"):
            raise ReportError(f"{where}.conditional is supported only for numeric columns")
    return Column(key=key, label=label, type=ctype, decimals=decimals, currency_symbol=symbol, total=total, conditional=conditional)


def _parse_value(value: Any, column: Column, where: str) -> Any:
    if column.type == "text":
        if not isinstance(value, str):
            raise ReportError(f"{where} must be a string (text column {column.key!r})")
        return value
    if column.type == "date":
        if not isinstance(value, str) or not _DATE_RE.fullmatch(value):
            raise ReportError(f"{where} must be an ISO date string YYYY-MM-DD (date column {column.key!r})")
        try:
            return date.fromisoformat(value)
        except ValueError as exc:
            raise ReportError(f"{where} is not a valid calendar date: {value!r}") from exc
    number = _number(value, where)
    if column.type == "integer":
        if isinstance(number, float):
            if not number.is_integer():
                raise ReportError(f"{where} must be a whole number (integer column {column.key!r})")
            number = int(number)
    return number


def _parse_rows(raw_rows: Any, columns: tuple[Column, ...]) -> tuple[Mapping[str, Any], ...]:
    if not isinstance(raw_rows, list) or not raw_rows:
        raise ReportError("rows must be a non-empty JSON array")
    if len(raw_rows) > MAX_ROWS:
        raise ReportError(f"rows exceeds the supported limit of {MAX_ROWS} data rows")
    keys = [column.key for column in columns]
    rows = []
    for index, raw in enumerate(raw_rows):
        where = f"rows[{index}]"
        _require_mapping(raw, where)
        missing = [key for key in keys if key not in raw]
        extra = sorted(key for key in raw if key not in keys)
        if missing or extra:
            details = []
            if missing:
                details.append("missing " + ", ".join(missing))
            if extra:
                details.append("unknown " + ", ".join(extra))
            raise ReportError(f"{where} does not match the column keys ({'; '.join(details)})")
        rows.append({column.key: _parse_value(raw[column.key], column, f"{where}.{column.key}") for column in columns})
    return tuple(rows)


def _parse_chart(raw: Any, columns: tuple[Column, ...]) -> ChartSpec:
    where = "chart"
    _require_mapping(raw, where)
    _check_keys(raw, _ALLOWED_CHART, ("categories", "series"), where)
    ctype = raw.get("type", "bar")
    if ctype not in CHART_TYPES:
        raise ReportError(f"chart.type must be one of {', '.join(CHART_TYPES)}; got {ctype!r}")
    title = raw.get("title")
    if title is not None:
        title = _string(title, "chart.title")
    by_key = {column.key: column for column in columns}
    categories = _string(raw["categories"], "chart.categories")
    if categories not in by_key:
        raise ReportError(f"chart.categories refers to unknown column {categories!r}")
    if by_key[categories].type not in ("text", "date"):
        raise ReportError("chart.categories must be a text or date column")
    series = raw["series"]
    if not isinstance(series, list) or not series:
        raise ReportError("chart.series must be a non-empty array of column keys")
    for index, key in enumerate(series):
        key = _string(key, f"chart.series[{index}]")
        if key not in by_key:
            raise ReportError(f"chart.series[{index}] refers to unknown column {key!r}")
        if by_key[key].type not in ("integer", "number", "currency", "percent"):
            raise ReportError(f"chart.series[{index}] must be a numeric column; {key!r} is {by_key[key].type}")
    if len(set(series)) != len(series):
        raise ReportError("chart.series contains duplicate column keys")
    if ctype == "pie" and len(series) != 1:
        raise ReportError("pie charts require exactly one series column")
    position = raw.get("position")
    if position is not None:
        position = _string(position, "chart.position")
        if not _CELL_RE.fullmatch(position):
            raise ReportError(f"chart.position must be a cell address like 'H2'; got {position!r}")
    return ChartSpec(type=ctype, title=title, categories=categories, series=tuple(series), position=position)


def parse_spec(raw: Any) -> ReportSpec:
    """Validate a decoded JSON object and return the typed spec."""
    _require_mapping(raw, "The report spec")
    _check_keys(raw, _ALLOWED_TOP, ("columns", "rows"), "The report spec")
    title = raw.get("title")
    if title is not None:
        title = _string(title, "title")
    sheet_name = _string(raw.get("sheet_name", "Report"), "sheet_name")
    if len(sheet_name) > 31:
        raise ReportError("sheet_name must be 31 characters or fewer")
    if any(character in _SHEET_FORBIDDEN for character in sheet_name):
        raise ReportError("sheet_name must not contain any of : \\ / ? * [ ]")
    if sheet_name.startswith("'") or sheet_name.endswith("'"):
        raise ReportError("sheet_name must not start or end with an apostrophe")
    raw_columns = raw["columns"]
    if not isinstance(raw_columns, list) or not raw_columns:
        raise ReportError("columns must be a non-empty JSON array")
    if len(raw_columns) > MAX_COLUMNS:
        raise ReportError(f"columns exceeds the supported limit of {MAX_COLUMNS}")
    columns = tuple(_parse_column(entry, index) for index, entry in enumerate(raw_columns))
    seen: set[str] = set()
    for column in columns:
        if column.key in seen:
            raise ReportError(f"Duplicate column key: {column.key!r}")
        seen.add(column.key)
    rows = _parse_rows(raw["rows"], columns)
    chart = _parse_chart(raw["chart"], columns) if raw.get("chart") is not None else None
    return ReportSpec(title=title, sheet_name=sheet_name, columns=columns, rows=rows, chart=chart)


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ReportError(f"Duplicate JSON key: {key}")
        result[key] = value
    return result


def _reject_constant(value):
    raise ReportError(f"Invalid JSON numeric constant: {value}")


def spec_from_file(path: str | Path) -> ReportSpec:
    """Read a JSON file strictly and return the validated spec."""
    target = Path(path)
    if not target.is_file():
        raise ReportError(f"Input JSON not found: {target}")
    try:
        text = target.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError) as exc:
        raise ReportError(f"Could not read input file {target}: {exc}") from exc
    try:
        raw = json.loads(text, object_pairs_hook=_unique_object, parse_constant=_reject_constant)
    except ReportError:
        raise
    except json.JSONDecodeError as exc:
        raise ReportError(f"Invalid JSON in {target}: {exc.msg} (line {exc.lineno}, column {exc.colno})") from exc
    return parse_spec(raw)
