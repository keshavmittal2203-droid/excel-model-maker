"""Preservation-compatible adapter for special formula structures.

The conservative composer previously refused every workbook containing
shared, array, or dataTable formula elements. This module replaces that
blanket refusal with structural validation plus explicit protection, so a
workbook is accepted only when every special structure is well-formed and
the edit machinery can prove it is preserved:

* shared formulas - master text, ``si``, and ``ref`` are kept byte-for-byte;
  followers must be empty, must resolve to a master, and must lie inside the
  master's reference range.
* array formulas - master text and ``ref`` are kept byte-for-byte; the whole
  reference range becomes a protected zone: mapped inputs may not intersect
  it and stale cached values inside it are invalidated.
* dataTable - the what-if result range becomes a protected zone; the
  row/column input references (``r1``/``r2``) are recorded for review.
* everything else (unknown ``t`` values, attributes on normal formulas,
  prefixed cell tags inside sheetData, broken references) keeps the
  conservative guard and blocks composition.

Nothing in this module evaluates formulas. A workbook whose special
structures pass this adapter still requires actual recalculation and
independent financial checks before any validated state (QUALITY.md).
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
import re

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Maximum cells a protected zone may cover; larger zones are refused so the
# stale-cache scan cannot be forced to expand a whole sheet.
MAX_ZONE_CELLS = 100_000

# Attributes permitted on each special formula kind. Anything else is
# refused with an explicit message instead of being guessed at.
ARRAY_ATTRIBUTES = {"ref"}
DATATABLE_ATTRIBUTES = {"ref", "dt2D", "dtr", "r1", "r2", "ca"}
SHARED_MASTER_ATTRIBUTES = {"si", "ref"}
SHARED_FOLLOWER_ATTRIBUTES = {"si"}

FORMULA_FRAGMENT = re.compile(rb"<f\b[^>]*?/>|<f\b[^>]*>.*?</f>", re.S)
CELL_FRAGMENT = re.compile(rb"<c\b[^>]*?/>|<c\b[^>]*>.*?</c>", re.S)
VALUE_FRAGMENT = re.compile(rb"<v\b[^>]*?/>|<v\b[^>]*>.*?</v>", re.S)
SHEET_DATA = re.compile(rb"<sheetData\b[^>]*?/>|<sheetData\b[^>]*>.*?</sheetData>", re.S)
ADDRESS = re.compile(r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")


def _ModelError(message):
    from modelmaker import ModelError  # deferred: avoids import cycle
    return ModelError(message)


def parse_range(ref):
    """Return (first_col, first_row, last_col, last_row) or raise ValueError."""
    lo, _, hi = ref.partition(":")
    hi = hi or lo
    if not ADDRESS.match(lo) or not ADDRESS.match(hi):
        raise ValueError("malformed range")

    def parts(address):
        col = 0
        for ch in re.match(r"^([A-Z]{1,3})", address)[1]:
            col = col * 26 + ord(ch) - 64
        return col, int(re.match(r"^[A-Z]{1,3}(\d+)$", address)[1])

    first = parts(lo)
    last = parts(hi)
    if first[0] > last[0] or first[1] > last[1]:
        raise ValueError("reversed range")
    if last[0] > 16384 or last[1] > 1048576:
        raise ValueError("range outside Excel bounds")
    return first[0], first[1], last[0], last[1]


def range_area(ref):
    first_col, first_row, last_col, last_row = parse_range(ref)
    return (last_col - first_col + 1) * (last_row - first_row + 1)


def in_range(ref, cell):
    """True when the single-cell address lies inside the range reference."""
    first_col, first_row, last_col, last_row = parse_range(ref)
    match = ADDRESS.match(cell)
    if not match:
        raise ValueError("malformed cell address")
    col = 0
    for ch in re.match(r"^([A-Z]{1,3})", cell)[1]:
        col = col * 26 + ord(ch) - 64
    row = int(re.match(r"^[A-Z]{1,3}(\d+)$", cell)[1])
    return first_col <= col <= last_col and first_row <= row <= last_row


def ranges_overlap(left, right):
    a = parse_range(left)
    b = parse_range(right)
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def range_cells(ref):
    """All single-cell addresses of a (bounded) range."""
    first_col, first_row, last_col, last_row = parse_range(ref)

    def letters(number):
        text = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            text = chr(65 + remainder) + text
        return text

    return [f"{letters(col)}{row}" for row in range(first_row, last_row + 1) for col in range(first_col, last_col + 1)]


def top_left(ref):
    return ref.partition(":")[0]


@dataclass
class FormulaReport:
    """Per-sheet classification of special formula structures."""

    sheet: str
    counts: dict = field(default_factory=dict)
    problems: list = field(default_factory=list)
    protected_zones: list = field(default_factory=list)
    table_input_cells: list = field(default_factory=list)
    stale_cells: set = field(default_factory=set)


def classify(tree, sheet):
    """Classify and validate every formula element on one worksheet."""
    cells = {c.get("r"): c for c in tree.iter(f"{{{NS}}}c") if c.get("r")}
    report = FormulaReport(sheet=sheet, counts=dict(Counter()))
    masters = {}
    followers = []
    zones = []

    for cell in tree.iter(f"{{{NS}}}c"):
        formula = cell.find(f"{{{NS}}}f")
        if formula is None:
            continue
        kind = formula.get("t", "normal")
        text = formula.text or ""
        ref = formula.get("ref")
        address = cell.get("r") or "?"
        attributes = set(formula.attrib) - {"t"}
        report.counts[kind] = report.counts.get(kind, 0) + 1

        if kind == "normal":
            if attributes:
                report.problems.append(f"{sheet}!{address}: attributes on a normal formula are unsupported: {sorted(attributes)}")
            continue

        if kind == "shared":
            si = formula.get("si")
            if text.strip():
                # master cell
                unexpected = attributes - SHARED_MASTER_ATTRIBUTES
                if unexpected:
                    report.problems.append(f"{sheet}!{address}: unsupported shared-master attributes: {sorted(unexpected)}")
                if si is None:
                    report.problems.append(f"{sheet}!{address}: shared master is missing si")
                elif si in masters:
                    report.problems.append(f"{sheet}!{address}: duplicate shared master si={si}")
                elif not ref:
                    report.problems.append(f"{sheet}!{address}: shared master is missing ref")
                else:
                    try:
                        if address != top_left(ref):
                            report.problems.append(f"{sheet}!{address}: shared master is not the top-left of {ref}")
                        if range_area(ref) > MAX_ZONE_CELLS:
                            report.problems.append(f"{sheet}!{address}: shared range too large: {ref}")
                    except ValueError as exc:
                        report.problems.append(f"{sheet}!{address}: malformed shared range {ref!r}: {exc}")
                    masters[si] = (address, ref)
            else:
                # follower cell
                unexpected = attributes - SHARED_FOLLOWER_ATTRIBUTES
                if unexpected:
                    report.problems.append(f"{sheet}!{address}: unsupported shared-follower attributes: {sorted(unexpected)}")
                if si is None:
                    report.problems.append(f"{sheet}!{address}: shared follower is missing si")
                else:
                    followers.append((si, address))
            continue

        if kind == "array":
            unexpected = attributes - ARRAY_ATTRIBUTES
            if unexpected:
                report.problems.append(f"{sheet}!{address}: unsupported array attributes: {sorted(unexpected)}")
            if not text.strip():
                report.problems.append(f"{sheet}!{address}: array master has empty formula text")
            if not ref:
                report.problems.append(f"{sheet}!{address}: array master is missing ref")
            else:
                try:
                    if address != top_left(ref):
                        report.problems.append(f"{sheet}!{address}: array master is not the top-left of {ref}")
                    if range_area(ref) > MAX_ZONE_CELLS:
                        report.problems.append(f"{sheet}!{address}: array range too large: {ref}")
                except ValueError as exc:
                    report.problems.append(f"{sheet}!{address}: malformed array range {ref!r}: {exc}")
                zones.append(ref)
            continue

        if kind == "dataTable":
            unexpected = attributes - DATATABLE_ATTRIBUTES
            if unexpected:
                report.problems.append(f"{sheet}!{address}: unsupported dataTable attributes: {sorted(unexpected)}")
            if text.strip():
                report.problems.append(f"{sheet}!{address}: dataTable formula text must be empty")
            if not ref:
                report.problems.append(f"{sheet}!{address}: dataTable is missing ref")
            else:
                try:
                    if address != top_left(ref):
                        report.problems.append(f"{sheet}!{address}: dataTable is not the top-left of {ref}")
                    if range_area(ref) > MAX_ZONE_CELLS:
                        report.problems.append(f"{sheet}!{address}: dataTable range too large: {ref}")
                except ValueError as exc:
                    report.problems.append(f"{sheet}!{address}: malformed dataTable range {ref!r}: {exc}")
                zones.append(ref)
            for attribute in ("dt2D", "dtr", "r1", "r2"):
                if attribute not in formula.attrib:
                    report.problems.append(f"{sheet}!{address}: dataTable is missing {attribute}")
            for attribute in ("r1", "r2"):
                value = (formula.get(attribute) or "").replace("$", "")
                if value and not ADDRESS.match(value):
                    report.problems.append(f"{sheet}!{address}: dataTable {attribute}={value!r} is not a single-cell reference")
                if value:
                    report.table_input_cells.append(value)
            continue

        report.problems.append(f"{sheet}!{address}: unsupported formula kind t={kind!r}")

    for si, address in followers:
        if si not in masters:
            report.problems.append(f"{sheet}!{address}: shared follower has unresolved si={si}")
            continue
        master_address, master_ref = masters[si]
        try:
            if not in_range(master_ref, address):
                report.problems.append(f"{sheet}!{address}: shared follower lies outside master range {master_ref} (si={si})")
        except ValueError as exc:
            report.problems.append(f"{sheet}!{address}: malformed shared master range {master_ref!r}: {exc}")
        if address == master_address:
            report.problems.append(f"{sheet}!{address}: shared follower duplicates its master cell")

    for index, zone in enumerate(zones):
        for other in zones[index + 1:]:
            if ranges_overlap(zone, other):
                report.problems.append(f"{sheet}: protected zones overlap: {zone} and {other}")

    report.protected_zones = zones
    for zone in zones:
        for address in range_cells(zone):
            cell = cells.get(address)
            if cell is None:
                continue
            if cell.find(f"{{{NS}}}f") is None and cell.find(f"{{{NS}}}v") is not None:
                report.stale_cells.add(address)
    return report


def formula_fragments(data):
    """Every <f> element in the worksheet bytes, in document order."""
    return [match[0] for match in FORMULA_FRAGMENT.finditer(data)]


def strip_stale(data, refs):
    """Remove cached <v> values from listed non-formula cells (protected zones)."""
    from modelmaker import attribute, set_attribute  # deferred

    if not refs:
        return data
    regions = list(SHEET_DATA.finditer(data))
    if len(regions) != 1:
        raise _ModelError("Worksheet must contain exactly one supported sheetData element")
    region = regions[0]

    def replace(match):
        fragment = match[0]
        opening_end = fragment.find(b">") + 1
        opening = fragment[:opening_end]
        reference = attribute(opening, "r")
        address = reference.decode("ascii") if reference else ""
        if address not in refs or re.search(rb"<f(?:\s|/?>)", fragment):
            return fragment
        opening = set_attribute(opening, "t", None)
        if fragment.endswith(b"/>"):
            return opening + fragment[opening_end:]
        return opening + VALUE_FRAGMENT.sub(b"", fragment[opening_end:])

    result = CELL_FRAGMENT.sub(replace, region[0])
    return data[:region.start()] + result + data[region.end():]


def zone_contains(protected_zones, cell):
    """True when the single-cell address falls inside any protected zone."""
    for zone in protected_zones:
        try:
            if in_range(zone, cell):
                return zone
        except ValueError:
            continue
    return None
