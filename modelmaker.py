"""Conservative XLSX input composer. Never certifies financial correctness."""
from __future__ import annotations

import argparse
from copy import copy, deepcopy
from datetime import datetime
from decimal import Decimal, InvalidOperation
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import re
from typing import Any
import xml.etree.ElementTree as ET
import zipfile
import zlib

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CELL = re.compile(rb"<c\b[^>]*?/>|<c\b[^>]*>.*?</c>", re.S)
VALUE = re.compile(rb"<v\b[^>]*?/>|<v\b[^>]*>.*?</v>", re.S)
FORMULA = re.compile(rb"<f\b[^>]*?/>|<f\b[^>]*>.*?</f>", re.S)
ADDRESS = re.compile(r"([A-Z]{1,3})([1-9][0-9]{0,6})\Z")
MAX_SIZE = 100 * 1024 * 1024
MAX_ENTRIES = 5000
DISALLOWED_REVIEWER_TOKENS = frozenset({
    "unapproved", "pending", "proposal", "todo", "tbd", "none", "null", "n/a",
    "agent", "ai", "argon", "gemini", "deepseek", "claude", "openai", "antigravity", "codex",
})


class ModelError(ValueError):
    """An input, package, or feature is not supported safely."""


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def numeric(value: Any) -> Decimal:
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ModelError("A finite JSON number is required; booleans/text are not numbers")
    try:
        result = Decimal(str(value))
    except InvalidOperation as exc:
        raise ModelError("Invalid number") from exc
    if not result.is_finite():
        raise ModelError("Non-finite numbers are forbidden")
    # Excel stores at most 15 significant decimal digits. Do not silently round.
    digits = list(result.as_tuple().digits)
    while len(digits) > 1 and digits[-1] == 0:
        digits.pop()
    significant = len(digits) if result else 1
    if significant > 15 or (result and not -307 <= result.adjusted() <= 307):
        raise ModelError("Number exceeds the supported Excel precision/range")
    return result


def coordinate(address: str) -> tuple[int, int]:
    match = ADDRESS.fullmatch(address) if isinstance(address, str) else None
    if not match:
        raise ModelError(f"Invalid cell address: {address!r}")
    col = 0
    for ch in match[1]:
        col = col * 26 + ord(ch) - 64
    row = int(match[2])
    if col > 16384 or row > 1048576:
        raise ModelError("Cell address outside Excel bounds")
    return col, row


def cell_range(ref: str) -> set[str]:
    if not isinstance(ref, str) or not ref:
        raise ModelError(f"Invalid range reference: {ref!r}")
    parts = ref.split(":")
    if len(parts) == 1:
        coordinate(parts[0])
        return {parts[0]}
    if len(parts) == 2:
        (c1, r1), (c2, r2) = coordinate(parts[0]), coordinate(parts[1])
        if c1 > c2 or r1 > r2:
            raise ModelError(f"Reversed range reference: {ref}")
        if (c2 - c1 + 1) * (r2 - r1 + 1) > 100000:
            raise ModelError(f"Formula range too large: {ref}")
        cells = set()
        for c in range(c1, c2 + 1):
            letters = ""
            temp = c
            while temp > 0:
                temp, rem = divmod(temp - 1, 26)
                letters = chr(65 + rem) + letters
            for r in range(r1, r2 + 1):
                cells.add(f"{letters}{r}")
        return cells
    raise ModelError(f"Invalid range reference: {ref}")


def xml(data: bytes) -> ET.Element:
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise ModelError("DTD/entity declarations are forbidden")
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise ModelError("Malformed package XML") from exc


def attribute(tag: bytes, name: str) -> bytes | None:
    match = re.search(rb"(?:\s)" + name.encode() + rb"\s*=\s*([\"'])(.*?)\1", tag, re.S)
    return match[2] if match else None


def set_attribute(tag: bytes, name: str, value: bytes | None) -> bytes:
    pattern = re.compile(rb"\s+" + name.encode() + rb"\s*=\s*([\"']).*?\1", re.S)
    tag = pattern.sub(b"", tag)
    if value is None:
        return tag
    position = tag.rfind(b"/>") if tag.endswith(b"/>") else tag.rfind(b">")
    return tag[:position] + b' ' + name.encode() + b'="' + value + b'"' + tag[position:]


class Package:
    def __init__(self, data: bytes):
        if len(data) > MAX_SIZE:
            raise ModelError("Workbook exceeds the supported size")
        self.data = data
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                entries = archive.infolist()
                names = [entry.filename for entry in entries]
                if len(entries) > MAX_ENTRIES or len(set(names)) != len(names):
                    raise ModelError("Duplicate or excessive ZIP entries")
                if sum(entry.file_size for entry in entries) > MAX_SIZE:
                    raise ModelError("Expanded workbook exceeds the supported size")
                for entry in entries:
                    path = PurePosixPath(entry.filename)
                    if path.is_absolute() or ".." in path.parts or "\\" in entry.filename:
                        raise ModelError("Unsafe ZIP member name")
                    if entry.flag_bits & 1:
                        raise ModelError("Encrypted ZIP entries are unsupported")
                self.entries = entries
                self.parts = {entry.filename: archive.read(entry) for entry in entries}
        except (zipfile.BadZipFile, RuntimeError, NotImplementedError, EOFError, zlib.error) as exc:
            raise ModelError("Unreadable or unsupported XLSX archive") from exc
        for required in ("[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
            if required not in self.parts:
                raise ModelError(f"Required package part missing: {required}")
        # Parse all XML once; fail before any output is written.
        for name, value in self.parts.items():
            if name.endswith((".xml", ".rels")):
                xml(value)
        self.workbook = xml(self.parts["xl/workbook.xml"])
        if self.workbook.tag != f"{{{NS}}}workbook":
            raise ModelError("Only transitional OOXML XLSX is currently supported")
        rels = xml(self.parts["xl/_rels/workbook.xml.rels"])
        relation_map = {}
        for rel in rels:
            if rel.get("Id") in relation_map:
                raise ModelError("Duplicate workbook relationship ID")
            relation_map[rel.get("Id")] = rel
        self.sheets: dict[str, tuple[str, ET.Element]] = {}
        for sheet in self.workbook.findall(f"{{{NS}}}sheets/{{{NS}}}sheet"):
            name = sheet.get("name")
            if not name or name in self.sheets:
                raise ModelError("Missing or duplicate sheet name")
            rel = relation_map.get(sheet.get(f"{{{REL}}}id"))
            if rel is None or rel.get("TargetMode") == "External":
                raise ModelError("Missing or external sheet relationship")
            if not (rel.get("Type") or "").endswith("/worksheet"):
                raise ModelError("Non-worksheet sheets need a dedicated compatibility adapter")
            target = rel.get("Target", "")
            if not target or "\\" in target or ".." in PurePosixPath(target).parts:
                raise ModelError("Unsupported sheet relationship target")
            part = target.lstrip("/") if target.startswith("/") else "xl/" + target
            if part not in self.parts:
                raise ModelError("Sheet relationship points to a missing part")
            tree = xml(self.parts[part])
            if tree.tag != f"{{{NS}}}worksheet":
                raise ModelError("Unexpected worksheet XML")
            if any(part == existing[0] for existing in self.sheets.values()):
                raise ModelError("Several sheets reference the same part")
            refs = [c.get("r") for c in tree.findall(f".//{{{NS}}}sheetData/{{{NS}}}row/{{{NS}}}c")]
            if len(refs) != len(set(refs)):
                raise ModelError("Duplicate cell reference")
            for ref in refs:
                coordinate(ref)
            self.sheets[name] = (part, tree)
        if not self.sheets:
            raise ModelError("Workbook has no worksheets")
        self.formula_cells: set[tuple[str, str]] = set()
        self.formula_followers: dict[str, set[str]] = {}
        self.special_formulas: dict[str, list[dict]] = {}
        self.formula_blockers: list[str] = []
        for sheet_name, (_, tree) in self.sheets.items():
            followers: set[str] = set()
            protected_zone_cells: set[str] = set()
            shared_masters: dict[str, tuple[str, set[str]]] = {}
            sheet_formulas: list[dict] = []
            for c in tree.findall(f".//{{{NS}}}sheetData/{{{NS}}}row/{{{NS}}}c"):
                ref = c.get("r")
                f = c.find(f"{{{NS}}}f")
                if f is None:
                    continue
                t = f.get("t", "normal")
                text = f.text or ""
                if "#REF!" in text.upper():
                    self.formula_blockers.append(f"Broken formula reference: {sheet_name}")
                if t == "normal":
                    if set(f.attrib) - {"t"}:
                        self.formula_blockers.append(f"Unsupported attributes on normal formula: {sheet_name}!{ref}")
                    if not text.strip():
                        self.formula_blockers.append(f"Empty normal formula: {sheet_name}!{ref}")
                    self.formula_cells.add((sheet_name, ref))
                elif t == "shared":
                    si = f.get("si")
                    if si is None or not si.isdigit():
                        self.formula_blockers.append(f"Malformed shared formula missing si: {sheet_name}!{ref}")
                        continue
                    if "ref" in f.attrib:
                        if set(f.attrib) - {"t", "si", "ref"}:
                            self.formula_blockers.append(f"Unsupported attributes on shared master: {sheet_name}!{ref}")
                            continue
                        if not text.strip():
                            self.formula_blockers.append(f"Shared formula master missing formula text: {sheet_name}!{ref}")
                            continue
                        if si in shared_masters:
                            self.formula_blockers.append(f"Duplicate shared formula group index: {sheet_name}")
                            continue
                        raw_ref = f.get("ref", "")
                        try:
                            rng = cell_range(raw_ref)
                        except ModelError:
                            self.formula_blockers.append(f"Invalid shared formula range: {sheet_name}!{ref}")
                            continue
                        top_left = raw_ref.partition(":")[0]
                        if top_left != ref:
                            self.formula_blockers.append(f"Shared formula master not at top-left of ref range: {sheet_name}!{ref}")
                            continue
                        shared_masters[si] = (ref, rng)
                        for cr in rng:
                            self.formula_cells.add((sheet_name, cr))
                        self.formula_cells.add((sheet_name, ref))
                        sheet_formulas.append({"type": "shared_master", "cell": ref, "si": si, "ref": raw_ref})
                    else:
                        if set(f.attrib) - {"t", "si"}:
                            self.formula_blockers.append(f"Unsupported attributes on shared follower: {sheet_name}!{ref}")
                            continue
                        if text.strip():
                            self.formula_blockers.append(f"Shared formula follower contains unexpected formula text: {sheet_name}!{ref}")
                            continue
                        self.formula_cells.add((sheet_name, ref))
                        sheet_formulas.append({"type": "shared_follower", "cell": ref, "si": si})
                elif t == "array":
                    if set(f.attrib) - {"t", "ref"}:
                        self.formula_blockers.append(f"Unsupported attributes on array formula: {sheet_name}!{ref}")
                        continue
                    if not text.strip():
                        self.formula_blockers.append(f"Array formula missing formula text: {sheet_name}!{ref}")
                        continue
                    array_ref = f.get("ref")
                    if not array_ref:
                        self.formula_blockers.append(f"Malformed array formula missing ref: {sheet_name}")
                        continue
                    try:
                        rng = cell_range(array_ref)
                    except ModelError:
                        self.formula_blockers.append(f"Invalid array formula range: {sheet_name}")
                        continue
                    if array_ref.partition(":")[0] != ref:
                        self.formula_blockers.append(f"Array formula master not at top-left of ref range: {sheet_name}!{ref}")
                        continue
                    if protected_zone_cells & rng:
                        self.formula_blockers.append(f"Overlapping special formula range: {sheet_name}!{ref}")
                        continue
                    protected_zone_cells.update(rng)
                    for cr in rng:
                        self.formula_cells.add((sheet_name, cr))
                        if cr != ref:
                            followers.add(cr)
                    sheet_formulas.append({"type": "array", "cell": ref, "ref": array_ref})
                elif t == "dataTable":
                    allowed_dt_attrs = {"t", "ref", "dt2D", "dtr", "r1", "r2", "ca"}
                    required_dt_attrs = {"t", "ref", "dt2D", "dtr", "r1", "r2"}
                    if set(f.attrib) - allowed_dt_attrs or not required_dt_attrs.issubset(set(f.attrib)):
                        self.formula_blockers.append(f"Malformed dataTable attributes: {sheet_name}!{ref}")
                        continue
                    if text.strip():
                        self.formula_blockers.append(f"dataTable formula must not contain text: {sheet_name}!{ref}")
                        continue
                    if f.get("dt2D") != "1" or f.get("dtr") != "1":
                        self.formula_blockers.append(f"Unsupported dataTable flags: {sheet_name}!{ref}")
                        continue
                    dt_ref = f.get("ref", "")
                    try:
                        coordinate(f.get("r1", ""))
                        coordinate(f.get("r2", ""))
                        rng = cell_range(dt_ref)
                    except ModelError:
                        self.formula_blockers.append(f"Invalid dataTable range or input cell reference: {sheet_name}!{ref}")
                        continue
                    if dt_ref.partition(":")[0] != ref:
                        self.formula_blockers.append(f"dataTable master not at top-left of ref range: {sheet_name}!{ref}")
                        continue
                    if protected_zone_cells & rng:
                        self.formula_blockers.append(f"Overlapping dataTable range: {sheet_name}!{ref}")
                        continue
                    protected_zone_cells.update(rng)
                    for cr in rng:
                        self.formula_cells.add((sheet_name, cr))
                        if cr != ref:
                            followers.add(cr)
                    sheet_formulas.append({"type": "dataTable", "cell": ref, "ref": dt_ref})
                else:
                    self.formula_blockers.append(f"Unsupported formula type: {sheet_name} ({t})")

            for item in sheet_formulas:
                if item["type"] == "shared_follower":
                    master = shared_masters.get(item["si"])
                    if master is None:
                        self.formula_blockers.append(f"Shared formula follower references missing group: {sheet_name}")
                    else:
                        master_ref, master_rng = master
                        if item["cell"] == master_ref or item["cell"] not in master_rng:
                            self.formula_blockers.append(
                                f"Shared formula follower outside master ref range: {sheet_name}!{item['cell']}"
                            )

            self.formula_followers[sheet_name] = followers
            if sheet_formulas:
                self.special_formulas[sheet_name] = sheet_formulas

    def blockers(self) -> list[str]:
        blockers = list(self.formula_blockers)
        calc = self.workbook.find(f"{{{NS}}}calcPr")
        if calc is not None and calc.get("iterate") in ("1", "true"):
            blockers.append("Iterative calculation requires a dedicated compatibility adapter")
        for name in self.parts:
            lower = name.lower()
            if any(mark in lower for mark in ("vbaproject", "_xmlsignatures/", "xl/externallinks/", "xl/querytables/")) or lower == "xl/connections.xml":
                blockers.append(f"Unsupported native/external feature: {name}")
        for sheet, (_, tree) in self.sheets.items():
            if tree.find(f"{{{NS}}}sheetProtection") is not None:
                blockers.append(f"Protected worksheet: {sheet}")
        return sorted(set(blockers))

    def summary(self) -> dict:
        sheets = []
        for name, (_, tree) in self.sheets.items():
            cells = list(tree.iter(f"{{{NS}}}c"))
            errors = [{"cell": c.get("r"), "error": c.findtext(f"{{{NS}}}v")} for c in cells if c.get("t") == "e"]
            formulas = [c for c in cells if c.find(f"{{{NS}}}f") is not None]
            sheets.append({"name": name, "cells": len(cells), "formulas": len(formulas), "cached_errors": errors})
        special = {s: len(items) for s, items in self.special_formulas.items()}
        return {"sha256": sha(self.data), "sheets": sheets, "special_formulas": special, "unsupported_features": self.blockers(), "financial_correctness": "not_evaluated"}


def validate_pack(pack: dict, package: Package, *, allow_proposal: bool = False) -> list[dict]:
    if not isinstance(pack, dict) or type(pack.get("schema_version")) is not int or pack["schema_version"] != 1:
        raise ModelError("Unsupported model-pack version")
    digest = sha(package.data)
    if pack.get("template_sha256") != digest:
        raise ModelError("Template hash does not match the approved pack")
    pack_status = pack.get("status", "approved")
    approval = pack.get("approval")
    if pack_status == "proposal" or approval is None:
        if not allow_proposal:
            raise ModelError("Pack is an unapproved proposal; independent human approval is required (or pass --allow-proposal for non-ready candidate testing)")
    else:
        if not isinstance(approval, dict) or approval.get("template_sha256") != digest or not all(isinstance(approval.get(k), str) and approval[k].strip() for k in ("reviewer", "approved_at")):
            raise ModelError("Pack requires recorded approval tied to this template hash")
        reviewer_words = set(re.findall(r"[a-z0-9/]+", approval["reviewer"].lower()))
        if reviewer_words & DISALLOWED_REVIEWER_TOKENS:
            raise ModelError(f"Pack approval reviewer cannot be an agent or placeholder: {approval['reviewer']!r}")
        try:
            approved_at = datetime.fromisoformat(approval["approved_at"].replace("Z", "+00:00"))
            if approved_at.tzinfo is None or "T" not in approval["approved_at"]:
                raise ValueError("Timezone is required")
        except ValueError as exc:
            raise ModelError("Approval time must be an ISO datetime with timezone") from exc
    fields = pack.get("inputs")
    checks = pack.get("checks")
    if not isinstance(fields, list) or not fields or not isinstance(checks, list) or not checks:
        raise ModelError("Pack requires input mappings and model-specific checks")
    names, destinations = set(), set()
    for field in fields:
        if not isinstance(field, dict):
            raise ModelError("Invalid input mapping")
        name, sheet, ref = field.get("name"), field.get("sheet"), field.get("cell")
        if not isinstance(name, str) or not name.strip() or name in names:
            raise ModelError("Input names must be non-empty and unique")
        coordinate(ref)
        if not isinstance(sheet, str) or sheet not in package.sheets or (sheet, ref) in destinations:
            raise ModelError("Input destination missing or duplicated")
        if field.get("kind") != "number" or not isinstance(field.get("unit"), str) or not field["unit"].strip():
            raise ModelError("Input mappings require numeric kind and explicit units")
        low = numeric(field["minimum"]) if "minimum" in field else None
        high = numeric(field["maximum"]) if "maximum" in field else None
        if low is not None and high is not None and low > high:
            raise ModelError("Reversed input bounds")
        names.add(name)
        destinations.add((sheet, ref))
    check_names = set()
    for check in checks:
        if not isinstance(check, dict) or not isinstance(check.get("name"), str) or not check["name"].strip() or check["name"] in check_names:
            raise ModelError("Checks require unique names")
        check_names.add(check["name"])
        coordinate(check.get("cell"))
        if check.get("sheet") not in package.sheets:
            raise ModelError("Check sheet is missing")
        numeric(check.get("expected"))
        if numeric(check.get("abs_tolerance")) < 0:
            raise ModelError("Check tolerance must be nonnegative")
        if "rel_tolerance" in check and numeric(check.get("rel_tolerance")) < 0:
            raise ModelError("Check relative tolerance must be nonnegative")
        tree = package.sheets[check["sheet"]][1]
        if not any(c.get("r") == check["cell"] for c in tree.iter(f"{{{NS}}}c")):
            raise ModelError("Check cell is missing")
    return fields


def input_cells(package: Package, fields: list[dict], supplied: dict) -> tuple[dict, list[dict]]:
    if not isinstance(supplied, dict) or set(supplied) != {"values"} or not isinstance(supplied["values"], dict):
        raise ModelError("Inputs must contain exactly a values object")
    values = supplied["values"]
    if set(values) != {field["name"] for field in fields}:
        raise ModelError("Input names must exactly match the approved map")
    edits, audit = {}, []
    for field in fields:
        item = values[field["name"]]
        if not isinstance(item, dict) or not {"value", "source"}.issubset(set(item)) or set(item) - {"value", "source", "unit", "period"} or not isinstance(item["source"], str) or not item["source"].strip():
            raise ModelError("Every input requires a number and a non-empty source reference")
        if "unit" in item and item["unit"] != field["unit"]:
            raise ModelError(f"Input unit mismatch for {field['name']}: expected {field['unit']!r}, got {item['unit']!r}")
        if "period" in item and "period" in field and item["period"] != field["period"]:
            raise ModelError(f"Input period mismatch for {field['name']}: expected {field['period']!r}, got {item['period']!r}")
        number = numeric(item["value"])
        if "minimum" in field and number < numeric(field["minimum"]) or "maximum" in field and number > numeric(field["maximum"]):
            raise ModelError(f"Input outside approved bounds: {field['name']}")
        if (field["sheet"], field["cell"]) in package.formula_cells:
            raise ModelError(f"Mapped input intersects a formula or formula range: {field['sheet']}!{field['cell']}")
        part, tree = package.sheets[field["sheet"]]
        cell = next((c for c in tree.iter(f"{{{NS}}}c") if c.get("r") == field["cell"]), None)
        if cell is None or cell.find(f"{{{NS}}}f") is not None or cell.get("t", "n") != "n":
            raise ModelError("Input must be an existing numeric non-formula cell")
        if any(child.tag != f"{{{NS}}}v" for child in cell):
            raise ModelError("Input cell contains unsupported native content")
        col, row = coordinate(field["cell"])
        for merge in tree.findall(f"{{{NS}}}mergeCells/{{{NS}}}mergeCell"):
            bounds = (merge.get("ref") or "").split(":")
            a = coordinate(bounds[0])
            b = coordinate(bounds[-1])
            if a[0] <= col <= b[0] and a[1] <= row <= b[1]:
                raise ModelError("Mapped input intersects a merged range")
        edits.setdefault(part, {})[field["cell"]] = str(number).encode("ascii")
        audit_entry = {"name": field["name"], "sheet": field["sheet"], "cell": field["cell"], "value_decimal": str(number), "unit": field["unit"], "source": item["source"]}
        if "period" in field:
            audit_entry["period"] = field["period"]
        if "role" in field:
            audit_entry["role"] = field["role"]
        audit.append(audit_entry)
    return edits, audit


def edit_sheet(data: bytes, edits: dict[str, bytes], followers: set[str] | None = None) -> bytes:
    if b"<worksheet" not in data or b"<!--" in data or b"<![CDATA[" in data:
        raise ModelError("Nonstandard worksheet XML requires a dedicated adapter")
    sheet_data = list(re.finditer(rb"<sheetData\b[^>]*?/>|<sheetData\b[^>]*>.*?</sheetData>", data, re.S))
    if len(sheet_data) != 1:
        raise ModelError("Worksheet must contain exactly one supported sheetData element")
    region = sheet_data[0]
    if re.search(rb"<[A-Za-z_][\w.-]*:(?:c|f|v)\b", region[0]):
        raise ModelError("Nonstandard worksheet cell XML requires a dedicated adapter")
    found = set()
    follower_set = followers or set()

    def replace(match):
        fragment = match[0]
        opening_end = fragment.find(b">") + 1
        opening = fragment[:opening_end]
        reference = attribute(opening, "r")
        ref = reference.decode("ascii") if reference else ""
        if ref in edits:
            found.add(ref)
            value = b"<v>" + edits[ref] + b"</v>"
            if fragment.endswith(b"/>"):
                return fragment[:-2] + b">" + value + b"</c>"
            if VALUE.search(fragment):
                return VALUE.sub(lambda _: value, fragment)
            return fragment[:-4] + value + b"</c>"
        if re.search(rb"<f(?:\s|/?>)", fragment) or ref in follower_set:
            opening = set_attribute(opening, "t", None)
            return opening + VALUE.sub(b"", fragment[opening_end:])
        return fragment

    replacement = CELL.sub(replace, region[0])
    result = data[:region.start()] + replacement + data[region.end():]
    if found != set(edits):
        raise ModelError("Could not locate every mapped cell in supported worksheet XML")
    return result


def verify_sheet_changes(original: bytes, candidate: bytes, edited: set[str], followers: set[str] | None = None) -> None:
    """Compare all worksheet XML semantics except explicitly allowed value/cache edits."""
    if FORMULA.findall(original) != FORMULA.findall(candidate):
        raise ModelError("Formula XML was modified during sheet editing")
    follower_set = followers or set()
    trees = [deepcopy(xml(original)), deepcopy(xml(candidate))]
    for index, tree in enumerate(trees):
        for cell in tree.findall(f"{{{NS}}}sheetData/{{{NS}}}row/{{{NS}}}c"):
            is_formula = cell.find(f"{{{NS}}}f") is not None or cell.get("r") in follower_set
            if index == 1 and is_formula and cell.find(f"{{{NS}}}v") is not None:
                raise ModelError("A formula cache survived invalidation")
            if is_formula or cell.get("r") in edited:
                for value in list(cell.findall(f"{{{NS}}}v")):
                    cell.remove(value)
            if is_formula:
                cell.attrib.pop("t", None)
        for node in tree.iter():
            if node.tail and not node.tail.strip():
                node.tail = None
            if len(node) and node.text and not node.text.strip():
                node.text = None
    if ET.tostring(trees[0]) != ET.tostring(trees[1]):
        raise ModelError("Unapproved worksheet change detected")


def calculation_flags(data: bytes) -> bytes:
    if b"<workbook" not in data:
        raise ModelError("Prefixed workbook XML requires a dedicated adapter")
    pattern = re.compile(
        rb"<(?:[A-Za-z_][\w.-]*:)?calcPr\b[^>]*?/>|<(?:[A-Za-z_][\w.-]*:)?calcPr\b[^>]*>.*?</(?:[A-Za-z_][\w.-]*:)?calcPr>",
        re.S,
    )
    matches = list(pattern.finditer(data))
    if len(matches) > 1:
        raise ModelError("Duplicate calculation properties")
    if matches:
        tag = matches[0][0]
        if not tag.endswith(b"/>"):
            raise ModelError("Unsupported calculation-properties shape")
        for key, value in (("calcMode", b"auto"), ("fullCalcOnLoad", b"1"), ("forceFullCalc", b"1")):
            tag = set_attribute(tag, key, value)
        return data[:matches[0].start()] + tag + data[matches[0].end():]
    tag = b'<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
    successor = re.search(rb"<(?:[A-Za-z_][\w.-]*:)?(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|</(?:[A-Za-z_][\w.-]*:)?workbook>", data)
    if not successor:
        raise ModelError("Could not insert calculation properties")
    return data[:successor.start()] + tag + data[successor.start():]


def without_chain_entry(data: bytes, element: str, key: str, value: bytes) -> bytes:
    pattern = re.compile(rb"<" + element.encode() + rb"\b[^>]*/>", re.S)
    return pattern.sub(lambda m: b"" if attribute(m[0], key) == value else m[0], data)


def compose(template: bytes, pack: dict, supplied: dict, *, allow_proposal: bool = False) -> tuple[bytes, dict]:
    package = Package(template)
    blockers = package.blockers()
    if blockers:
        raise ModelError("; ".join(blockers))
    fields = validate_pack(pack, package, allow_proposal=allow_proposal)
    edits, audit = input_cells(package, fields, supplied)
    changed = dict(package.parts)
    for sheet_name, (part, _) in package.sheets.items():
        followers = package.formula_followers.get(sheet_name, set())
        changed[part] = edit_sheet(changed[part], edits.get(part, {}), followers)
        verify_sheet_changes(package.parts[part], changed[part], set(edits.get(part, {})), followers)
    changed["xl/workbook.xml"] = calculation_flags(changed["xl/workbook.xml"])
    changed.pop("xl/calcChain.xml", None)
    rel_name = "xl/_rels/workbook.xml.rels"
    changed[rel_name] = without_chain_entry(changed[rel_name], "Relationship", "Type", (REL + "/calcChain").encode())
    changed["[Content_Types].xml"] = without_chain_entry(changed["[Content_Types].xml"], "Override", "PartName", b"/xl/calcChain.xml")
    # Abort if a nonstandard declaration would leave a dangling calcChain reference.
    if any((r.get("Type") or "").endswith("/calcChain") for r in xml(changed[rel_name])):
        raise ModelError("Unsupported calculation-chain relationship shape")
    if any(r.get("PartName") == "/xl/calcChain.xml" for r in xml(changed["[Content_Types].xml"])):
        raise ModelError("Unsupported calculation-chain content-type shape")
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for entry in package.entries:
            if entry.filename in changed:
                archive.writestr(copy(entry), changed[entry.filename])
    output = stream.getvalue()
    # Reparse generated XML and verify every requested input survived as a number.
    result = Package(output)
    for field, event in zip(fields, audit):
        tree = result.sheets[field["sheet"]][1]
        cell = next(c for c in tree.iter(f"{{{NS}}}c") if c.get("r") == field["cell"])
        if Decimal(cell.findtext(f"{{{NS}}}v")) != Decimal(event["value_decimal"]):
            raise ModelError("Post-write input verification failed")
    is_proposal = pack.get("status") == "proposal" or pack.get("approval") is None
    pending_gates = ["engine_recalculation", "formula_error_scan", "independent_model_checks", "input_perturbation_and_restoration", "render_review"]
    if is_proposal:
        pending_gates.insert(0, "human_pack_approval")
    manifest = {
        "status": "needs_recalculation", "ready": False,
        "pack_status": "proposal" if is_proposal else "approved",
        "template_sha256": sha(template), "output_sha256": sha(output),
        "pack_sha256": sha(json.dumps(pack, sort_keys=True, default=str).encode()),
        "input_sha256": sha(json.dumps(supplied, sort_keys=True, default=str).encode()),
        "changed_inputs": audit,
        "changed_parts": sorted(k for k, v in changed.items() if v != package.parts.get(k)),
        "removed_parts": sorted(set(package.parts) - set(changed)),
        "pending_gates": pending_gates,
    }
    return output, manifest


def read_json(path: str) -> dict:
    def unique_object(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ModelError(f"Duplicate JSON key: {key}")
            result[key] = value
        return result

    def invalid_constant(value):
        raise ModelError(f"Invalid JSON numeric constant: {value}")

    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"), parse_float=Decimal, object_pairs_hook=unique_object, parse_constant=invalid_constant)
    except (json.JSONDecodeError, UnicodeError) as exc:
        raise ModelError("Invalid JSON input") from exc


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    inspect = commands.add_parser("inspect")
    inspect.add_argument("template")
    build = commands.add_parser("build")
    for name in ("template", "pack", "inputs", "output"):
        build.add_argument("--" + name, required=True)
    build.add_argument("--allow-proposal", action="store_true", help="Allow building a non-ready candidate from an unapproved proposal pack")
    validate = commands.add_parser("validate")
    validate.add_argument("workbook")
    validate.add_argument("--pack", required=True)
    validate.add_argument("--recalculate", action="store_true")
    validate.add_argument("--recalculated-output")
    validate.add_argument("--perturb-check", action="store_true")
    validate.add_argument("--template")
    validate.add_argument("--inputs")
    validate.add_argument("--report")
    validate.add_argument("--allow-proposal", action="store_true", help="Allow validating a candidate against an unapproved proposal pack")
    args = parser.parse_args()
    try:
        if args.command == "inspect":
            data = Path(args.template).read_bytes()
            print(json.dumps(Package(data).summary(), indent=2))
        elif args.command == "build":
            data = Path(args.template).read_bytes()
            if Path(args.template).resolve() == Path(args.output).resolve():
                raise ModelError("Output must be different from the original template")
            output, manifest = compose(data, read_json(args.pack), read_json(args.inputs), allow_proposal=args.allow_proposal)
            with Path(args.output).open("xb") as destination:
                destination.write(output)
            print(json.dumps(manifest, indent=2))
        elif args.command == "validate":
            from validate_workbook import validate_candidate

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
        return 0
    except (ModelError, OSError, KeyError, TypeError) as exc:
        print(json.dumps({"status": "blocked", "ready": False, "error": str(exc)}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

