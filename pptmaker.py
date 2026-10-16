"""Dependency-free OOXML PPTX inspector and template-preserving presentation composer."""
from __future__ import annotations

import argparse
from copy import copy, deepcopy
from datetime import datetime
from decimal import Decimal, InvalidOperation
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import posixpath
from typing import Any
import xml.etree.ElementTree as ET
import zipfile
import zlib

from modelmaker import ModelError, NS as XLSX_NS, Package as WorkbookPackage, numeric, read_json, sha

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

MAX_PPTX_SIZE = 300 * 1024 * 1024
MAX_PPTX_ENTRIES = 10000


class DeckError(ValueError):
    """An input, deck pack, or PPTX feature is not supported safely."""


def parse_xml(data: bytes) -> ET.Element:
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise DeckError("DTD/entity declarations are forbidden in PPTX XML")
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise DeckError("Malformed PPTX XML") from exc


def resolve_rel_target(base_part: str, target: str) -> str:
    if not target or "\\" in target:
        raise DeckError(f"Unsupported relationship target: {target!r}")
    if target.startswith("/"):
        norm = posixpath.normpath(target.lstrip("/"))
    else:
        base_dir = posixpath.dirname(base_part)
        norm = posixpath.normpath(posixpath.join(base_dir, target))
    if norm.startswith("../") or norm == ".." or "/../" in norm:
        raise DeckError(f"Relationship target escapes archive root: {target!r}")
    return norm


def _paragraph_text(p_elem: ET.Element) -> str:
    return "".join(t.text or "" for t in p_elem.findall(f".//{{{A_NS}}}t"))


def _txbody_paragraphs(tx_body: ET.Element) -> list[str]:
    return [_paragraph_text(p) for p in tx_body.findall(f"{{{A_NS}}}p")]


def _txbody_full_text(tx_body: ET.Element) -> str:
    return "\n".join(_txbody_paragraphs(tx_body))


def _set_paragraph_text_preserving_rpr(p_elem: ET.Element, new_text: str) -> None:
    """Replace text inside <a:p> while preserving <a:pPr>, first <a:rPr>, and <a:endParaRPr>."""
    p_pr = p_elem.find(f"{{{A_NS}}}pPr")
    end_rpr = p_elem.find(f"{{{A_NS}}}endParaRPr")
    first_r = p_elem.find(f"{{{A_NS}}}r")
    first_rpr = deepcopy(first_r.find(f"{{{A_NS}}}rPr")) if first_r is not None and first_r.find(f"{{{A_NS}}}rPr") is not None else None
    p_pr_copy = deepcopy(p_pr) if p_pr is not None else None
    end_rpr_copy = deepcopy(end_rpr) if end_rpr is not None else None

    for child in list(p_elem):
        p_elem.remove(child)

    if p_pr_copy is not None:
        p_elem.append(p_pr_copy)
    r_elem = ET.SubElement(p_elem, f"{{{A_NS}}}r")
    if first_rpr is not None:
        r_elem.append(first_rpr)
    t_elem = ET.SubElement(r_elem, f"{{{A_NS}}}t")
    t_elem.text = new_text
    if end_rpr_copy is not None:
        p_elem.append(end_rpr_copy)


def _set_txbody_text_preserving_format(
    tx_body: ET.Element,
    new_text: str,
    paragraph_index: int | None = None,
) -> None:
    paragraphs = tx_body.findall(f"{{{A_NS}}}p")
    if not paragraphs:
        raise DeckError("Target text body has no paragraphs")
    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            raise DeckError(f"Paragraph index {paragraph_index} out of range (0..{len(paragraphs) - 1})")
        _set_paragraph_text_preserving_rpr(paragraphs[paragraph_index], new_text)
        return

    lines = new_text.split("\n")
    template_p = deepcopy(paragraphs[0])
    for p in paragraphs:
        tx_body.remove(p)
    for idx, line in enumerate(lines):
        if idx < len(paragraphs):
            p_node = deepcopy(paragraphs[idx])
        else:
            p_node = deepcopy(template_p)
        _set_paragraph_text_preserving_rpr(p_node, line)
        tx_body.append(p_node)


def _replace_token_in_txbody(tx_body: ET.Element, token: str, replacement: str) -> int:
    """Replace token across runs or within paragraphs while preserving <a:rPr>."""
    if not token:
        raise DeckError("Token must be non-empty")
    count = 0
    for p_elem in tx_body.findall(f"{{{A_NS}}}p"):
        runs = p_elem.findall(f"{{{A_NS}}}r")
        # First try run-local replacement (preserves multi-run rich styling inside the same paragraph)
        run_replaced = 0
        for r in runs:
            t = r.find(f"{{{A_NS}}}t")
            if t is not None and t.text and token in t.text:
                occ = t.text.count(token)
                t.text = t.text.replace(token, replacement)
                run_replaced += occ
        if run_replaced > 0:
            count += run_replaced
            continue
        # If token spans across adjacent runs in the paragraph, coalesce paragraph text
        full_p = _paragraph_text(p_elem)
        if token in full_p:
            occ = full_p.count(token)
            _set_paragraph_text_preserving_rpr(p_elem, full_p.replace(token, replacement))
            count += occ
    return count


class PptxPackage:
    def __init__(self, data: bytes):
        if len(data) > MAX_PPTX_SIZE:
            raise DeckError("Presentation exceeds the supported size")
        self.data = data
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                entries = archive.infolist()
                names = [entry.filename for entry in entries]
                if len(entries) > MAX_PPTX_ENTRIES or len(set(names)) != len(names):
                    raise DeckError("Duplicate or excessive ZIP entries in PPTX")
                if sum(entry.file_size for entry in entries) > MAX_PPTX_SIZE:
                    raise DeckError("Expanded PPTX exceeds the supported size")
                for entry in entries:
                    path = PurePosixPath(entry.filename)
                    if path.is_absolute() or ".." in path.parts or "\\" in entry.filename:
                        raise DeckError("Unsafe ZIP member name in PPTX")
                    if entry.flag_bits & 1:
                        raise DeckError("Encrypted ZIP entries are unsupported in PPTX")
                self.entries = entries
                self.parts = {entry.filename: archive.read(entry) for entry in entries}
        except (zipfile.BadZipFile, RuntimeError, NotImplementedError, EOFError, zlib.error) as exc:
            raise DeckError("Unreadable or unsupported PPTX archive") from exc

        for required in ("[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"):
            if required not in self.parts:
                raise DeckError(f"Required PPTX part missing: {required}")

        for name, value in self.parts.items():
            if name.endswith((".xml", ".rels")):
                parse_xml(value)

        self.presentation = parse_xml(self.parts["ppt/presentation.xml"])
        if self.presentation.tag != f"{{{P_NS}}}presentation":
            raise DeckError("Invalid presentation root element")

        pres_rels = parse_xml(self.parts["ppt/_rels/presentation.xml.rels"])
        rel_map = {}
        for rel in pres_rels:
            rid = rel.get("Id")
            if rid in rel_map:
                raise DeckError("Duplicate presentation relationship ID")
            rel_map[rid] = rel

        self.slides: list[tuple[int, str, ET.Element]] = []
        sld_id_lst = self.presentation.find(f"{{{P_NS}}}sldIdLst")
        if sld_id_lst is not None:
            for idx, sld_id in enumerate(sld_id_lst.findall(f"{{{P_NS}}}sldId"), start=1):
                rid = sld_id.get(f"{{{R_NS}}}id")
                rel = rel_map.get(rid)
                if rel is None or rel.get("TargetMode") == "External":
                    raise DeckError(f"Missing or external slide relationship for slide {idx}")
                part = resolve_rel_target("ppt/presentation.xml", rel.get("Target", ""))
                if part not in self.parts:
                    raise DeckError(f"Slide relationship points to missing part: {part}")
                tree = parse_xml(self.parts[part])
                if tree.tag != f"{{{P_NS}}}sld":
                    raise DeckError(f"Unexpected slide XML root in {part}")
                self.slides.append((idx, part, tree))
        if not self.slides:
            raise DeckError("Presentation has no slides")

    def blockers(self) -> list[str]:
        blockers = []
        for name in self.parts:
            lower = name.lower()
            if "vbaproject" in lower or "_xmlsignatures/" in lower:
                blockers.append(f"Unsupported macro/signature part: {name}")
        return sorted(set(blockers))

    def _slide_layout_and_charts(self, slide_part: str) -> tuple[str | None, list[str]]:
        rels_part = posixpath.join(posixpath.dirname(slide_part), "_rels", posixpath.basename(slide_part) + ".rels")
        layout = None
        charts = []
        if rels_part in self.parts:
            rels_root = parse_xml(self.parts[rels_part])
            for rel in rels_root:
                rtype = rel.get("Type") or ""
                if rtype.endswith("/slideLayout") and rel.get("Target"):
                    layout = resolve_rel_target(slide_part, rel.get("Target", ""))
                elif rtype.endswith("/chart") and rel.get("Target"):
                    charts.append(resolve_rel_target(slide_part, rel.get("Target", "")))
        return layout, sorted(charts)

    def summary(self) -> dict[str, Any]:
        layouts = [n for n in self.parts if n.startswith("ppt/slideLayouts/slideLayout") and n.endswith(".xml")]
        masters = [n for n in self.parts if n.startswith("ppt/slideMasters/slideMaster") and n.endswith(".xml")]
        charts = [n for n in self.parts if n.startswith("ppt/charts/chart") and n.endswith(".xml")]
        media = [n for n in self.parts if n.startswith("ppt/media/")]

        slides_info = []
        for s_num, s_part, tree in self.slides:
            layout, slide_charts = self._slide_layout_and_charts(s_part)
            text_slots = []
            for sp in tree.findall(f".//{{{P_NS}}}sp"):
                c_nv_pr = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
                tx_body = sp.find(f"{{{P_NS}}}txBody")
                if c_nv_pr is None or tx_body is None:
                    continue
                paras = _txbody_paragraphs(tx_body)
                full_text = "\n".join(p for p in paras if p.strip())
                if full_text:
                    text_slots.append({
                        "shape_id": int(c_nv_pr.get("id", "0")),
                        "shape_name": c_nv_pr.get("name", ""),
                        "paragraph_count": len(paras),
                        "text": full_text,
                    })

            tables_info = []
            for gf in tree.findall(f".//{{{P_NS}}}graphicFrame"):
                tbl = gf.find(f".//{{{A_NS}}}tbl")
                if tbl is None:
                    continue
                c_nv_pr = gf.find(f"{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr")
                rows = tbl.findall(f"{{{A_NS}}}tr")
                matrix = []
                for tr in rows:
                    tcs = tr.findall(f"{{{A_NS}}}tc")
                    matrix.append([
                        "\n".join(p for p in _txbody_paragraphs(tc.find(f"{{{A_NS}}}txBody")) if p.strip())
                        if tc.find(f"{{{A_NS}}}txBody") is not None
                        else ""
                        for tc in tcs
                    ])
                tables_info.append({
                    "shape_id": int(c_nv_pr.get("id", "0")) if c_nv_pr is not None else 0,
                    "shape_name": c_nv_pr.get("name", "") if c_nv_pr is not None else "",
                    "rows": len(matrix),
                    "cols": max((len(r) for r in matrix), default=0),
                    "cells": matrix,
                })

            slides_info.append({
                "slide": s_num,
                "part": s_part,
                "layout": layout,
                "text_slot_count": len(text_slots),
                "table_count": len(tables_info),
                "charts": slide_charts,
                "text_slots": text_slots,
                "tables": tables_info,
            })

        return {
            "sha256": sha(self.data),
            "slide_count": len(self.slides),
            "layout_count": len(layouts),
            "master_count": len(masters),
            "chart_count": len(charts),
            "media_count": len(media),
            "unsupported_features": self.blockers(),
            "slides": slides_info,
        }


def _find_shape_on_slide(slide_tree: ET.Element, shape_id: int, shape_name: str | None = None) -> ET.Element:
    for sp in slide_tree.findall(f".//{{{P_NS}}}sp"):
        c_nv_pr = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
        if c_nv_pr is None:
            continue
        if int(c_nv_pr.get("id", "-1")) == shape_id:
            if shape_name is not None and c_nv_pr.get("name") != shape_name:
                raise DeckError(f"Shape id={shape_id} has name {c_nv_pr.get('name')!r}, expected {shape_name!r}")
            return sp
    raise DeckError(f"Shape id={shape_id} not found on slide")


def _find_table_on_slide(slide_tree: ET.Element, shape_id: int, shape_name: str | None = None) -> ET.Element:
    for gf in slide_tree.findall(f".//{{{P_NS}}}graphicFrame"):
        tbl = gf.find(f".//{{{A_NS}}}tbl")
        if tbl is None:
            continue
        c_nv_pr = gf.find(f"{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr")
        if c_nv_pr is not None and int(c_nv_pr.get("id", "-1")) == shape_id:
            if shape_name is not None and c_nv_pr.get("name") != shape_name:
                raise DeckError(f"Table id={shape_id} has name {c_nv_pr.get('name')!r}, expected {shape_name!r}")
            return tbl
    raise DeckError(f"Table id={shape_id} not found on slide")


def validate_deck_pack(pack: dict, package: PptxPackage) -> list[dict]:
    if not isinstance(pack, dict) or type(pack.get("schema_version")) is not int or pack["schema_version"] != 1:
        raise DeckError("Unsupported deck-pack schema_version")
    digest = sha(package.data)
    if pack.get("template_sha256") != digest:
        raise DeckError("Template SHA-256 does not match the approved deck pack")
    approval = pack.get("approval")
    if (
        not isinstance(approval, dict)
        or approval.get("template_sha256") != digest
        or not all(isinstance(approval.get(k), str) and approval[k].strip() for k in ("reviewer", "approved_at"))
    ):
        raise DeckError("Deck pack requires recorded approval tied to this template hash")
    try:
        approved_at = datetime.fromisoformat(approval["approved_at"].replace("Z", "+00:00"))
        if approved_at.tzinfo is None or "T" not in approval["approved_at"]:
            raise ValueError("Timezone is required")
    except ValueError as exc:
        raise DeckError("Approval time must be an ISO datetime with timezone") from exc

    slots = pack.get("slots")
    if not isinstance(slots, list) or not slots:
        raise DeckError("Deck pack requires a non-empty slots list")

    slide_map = {s_num: tree for s_num, _, tree in package.slides}
    names = set()
    targets = set()
    for slot in slots:
        if not isinstance(slot, dict):
            raise DeckError("Invalid slot definition")
        name = slot.get("name")
        if not isinstance(name, str) or not name.strip() or name in names:
            raise DeckError(f"Slot names must be non-empty and unique: {name!r}")
        names.add(name)

        slide_num = slot.get("slide")
        if type(slide_num) is not int or slide_num not in slide_map:
            raise DeckError(f"Invalid slide number in slot {name!r}: {slide_num!r}")
        slide_tree = slide_map[slide_num]

        target_type = slot.get("target_type")
        shape_id = slot.get("shape_id")
        shape_name = slot.get("shape_name")
        if type(shape_id) is not int:
            raise DeckError(f"Slot {name!r} requires integer shape_id")

        if target_type == "shape_text":
            sp = _find_shape_on_slide(slide_tree, shape_id, shape_name)
            tx_body = sp.find(f"{{{P_NS}}}txBody")
            if tx_body is None:
                raise DeckError(f"Shape {shape_id} on slide {slide_num} has no text body")
            p_idx = slot.get("paragraph_index")
            if p_idx is not None:
                paras = tx_body.findall(f"{{{A_NS}}}p")
                if type(p_idx) is not int or p_idx < 0 or p_idx >= len(paras):
                    raise DeckError(f"Invalid paragraph_index {p_idx!r} in slot {name!r}")
            t_key = (slide_num, "shape_text", shape_id, p_idx)
        elif target_type == "shape_token":
            token = slot.get("token")
            if not isinstance(token, str) or not token:
                raise DeckError(f"Slot {name!r} of type shape_token requires non-empty token")
            sp = _find_shape_on_slide(slide_tree, shape_id, shape_name)
            tx_body = sp.find(f"{{{P_NS}}}txBody")
            if tx_body is None or token not in _txbody_full_text(tx_body):
                raise DeckError(f"Token {token!r} not found in shape {shape_id} on slide {slide_num}")
            t_key = (slide_num, "shape_token", shape_id, token)
        elif target_type == "table_cell":
            row_idx = slot.get("row")
            col_idx = slot.get("col")
            if type(row_idx) is not int or type(col_idx) is not int or row_idx < 0 or col_idx < 0:
                raise DeckError(f"Slot {name!r} of type table_cell requires non-negative integer row and col")
            tbl = _find_table_on_slide(slide_tree, shape_id, shape_name)
            rows = tbl.findall(f"{{{A_NS}}}tr")
            if row_idx >= len(rows):
                raise DeckError(f"Row {row_idx} out of range for table {shape_id} on slide {slide_num}")
            tcs = rows[row_idx].findall(f"{{{A_NS}}}tc")
            if col_idx >= len(tcs):
                raise DeckError(f"Col {col_idx} out of range for table {shape_id} on slide {slide_num}")
            t_key = (slide_num, "table_cell", shape_id, row_idx, col_idx)
        else:
            raise DeckError(f"Unsupported target_type in slot {name!r}: {target_type!r}")

        if t_key in targets:
            raise DeckError(f"Duplicate slot target destination: {t_key}")
        targets.add(t_key)

    return slots


def _resolve_workbook_binding(wb_pkg: WorkbookPackage | None, binding: dict, slot_name: str) -> tuple[Decimal, str]:
    if wb_pkg is None:
        raise DeckError(f"Slot {slot_name!r} has workbook_binding, so a recalculated --workbook must be provided")
    sheet = binding.get("sheet")
    cell_ref = binding.get("cell")
    if not isinstance(sheet, str) or sheet not in wb_pkg.sheets or not isinstance(cell_ref, str):
        raise DeckError(f"Invalid workbook_binding in slot {slot_name!r}")
    tree = wb_pkg.sheets[sheet][1]
    cell = next((c for c in tree.iter(f"{{{XLSX_NS}}}c") if c.get("r") == cell_ref), None)
    if cell is None:
        raise DeckError(f"Bound cell {sheet}!{cell_ref} does not exist in workbook")
    if cell.get("t") == "e":
        raise DeckError(f"Bound cell {sheet}!{cell_ref} has cached Excel error: {cell.findtext(f'{{{XLSX_NS}}}v')}")
    raw_v = cell.findtext(f"{{{XLSX_NS}}}v")
    if raw_v is None or raw_v.strip() == "":
        raise DeckError(f"Bound cell {sheet}!{cell_ref} has no cached value (workbook must be recalculated first)")
    try:
        val = Decimal(raw_v)
    except InvalidOperation as exc:
        raise DeckError(f"Bound cell {sheet}!{cell_ref} is not numeric: {raw_v!r}") from exc

    scale = Decimal(str(binding.get("scale", 1)))
    scaled = val * scale
    decimals = binding.get("decimals")
    if decimals is not None:
        quant = Decimal("10") ** (-int(decimals))
        formatted_num = f"{scaled.quantize(quant):,}" if binding.get("thousands_sep") else str(scaled.quantize(quant))
    else:
        formatted_num = str(scaled)
    prefix = binding.get("prefix", "")
    suffix = binding.get("suffix", "")
    return scaled, f"{prefix}{formatted_num}{suffix}"


def verify_slide_changes(original: bytes, candidate: bytes, approved_shape_ids: set[int]) -> None:
    """Ensure no unapproved shapes, styles, or run formatting properties (<a:rPr>) were altered."""
    trees = [deepcopy(parse_xml(original)), deepcopy(parse_xml(candidate))]
    for tree in trees:
        for sp in tree.findall(f".//{{{P_NS}}}sp"):
            c_nv_pr = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if c_nv_pr is not None and int(c_nv_pr.get("id", "-1")) in approved_shape_ids:
                tx_body = sp.find(f"{{{P_NS}}}txBody")
                if tx_body is not None:
                    # Preserve checks on nvSpPr and spPr; normalize <a:p> children inside approved txBody
                    for p in list(tx_body.findall(f"{{{A_NS}}}p")):
                        tx_body.remove(p)
        for gf in tree.findall(f".//{{{P_NS}}}graphicFrame"):
            c_nv_pr = gf.find(f"{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr")
            if c_nv_pr is not None and int(c_nv_pr.get("id", "-1")) in approved_shape_ids:
                for tc in gf.findall(f".//{{{A_NS}}}tc"):
                    tx_body = tc.find(f"{{{A_NS}}}txBody")
                    if tx_body is not None:
                        for p in list(tx_body.findall(f"{{{A_NS}}}p")):
                            tx_body.remove(p)
        for node in tree.iter():
            if node.tail and not node.tail.strip():
                node.tail = None
            if len(node) and node.text and not node.text.strip():
                node.text = None
    if ET.tostring(trees[0]) != ET.tostring(trees[1]):
        raise DeckError("Unapproved slide structural or formatting change detected")


def compose_deck(
    template: bytes,
    pack: dict,
    supplied: dict,
    workbook_bytes: bytes | None = None,
) -> tuple[bytes, dict[str, Any]]:
    package = PptxPackage(template)
    blockers = package.blockers()
    if blockers:
        raise DeckError("; ".join(blockers))
    slots = validate_deck_pack(pack, package)

    if not isinstance(supplied, dict) or set(supplied) != {"values"} or not isinstance(supplied["values"], dict):
        raise DeckError("Deck inputs must contain exactly a values object")
    values = supplied["values"]
    if set(values) != {s["name"] for s in slots}:
        raise DeckError("Deck input names must exactly match the approved deck-pack slots")

    wb_pkg = WorkbookPackage(workbook_bytes) if workbook_bytes is not None else None
    slide_part_map = {s_num: (s_part, deepcopy(tree)) for s_num, s_part, tree in package.slides}
    approved_shapes_per_slide: dict[int, set[int]] = {}
    audit: list[dict[str, Any]] = []

    for slot in slots:
        name = slot["name"]
        item = values[name]
        if (
            not isinstance(item, dict)
            or set(item) != {"value", "source"}
            or not isinstance(item["source"], str)
            or not item["source"].strip()
        ):
            raise DeckError(f"Slot {name!r} requires a value and a non-empty source reference")

        raw_val = item["value"]
        if isinstance(raw_val, bool) or raw_val is None:
            raise DeckError(f"Slot {name!r} value must be a string or finite number")

        binding = slot.get("workbook_binding")
        if binding is not None:
            bound_dec, bound_formatted = _resolve_workbook_binding(wb_pkg, binding, name)
            if isinstance(raw_val, (int, float, Decimal)):
                supplied_dec = numeric(raw_val)
                tol = Decimal(str(binding.get("abs_tolerance", "0.05")))
                if abs(supplied_dec - bound_dec) > tol:
                    raise DeckError(
                        f"Slot {name!r} value {supplied_dec} diverges from recalculated workbook cell "
                        f"{binding['sheet']}!{binding['cell']} ({bound_dec}) beyond tolerance {tol}"
                    )
                text_val = bound_formatted
            elif isinstance(raw_val, str):
                if raw_val.strip() != bound_formatted.strip():
                    raise DeckError(
                        f"Slot {name!r} text {raw_val!r} does not match bound workbook value {bound_formatted!r}"
                    )
                text_val = raw_val
            else:
                raise DeckError(f"Invalid value type for bound slot {name!r}")
        else:
            text_val = str(raw_val)
            if not text_val.strip():
                raise DeckError(f"Slot {name!r} replacement text cannot be empty")

        slide_num = slot["slide"]
        shape_id = slot["shape_id"]
        shape_name = slot.get("shape_name")
        target_type = slot["target_type"]
        _, slide_tree = slide_part_map[slide_num]
        approved_shapes_per_slide.setdefault(slide_num, set()).add(shape_id)

        if target_type == "shape_text":
            sp = _find_shape_on_slide(slide_tree, shape_id, shape_name)
            tx_body = sp.find(f"{{{P_NS}}}txBody")
            _set_txbody_text_preserving_format(tx_body, text_val, slot.get("paragraph_index"))
        elif target_type == "shape_token":
            sp = _find_shape_on_slide(slide_tree, shape_id, shape_name)
            tx_body = sp.find(f"{{{P_NS}}}txBody")
            replaced_count = _replace_token_in_txbody(tx_body, slot["token"], text_val)
            if replaced_count == 0:
                raise DeckError(f"Token {slot['token']!r} could not be replaced in slot {name!r}")
        elif target_type == "table_cell":
            tbl = _find_table_on_slide(slide_tree, shape_id, shape_name)
            tr = tbl.findall(f"{{{A_NS}}}tr")[slot["row"]]
            tc = tr.findall(f"{{{A_NS}}}tc")[slot["col"]]
            tx_body = tc.find(f"{{{A_NS}}}txBody")
            if tx_body is None:
                raise DeckError(f"Table cell ({slot['row']}, {slot['col']}) has no txBody")
            _set_txbody_text_preserving_format(tx_body, text_val, None)

        audit.append({
            "name": name,
            "slide": slide_num,
            "target_type": target_type,
            "shape_id": shape_id,
            "value": text_val,
            "source": item["source"],
            "workbook_binding": binding,
        })

    ET.register_namespace("a", A_NS)
    ET.register_namespace("r", R_NS)
    ET.register_namespace("p", P_NS)
    ET.register_namespace("c", C_NS)

    changed = dict(package.parts)
    for slide_num, shape_ids in approved_shapes_per_slide.items():
        s_part, mod_tree = slide_part_map[slide_num]
        mod_bytes = ET.tostring(mod_tree, encoding="utf-8", xml_declaration=True)
        verify_slide_changes(package.parts[s_part], mod_bytes, shape_ids)
        changed[s_part] = mod_bytes

    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for entry in package.entries:
            if entry.filename in changed:
                archive.writestr(copy(entry), changed[entry.filename])
    output = stream.getvalue()

    # Reparse generated PPTX archive to verify structural validity
    result_pkg = PptxPackage(output)
    if len(result_pkg.slides) != len(package.slides):
        raise DeckError("Post-write slide count mismatch")

    manifest = {
        "status": "composed",
        "ready": True,
        "template_sha256": sha(template),
        "output_sha256": sha(output),
        "pack_sha256": sha(json.dumps(pack, sort_keys=True, default=str).encode()),
        "input_sha256": sha(json.dumps(supplied, sort_keys=True, default=str).encode()),
        "workbook_sha256": sha(workbook_bytes) if workbook_bytes is not None else None,
        "changed_slots": audit,
        "changed_parts": sorted(k for k, v in changed.items() if v != package.parts.get(k)),
        "preserved_part_count": sum(1 for k, v in changed.items() if v == package.parts.get(k)),
    }
    return output, manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    inspect_cmd = commands.add_parser("inspect")
    inspect_cmd.add_argument("template")
    build_cmd = commands.add_parser("build")
    for name in ("template", "pack", "inputs", "output"):
        build_cmd.add_argument("--" + name, required=True)
    build_cmd.add_argument("--workbook", help="Optional recalculated XLSX workbook for bound metrics")
    args = parser.parse_args()
    try:
        data = Path(args.template).read_bytes()
        if args.command == "inspect":
            print(json.dumps(PptxPackage(data).summary(), indent=2))
        else:
            if Path(args.template).resolve() == Path(args.output).resolve():
                raise DeckError("Output must be different from the original template")
            wb_bytes = Path(args.workbook).read_bytes() if args.workbook else None
            output, manifest = compose_deck(
                data,
                read_json(args.pack),
                read_json(args.inputs),
                workbook_bytes=wb_bytes,
            )
            Path(args.output).parent.mkdir(parents=True, exist_ok=True)
            with Path(args.output).open("xb") as destination:
                destination.write(output)
            print(json.dumps(manifest, indent=2, default=str))
        return 0
    except (DeckError, ModelError, OSError, KeyError, TypeError) as exc:
        print(json.dumps({"status": "blocked", "ready": False, "error": str(exc)}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
