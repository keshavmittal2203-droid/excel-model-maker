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
import re
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
PACKAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"

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


def validate_deck_pack(pack: dict, package: PptxPackage, *, allow_proposal: bool = False) -> list[dict]:
    from modelmaker import DISALLOWED_REVIEWER_TOKENS

    if not isinstance(pack, dict) or type(pack.get("schema_version")) is not int or pack["schema_version"] != 1:
        raise DeckError("Unsupported deck-pack schema_version")
    digest = sha(package.data)
    if pack.get("template_sha256") != digest:
        raise DeckError("Template SHA-256 does not match the approved deck pack")
    pack_status = pack.get("status", "approved")
    approval = pack.get("approval")
    if pack_status == "proposal" or approval is None:
        if not allow_proposal:
            raise DeckError("Deck pack is an unapproved proposal; independent human approval is required (or pass --allow-proposal for non-ready candidate testing)")
    else:
        if (
            not isinstance(approval, dict)
            or approval.get("template_sha256") != digest
            or not all(isinstance(approval.get(k), str) and approval[k].strip() for k in ("reviewer", "approved_at"))
        ):
            raise DeckError("Deck pack requires recorded approval tied to this template hash")
        reviewer_words = set(re.findall(r"[a-z0-9/]+", approval["reviewer"].lower()))
        if reviewer_words & DISALLOWED_REVIEWER_TOKENS:
            raise DeckError(f"Deck pack approval reviewer cannot be an agent or placeholder: {approval['reviewer']!r}")
        try:
            approved_at = datetime.fromisoformat(approval["approved_at"].replace("Z", "+00:00"))
            if approved_at.tzinfo is None or "T" not in approval["approved_at"]:
                raise ValueError("Timezone is required")
        except ValueError as exc:
            raise DeckError("Approval time must be an ISO datetime with timezone") from exc

    slots = pack.get("slots")
    if not isinstance(slots, list) or not slots:
        raise DeckError("Deck pack requires a non-empty slots list")

    retained_slides = pack.get("retained_slides")
    if retained_slides is not None:
        if not isinstance(retained_slides, list) or not retained_slides:
            raise DeckError("retained_slides must be a non-empty list of integers")
        all_slide_nums = {s_num for s_num, _, _ in package.slides}
        for s_num in retained_slides:
            if type(s_num) is not int or s_num not in all_slide_nums:
                raise DeckError(f"retained_slides contains invalid slide number: {s_num}")
        if len(set(retained_slides)) != len(retained_slides):
            raise DeckError("Duplicate slide numbers in retained_slides")

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
        if retained_slides is not None and slide_num not in retained_slides:
            raise DeckError(f"Slot {name!r} references pruned slide {slide_num}")
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
        elif target_type == "picture_replacement":
            pic = next((p for p in slide_tree.findall(f".//{{{P_NS}}}pic") if int(p.find(f"{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr").get("id", "-1")) == shape_id), None)
            if pic is None:
                raise DeckError(f"Picture shape {shape_id} not found on slide {slide_num}")
            if shape_name is not None:
                p_name = pic.find(f"{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr").get("name")
                if p_name != shape_name:
                    raise DeckError(f"Picture shape id={shape_id} has name {p_name!r}, expected {shape_name!r}")
            blip = pic.find(f".//{{{A_NS}}}blip")
            if blip is None or not blip.get(f"{{{R_NS}}}embed"):
                raise DeckError(f"Picture shape {shape_id} on slide {slide_num} has no blip embed reference")
            t_key = (slide_num, "picture_replacement", shape_id)
        elif target_type in ("chart_series", "chart_data"):
            gf = next((g for g in slide_tree.findall(f".//{{{P_NS}}}graphicFrame") if int(g.find(f"{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr").get("id", "-1")) == shape_id), None)
            if gf is None:
                raise DeckError(f"GraphicFrame shape {shape_id} not found on slide {slide_num}")
            if shape_name is not None:
                g_name = gf.find(f"{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr").get("name")
                if g_name != shape_name:
                    raise DeckError(f"GraphicFrame shape id={shape_id} has name {g_name!r}, expected {shape_name!r}")
            chart_elem = gf.find(f".//{{{C_NS}}}chart")
            if chart_elem is None or not chart_elem.get(f"{{{R_NS}}}id"):
                raise DeckError(f"Shape {shape_id} on slide {slide_num} is not a chart or missing chart relationship")
            t_key = (slide_num, "chart_data", shape_id)
        elif target_type == "suppress_shape":
            found = False
            for tag in ("sp", "pic", "graphicFrame"):
                for elem in slide_tree.findall(f".//{{{P_NS}}}{tag}"):
                    cnv = elem.find(f".//{{{P_NS}}}cNvPr")
                    if cnv is not None and int(cnv.get("id", "-1")) == shape_id:
                        if shape_name is not None and cnv.get("name") != shape_name:
                            raise DeckError(f"Shape id={shape_id} has name {cnv.get('name')!r}, expected {shape_name!r}")
                        found = True
                        break
                if found:
                    break
            if not found:
                raise DeckError(f"Shape to suppress {shape_id} not found on slide {slide_num}")
            t_key = (slide_num, "suppress_shape", shape_id)
        else:
            raise DeckError(f"Unsupported target_type in slot {name!r}: {target_type!r}")

        if t_key in targets:
            raise DeckError(f"Duplicate slot target destination: {t_key}")
        targets.add(t_key)

    return slots


def _resolve_workbook_binding(
    wb_pkg: WorkbookPackage | None,
    wb_sha256: str | None,
    binding: dict,
    slot_name: str,
) -> tuple[Decimal, str]:
    if wb_pkg is None or wb_sha256 is None:
        raise DeckError(f"Slot {slot_name!r} has workbook_binding, so a recalculated --workbook must be provided")
    expected_wb_sha = binding.get("workbook_sha256")
    if expected_wb_sha is not None and expected_wb_sha != wb_sha256:
        raise DeckError(
            f"Slot {slot_name!r} workbook_sha256 mismatch: expected {expected_wb_sha}, got {wb_sha256}"
        )
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

PML_NAMESPACES = {
    "a": A_NS,
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "c": C_NS,
    "c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
    "c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
    "c16": "http://schemas.microsoft.com/office/drawing/2014/chart",
    "c16r2": "http://schemas.microsoft.com/office/drawing/2015/06/chart",
    "c16r3": "http://schemas.microsoft.com/office/drawing/2017/03/chart",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p": P_NS,
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "r": R_NS,
}


def serialize_pml(root: ET.Element) -> bytes:
    for prefix, uri in PML_NAMESPACES.items():
        ET.register_namespace(prefix, uri)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def serialize_rels(root: ET.Element) -> bytes:
    ET.register_namespace("", PKG_REL_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def serialize_content_types(root: ET.Element) -> bytes:
    ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/content-types")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def serialize_xlsx(root: ET.Element) -> bytes:
    ET.register_namespace("", XLSX_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _update_embedded_xlsx(
    data: bytes,
    values: list[Decimal],
    categories: list[str] | None = None,
) -> bytes:
    """Update values (column B) and categories (column A) in embedded worksheet."""
    with zipfile.ZipFile(io.BytesIO(data)) as z_in:
        wb_parts = {name: z_in.read(name) for name in z_in.namelist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    if sheet_name not in wb_parts:
        return data
    sheet_tree = parse_xml(wb_parts[sheet_name])
    sheet_data = sheet_tree.find(f"{{{XLSX_NS}}}sheetData")
    if sheet_data is None:
        return data
    row_map = {int(r.get("r", "0")): r for r in sheet_data.findall(f"{{{XLSX_NS}}}row")}
    for idx, val in enumerate(values):
        row_num = 2 + idx
        if row_num not in row_map:
            row_elem = ET.SubElement(sheet_data, f"{{{XLSX_NS}}}row", r=str(row_num))
            row_map[row_num] = row_elem
        row_elem = row_map[row_num]
        cell_ref = f"B{row_num}"
        c_elem = next((c for c in row_elem.findall(f"{{{XLSX_NS}}}c") if c.get("r") == cell_ref), None)
        if c_elem is None:
            c_elem = ET.SubElement(row_elem, f"{{{XLSX_NS}}}c", r=cell_ref)
        if "t" in c_elem.attrib:
            del c_elem.attrib["t"]
        v_elem = c_elem.find(f"{{{XLSX_NS}}}v")
        if v_elem is None:
            v_elem = ET.SubElement(c_elem, f"{{{XLSX_NS}}}v")
        v_elem.text = str(val)
        if categories and idx < len(categories):
            cat_ref = f"A{row_num}"
            cat_elem = next((c for c in row_elem.findall(f"{{{XLSX_NS}}}c") if c.get("r") == cat_ref), None)
            if cat_elem is None:
                cat_elem = ET.SubElement(row_elem, f"{{{XLSX_NS}}}c", r=cat_ref)
            cat_elem.attrib["t"] = "inlineStr"
            for child in list(cat_elem):
                cat_elem.remove(child)
            is_elem = ET.SubElement(cat_elem, f"{{{XLSX_NS}}}is")
            t_elem = ET.SubElement(is_elem, f"{{{XLSX_NS}}}t")
            t_elem.text = str(categories[idx])
    wb_parts[sheet_name] = serialize_xlsx(sheet_tree)
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
        for name, b_data in wb_parts.items():
            z_out.writestr(name, b_data)
    return out_buf.getvalue()


def verify_slide_changes(
    original: bytes,
    candidate: bytes,
    approved_shape_ids: set[int],
    suppressed_shape_ids: set[int] | None = None,
) -> None:
    """Ensure no unapproved shapes, styles, or run formatting properties (<a:rPr>) were altered."""
    trees = [deepcopy(parse_xml(original)), deepcopy(parse_xml(candidate))]
    suppressed = suppressed_shape_ids or set()
    for tree in trees:
        # Strip suppressed shapes completely
        for sp_tree in tree.findall(f".//{{{P_NS}}}spTree"):
            for child in list(sp_tree):
                cnv = child.find(f".//{{{P_NS}}}cNvPr")
                if cnv is not None and int(cnv.get("id", "-1")) in suppressed:
                    sp_tree.remove(child)
        for sp in tree.findall(f".//{{{P_NS}}}sp"):
            c_nv_pr = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if c_nv_pr is not None and int(c_nv_pr.get("id", "-1")) in approved_shape_ids:
                tx_body = sp.find(f"{{{P_NS}}}txBody")
                if tx_body is not None:
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
        for pic in tree.findall(f".//{{{P_NS}}}pic"):
            c_nv_pr = pic.find(f"{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr")
            if c_nv_pr is not None and int(c_nv_pr.get("id", "-1")) in approved_shape_ids:
                blip = pic.find(f".//{{{A_NS}}}blip")
                if blip is not None:
                    blip.attrib.clear()
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
    *,
    allow_proposal: bool = False,
) -> tuple[bytes, dict[str, Any]]:
    package = PptxPackage(template)
    blockers = package.blockers()
    if blockers:
        raise DeckError("; ".join(blockers))
    slots = validate_deck_pack(pack, package, allow_proposal=allow_proposal)

    if not isinstance(supplied, dict) or set(supplied) != {"values"} or not isinstance(supplied["values"], dict):
        raise DeckError("Deck inputs must contain exactly a values object")
    values = supplied["values"]
    if set(values) != {s["name"] for s in slots}:
        raise DeckError("Deck input names must exactly match the approved deck-pack slots")

    wb_pkg: WorkbookPackage | None = None
    wb_sha256: str | None = None
    if workbook_bytes is not None:
        wb_pkg = WorkbookPackage(workbook_bytes)
        wb_sha256 = sha(workbook_bytes)
        if wb_pkg.blockers():
            raise DeckError(f"Bound workbook has blockers: {'; '.join(wb_pkg.blockers())}")

    slide_part_map = {s_num: (s_part, deepcopy(tree)) for s_num, s_part, tree in package.slides}
    approved_shapes_per_slide: dict[int, set[int]] = {}
    suppressed_shapes_per_slide: dict[int, set[int]] = {}
    audit: list[dict[str, Any]] = []
    changed = dict(package.parts)

    for slot in slots:
        name = slot["name"]
        item = values[name]
        if (
            not isinstance(item, dict)
            or not {"value", "source"}.issubset(set(item))
            or set(item) - {"value", "source", "unit", "period", "categories", "series_name", "chart_title"}
            or not isinstance(item["source"], str)
            or not item["source"].strip()
        ):
            raise DeckError(f"Slot {name!r} requires a value and a non-empty source reference")

        raw_val = item["value"]
        slide_num = slot["slide"]
        shape_id = slot["shape_id"]
        shape_name = slot.get("shape_name")
        target_type = slot["target_type"]
        s_part, slide_tree = slide_part_map[slide_num]

        if target_type == "suppress_shape":
            removed = False
            for sp_tree in slide_tree.findall(f".//{{{P_NS}}}spTree"):
                for child in list(sp_tree):
                    cnv = child.find(f".//{{{P_NS}}}cNvPr")
                    if cnv is not None and int(cnv.get("id", "-1")) == shape_id:
                        sp_tree.remove(child)
                        removed = True
                        break
                if removed:
                    break
            if not removed:
                raise DeckError(f"Shape {shape_id} on slide {slide_num} could not be suppressed")
            suppressed_shapes_per_slide.setdefault(slide_num, set()).add(shape_id)
            audit.append({
                "name": name,
                "slide": slide_num,
                "target_type": "suppress_shape",
                "shape_id": shape_id,
                "action": "suppressed",
                "source": item["source"],
            })
            continue

        if target_type == "picture_replacement":
            if isinstance(raw_val, str):
                p_asset = Path(raw_val)
                if not p_asset.exists():
                    raise DeckError(f"Image asset file not found: {raw_val}")
                new_img_bytes = p_asset.read_bytes()
            elif isinstance(raw_val, bytes):
                new_img_bytes = raw_val
            else:
                raise DeckError(f"Picture replacement value must be a path string or bytes: {raw_val!r}")
            new_img_sha = sha(new_img_bytes)

            pic = next((p for p in slide_tree.findall(f".//{{{P_NS}}}pic") if int(p.find(f"{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr").get("id", "-1")) == shape_id), None)
            if pic is None:
                raise DeckError(f"Picture shape {shape_id} not found on slide {slide_num}")
            blip = pic.find(f".//{{{A_NS}}}blip")
            r_embed = blip.get(f"{{{R_NS}}}embed")

            s_rels_part = posixpath.join(posixpath.dirname(s_part), "_rels", posixpath.basename(s_part) + ".rels")
            if s_rels_part not in changed:
                raise DeckError(f"Missing slide rels part: {s_rels_part}")
            s_rels = parse_xml(changed[s_rels_part])
            rel = next((r for r in s_rels if r.get("Id") == r_embed), None)
            if rel is None:
                raise DeckError(f"Relationship {r_embed} not found in {s_rels_part}")
            orig_media_part = resolve_rel_target(s_part, rel.get("Target"))
            orig_bytes = package.parts[orig_media_part]
            orig_sha = sha(orig_bytes)

            ref_count = 0
            for part_name, part_content in package.parts.items():
                if part_name.endswith(".rels"):
                    r_root = parse_xml(part_content)
                    for r in r_root:
                        if r.get("TargetMode") == "External":
                            continue
                        tgt = r.get("Target")
                        if not tgt or "\\" in tgt or "://" in tgt:
                            continue
                        try:
                            if resolve_rel_target(posixpath.dirname(part_name), tgt) == orig_media_part:
                                ref_count += 1
                        except DeckError:
                            continue

            if ref_count > 1:
                ext = posixpath.splitext(orig_media_part)[1] or ".png"
                new_media_part = f"ppt/media/img_s{slide_num}_sh{shape_id}_{new_img_sha[:8]}{ext}"
                changed[new_media_part] = new_img_bytes
                rel_target = posixpath.relpath(new_media_part, posixpath.dirname(s_part)).replace("\\", "/")
                rel.set("Target", rel_target)
                changed[s_rels_part] = serialize_rels(s_rels)
                final_media = new_media_part
            else:
                changed[orig_media_part] = new_img_bytes
                final_media = orig_media_part

            approved_shapes_per_slide.setdefault(slide_num, set()).add(shape_id)
            audit.append({
                "name": name,
                "slide": slide_num,
                "target_type": "picture_replacement",
                "shape_id": shape_id,
                "original_media": orig_media_part,
                "original_sha256": orig_sha,
                "new_media": final_media,
                "new_sha256": new_img_sha,
                "source": item["source"],
                "workbook_binding": slot.get("workbook_binding"),
            })
            continue

        if target_type in ("chart_series", "chart_data"):
            if isinstance(raw_val, list):
                vals = [Decimal(str(v)) for v in raw_val]
            elif isinstance(raw_val, (int, float, Decimal, str)):
                vals = [Decimal(str(raw_val))]
            else:
                raise DeckError(f"Chart slot {name!r} value must be a number or list of numbers")

            gf = next((g for g in slide_tree.findall(f".//{{{P_NS}}}graphicFrame") if int(g.find(f"{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr").get("id", "-1")) == shape_id), None)
            if gf is None:
                raise DeckError(f"GraphicFrame shape {shape_id} not found on slide {slide_num}")
            chart_elem = gf.find(f".//{{{C_NS}}}chart")
            chart_rid = chart_elem.get(f"{{{R_NS}}}id")

            s_rels_part = posixpath.join(posixpath.dirname(s_part), "_rels", posixpath.basename(s_part) + ".rels")
            s_rels = parse_xml(changed[s_rels_part])
            rel = next((r for r in s_rels if r.get("Id") == chart_rid), None)
            if rel is None:
                raise DeckError(f"Chart relationship {chart_rid} not found in {s_rels_part}")
            chart_part = resolve_rel_target(s_part, rel.get("Target"))

            chart_root = parse_xml(changed[chart_part])
            series_list = chart_root.findall(f".//{{{C_NS}}}ser")
            ser_idx = slot.get("series_index", 0)
            if ser_idx >= len(series_list):
                raise DeckError(f"Series index {ser_idx} out of range (found {len(series_list)} series)")
            ser = series_list[ser_idx]

            num_cache = ser.find(f".//{{{C_NS}}}val//{{{C_NS}}}numCache")
            if num_cache is None:
                val_node = ser.find(f"{{{C_NS}}}val")
                if val_node is None:
                    val_node = ET.SubElement(ser, f"{{{C_NS}}}val")
                num_ref = val_node.find(f"{{{C_NS}}}numRef")
                if num_ref is None:
                    num_ref = ET.SubElement(val_node, f"{{{C_NS}}}numRef")
                num_cache = ET.SubElement(num_ref, f"{{{C_NS}}}numCache")

            pt_count = num_cache.find(f"{{{C_NS}}}ptCount")
            if pt_count is None:
                pt_count = ET.SubElement(num_cache, f"{{{C_NS}}}ptCount")
            pt_count.set("val", str(len(vals)))
            for pt in list(num_cache.findall(f"{{{C_NS}}}pt")):
                num_cache.remove(pt)
            for idx, v in enumerate(vals):
                pt = ET.SubElement(num_cache, f"{{{C_NS}}}pt", idx=str(idx))
                v_node = ET.SubElement(pt, f"{{{C_NS}}}v")
                v_node.text = str(v)

            categories = slot.get("categories") or item.get("categories")
            if categories:
                str_cache = ser.find(f".//{{{C_NS}}}cat//{{{C_NS}}}strCache")
                if str_cache is None:
                    cat_node = ser.find(f"{{{C_NS}}}cat")
                    if cat_node is None:
                        cat_node = ET.SubElement(ser, f"{{{C_NS}}}cat")
                    str_ref = cat_node.find(f"{{{C_NS}}}strRef")
                    if str_ref is None:
                        str_ref = ET.SubElement(cat_node, f"{{{C_NS}}}strRef")
                    str_cache = ET.SubElement(str_ref, f"{{{C_NS}}}strCache")
                c_pt_count = str_cache.find(f"{{{C_NS}}}ptCount")
                if c_pt_count is None:
                    c_pt_count = ET.SubElement(str_cache, f"{{{C_NS}}}ptCount")
                c_pt_count.set("val", str(len(categories)))
                for pt in list(str_cache.findall(f"{{{C_NS}}}pt")):
                    str_cache.remove(pt)
                for idx, c in enumerate(categories):
                    pt = ET.SubElement(str_cache, f"{{{C_NS}}}pt", idx=str(idx))
                    v_node = ET.SubElement(pt, f"{{{C_NS}}}v")
                    v_node.text = str(c)

            ser_name = slot.get("series_name") or item.get("series_name")
            if ser_name:
                tx = ser.find(f"{{{C_NS}}}tx")
                if tx is not None:
                    tx_v = tx.find(f"{{{C_NS}}}v")
                    if tx_v is None:
                        tx_v = tx.find(f".//{{{C_NS}}}pt/{{{C_NS}}}v")
                    if tx_v is not None:
                        tx_v.text = str(ser_name)

            chart_title = slot.get("chart_title") or item.get("chart_title")
            if chart_title:
                title_elem = chart_root.find(f".//{{{C_NS}}}title")
                if title_elem is not None:
                    p_elem = title_elem.find(f".//{{{A_NS}}}p")
                    if p_elem is not None:
                        _set_paragraph_text_preserving_rpr(p_elem, str(chart_title))

            changed[chart_part] = serialize_pml(chart_root)

            c_rels_part = posixpath.join(posixpath.dirname(chart_part), "_rels", posixpath.basename(chart_part) + ".rels")
            emb_part = None
            if c_rels_part in changed:
                c_rels = parse_xml(changed[c_rels_part])
                pkg_rel = next((r for r in c_rels if r.get("Type") == PACKAGE_REL_TYPE or (r.get("Type") or "").endswith("/package")), None)
                if pkg_rel is not None:
                    emb_part = resolve_rel_target(chart_part, pkg_rel.get("Target"))
                    if emb_part in changed:
                        changed[emb_part] = _update_embedded_xlsx(changed[emb_part], vals, categories)

            approved_shapes_per_slide.setdefault(slide_num, set()).add(shape_id)
            audit.append({
                "name": name,
                "slide": slide_num,
                "target_type": target_type,
                "shape_id": shape_id,
                "chart_part": chart_part,
                "embedded_workbook": emb_part,
                "series_index": ser_idx,
                "values_count": len(vals),
                "source": item["source"],
            })
            continue

        if isinstance(raw_val, bool) or raw_val is None:
            raise DeckError(f"Slot {name!r} value must be a string or finite number")

        binding = slot.get("workbook_binding")
        if binding is not None:
            if "unit" in binding and "unit" in item and item["unit"] != binding["unit"]:
                raise DeckError(f"Slot {name!r} unit mismatch: expected {binding['unit']!r}, got {item['unit']!r}")
            if "period" in binding and "period" in item and item["period"] != binding["period"]:
                raise DeckError(f"Slot {name!r} period mismatch: expected {binding['period']!r}, got {item['period']!r}")
            bound_dec, bound_formatted = _resolve_workbook_binding(wb_pkg, wb_sha256, binding, name)
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

    retained_slides = pack.get("retained_slides")
    for slide_num, (s_part, mod_tree) in slide_part_map.items():
        if retained_slides is not None and slide_num not in retained_slides:
            continue
        shape_ids = approved_shapes_per_slide.get(slide_num, set())
        suppressed_ids = suppressed_shapes_per_slide.get(slide_num, set())
        if shape_ids or suppressed_ids:
            mod_bytes = serialize_pml(mod_tree)
            verify_slide_changes(package.parts[s_part], mod_bytes, shape_ids, suppressed_ids)
            changed[s_part] = mod_bytes

    if retained_slides is not None:
        pres_tree = parse_xml(changed["ppt/presentation.xml"])
        sld_id_lst = pres_tree.find(f"{{{P_NS}}}sldIdLst")
        orig_sld_ids = list(sld_id_lst.findall(f"{{{P_NS}}}sldId")) if sld_id_lst is not None else []
        idx_to_sld_elem = {s_num: orig_sld_ids[s_num - 1] for s_num, _, _ in package.slides}
        for elem in list(sld_id_lst):
            sld_id_lst.remove(elem)
        for s_num in retained_slides:
            sld_id_lst.append(idx_to_sld_elem[s_num])
        changed["ppt/presentation.xml"] = serialize_pml(pres_tree)

        pres_rels_tree = parse_xml(changed["ppt/_rels/presentation.xml.rels"])
        retained_rids = {idx_to_sld_elem[s_num].get(f"{{{R_NS}}}id") for s_num in retained_slides}
        for rel in list(pres_rels_tree):
            rtype = rel.get("Type") or ""
            if rtype.endswith("/slide") and rel.get("Id") not in retained_rids:
                pres_rels_tree.remove(rel)
        changed["ppt/_rels/presentation.xml.rels"] = serialize_rels(pres_rels_tree)

        pruned_slide_parts = set()
        for s_num, s_part, _ in package.slides:
            if s_num not in retained_slides:
                pruned_slide_parts.add(s_part)
                if s_part in changed:
                    del changed[s_part]
                s_rels = posixpath.join(posixpath.dirname(s_part), "_rels", posixpath.basename(s_part) + ".rels")
                if s_rels in changed:
                    del changed[s_rels]

        ct_tree = parse_xml(changed["[Content_Types].xml"])
        for override in list(ct_tree.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")):
            part_name = override.get("PartName", "").lstrip("/")
            if part_name in pruned_slide_parts:
                ct_tree.remove(override)
        changed["[Content_Types].xml"] = serialize_content_types(ct_tree)

    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for entry in package.entries:
            if entry.filename in changed:
                archive.writestr(copy(entry), changed[entry.filename])
        for new_file, new_bytes in changed.items():
            if new_file not in package.parts:
                archive.writestr(new_file, new_bytes)
    output = stream.getvalue()

    result_pkg = PptxPackage(output)
    expected_slide_count = len(retained_slides) if retained_slides is not None else len(package.slides)
    if len(result_pkg.slides) != expected_slide_count:
        raise DeckError(f"Post-write slide count mismatch: expected {expected_slide_count}, got {len(result_pkg.slides)}")

    is_proposal = pack.get("status") == "proposal" or pack.get("approval") is None
    pending_gates = ["render_review"]
    if is_proposal:
        pending_gates.insert(0, "human_pack_approval")

    manifest = {
        "status": "composed_pending_render_review",
        "ready": False,
        "pack_status": "proposal" if is_proposal else "approved",
        "template_sha256": sha(template),
        "output_sha256": sha(output),
        "pack_sha256": sha(json.dumps(pack, sort_keys=True, default=str).encode()),
        "input_sha256": sha(json.dumps(supplied, sort_keys=True, default=str).encode()),
        "workbook_sha256": wb_sha256,
        "changed_slots": audit,
        "changed_parts": sorted(k for k, v in changed.items() if v != package.parts.get(k)),
        "preserved_part_count": sum(1 for k, v in changed.items() if v == package.parts.get(k)),
        "pending_gates": pending_gates,
    }
    if retained_slides is not None:
        manifest["retained_slides"] = retained_slides
        manifest["source_slide_lineage"] = [
            {"curated_slide": i + 1, "source_slide": s_num}
            for i, s_num in enumerate(retained_slides)
        ]
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
    build_cmd.add_argument("--allow-proposal", action="store_true", help="Allow building from an unapproved proposal deck pack")
    args = parser.parse_args()
    try:
        data = Path(args.template).read_bytes()
        if args.command == "inspect":
            print(json.dumps(PptxPackage(data).summary(), indent=2))
        else:
            if Path(args.template).resolve() == Path(args.output).resolve():
                raise DeckError("Output must be different from the original template")
            if Path(args.output).exists():
                raise DeckError(f"Output already exists and will not be overwritten: {args.output}")
            wb_bytes = Path(args.workbook).read_bytes() if args.workbook else None
            output, manifest = compose_deck(
                data,
                read_json(args.pack),
                read_json(args.inputs),
                workbook_bytes=wb_bytes,
                allow_proposal=args.allow_proposal,
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
