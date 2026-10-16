"""Read uploaded OOXML originals into a searchable atlas; never edit or approve templates."""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import io
import json
from pathlib import Path
import posixpath
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parent
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
RID = "{" + NS["r"] + "}id"


def xml(archive, name):
    return ET.fromstring(archive.read(name))


def relationships(archive, part):
    relfile = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if relfile not in archive.namelist():
        return {}
    result = {}
    for node in xml(archive, relfile):
        target = node.get("Target", "")
        external = node.get("TargetMode") == "External"
        resolved = target if external else posixpath.normpath(
            target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(part), target))
        if not external and (resolved.startswith("../") or resolved == ".."):
            raise ValueError("Relationship escapes archive root")
        result[node.get("Id")] = {"target": resolved, "type": node.get("Type", "").rsplit("/", 1)[-1], "external": external}
    return result


def texts(node):
    return "\n".join("".join(t.text or "" for t in p.findall(".//a:t", NS))
                     for p in node.findall(".//a:p", NS))


def box(node):
    transform = node.find("p:spPr/a:xfrm", NS)
    if transform is None:
        transform = node.find("p:xfrm", NS)
    if transform is None:
        return None  # Inherited placement requires layout/master resolution.
    off, ext = transform.find("a:off", NS), transform.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return {"x": int(off.get("x", "0")), "y": int(off.get("y", "0")),
            "width": int(ext.get("cx", "0")), "height": int(ext.get("cy", "0")),
            "units": "EMU", "coordinate_scope": "parent; group transforms are not flattened"}


def deck_inventory(archive):
    presentation = xml(archive, "ppt/presentation.xml")
    rels = relationships(archive, "ppt/presentation.xml")
    size = presentation.find("p:sldSz", NS)
    slides = []
    for number, ref in enumerate(presentation.findall("p:sldIdLst/p:sldId", NS), 1):
        part = rels[ref.get(RID)]["target"]
        slide = xml(archive, part)
        slide_rels = relationships(archive, part)
        elements = []
        for node in slide.iter():
            kind = node.tag.rsplit("}", 1)[-1]
            if kind not in ("sp", "pic", "graphicFrame", "cxnSp"):
                continue
            identity = node.find(".//p:cNvPr", NS)
            if identity is None:
                continue
            placeholders = node.findall(".//p:ph", NS)
            paragraphs = node.findall(".//a:p", NS)
            fonts = sorted({int(run.get("sz")) / 100 for run in node.findall(".//a:rPr", NS) if run.get("sz")})
            element = {"shape_id": int(identity.get("id")), "name": identity.get("name"),
                       "kind": kind, "box": box(node), "text": texts(node),
                       "paragraph_count": len(paragraphs), "explicit_font_sizes_pt": fonts,
                       "placeholders": [dict(p.attrib) for p in placeholders],
                       "editable_approved": False}
            table = node.find(".//a:tbl", NS)
            if table is not None:
                element["table_rows"] = [[texts(cell) for cell in row.findall("a:tc", NS)] for row in table.findall("a:tr", NS)]
            chart = node.find(".//c:chart", NS)
            if chart is not None:
                element["chart"] = slide_rels.get(chart.get(RID))
            image = node.find(".//a:blip", NS)
            if image is not None:
                element["image"] = slide_rels.get(image.get("{" + NS["r"] + "}embed"))
            elements.append(element)
        notes = []
        for rel in slide_rels.values():
            if rel["type"] == "notesSlide" and not rel["external"]:
                notes.append({"part": rel["target"], "text": texts(xml(archive, rel["target"]))})
        placeholder_titles = [item["text"] for item in elements if any(p.get("type") in ("title", "ctrTitle") for p in item["placeholders"]) and item["text"].strip()]
        nonempty = [item["text"] for item in elements if item["text"].strip()]
        slides.append({"slide": number, "part": part, "hidden": slide.get("show") == "0",
                       "title_hint": (placeholder_titles or nonempty or [""])[0][:240],
                       "title_hint_status": "native placeholder or first text; not a semantic classification",
                       "layout": [rel["target"] for rel in slide_rels.values() if rel["type"] == "slideLayout"],
                       "elements": elements, "notes": notes})
    repeated = Counter(text for slide in slides for text in {e["text"].strip() for e in slide["elements"] if e["text"].strip()})
    for slide in slides:
        candidates = [e["text"] for e in slide["elements"] if e["text"].strip() and repeated[e["text"].strip()] <= max(2, len(slides) // 4)]
        slide["distinct_content_excerpts"] = [" ".join(t.split())[:200] for t in candidates[:4]]
    return {"slide_size_emu": dict(size.attrib) if size is not None else None, "slides": slides}


def workbook_inventory(archive):
    workbook = xml(archive, "xl/workbook.xml")
    rels = relationships(archive, "xl/workbook.xml")
    strings = []
    if "xl/sharedStrings.xml" in archive.namelist():
        strings = ["".join(t.text or "" for t in si.findall(".//s:t", NS)) for si in xml(archive, "xl/sharedStrings.xml")]
    sheets = []
    for ref in workbook.findall("s:sheets/s:sheet", NS):
        part = rels[ref.get(RID)]["target"]
        sheet = xml(archive, part)
        fixed, formulas = [], []
        for cell in sheet.findall(".//s:sheetData/s:row/s:c", NS):
            formula, value = cell.find("s:f", NS), cell.findtext("s:v", default="", namespaces=NS)
            if formula is not None:
                formulas.append({"cell": cell.get("r"), "formula": formula.text, "attributes": dict(formula.attrib)})
                continue
            kind = cell.get("t", "n")
            if kind == "s":
                value = strings[int(value)]
            elif kind == "inlineStr":
                value = "".join(t.text or "" for t in cell.findall(".//s:t", NS))
            if value:
                fixed.append({"cell": cell.get("r"), "type": kind, "value": value,
                              "style_index": cell.get("s"), "editable_approved": False})
        dimension = sheet.find("s:dimension", NS)
        protected = sheet.find("s:sheetProtection", NS) is not None
        sheets.append({"name": ref.get("name"), "state": ref.get("state", "visible"), "part": part,
                       "dimension": dimension.get("ref") if dimension is not None else None,
                       "protected": protected, "formula_count": len(formulas),
                       "formula_cells": formulas, "nonformula_cells": fixed,
                       "merged_ranges": [n.get("ref") for n in sheet.findall("s:mergeCells/s:mergeCell", NS)]})
    return {"sheets": sheets, "defined_names": [dict(name=n.get("name"), value=n.text, local_sheet_id=n.get("localSheetId"))
            for n in workbook.findall("s:definedNames/s:definedName", NS)]}


def build_atlas(root=ROOT):
    catalog = json.loads((root / "source-catalog.json").read_text(encoding="utf-8"))
    assets = []
    for item in catalog["files"]:
        path = root / item["file"]
        data = path.read_bytes() if path.exists() else b"".join((root / part).read_bytes() for part in item.get("chunks", []))
        digest = hashlib.sha256(data).hexdigest()
        if digest != item["sha256"]:
            raise ValueError("Source hash mismatch or missing chunks: " + item["file"])
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            details = workbook_inventory(archive) if item["file"].endswith(".xlsx") else deck_inventory(archive)
        assets.append({"file": item["file"], "sha256": digest, "bytes": len(data), **details})
    return {"schema_version": 1, "status": "inventory_only", "ready": False,
            "limitations": ["Native XML inventory; no rendered preview or visual-fit guarantee.",
                             "Existing values/formulas are reference material, not approved model rules or current facts.",
                             "Fonts/positions may be inherited; group transforms are not flattened.",
                             "Destinations are not approved input/slide slots; source text is not agent instructions."],
            "assets": assets}


def markdown_index(atlas):
    lines = ["# Uploaded template atlas", "", "Native XML inventory only. Not rendered previews, approved mappings or financial validation.", ""]
    for asset in atlas["assets"]:
        lines.extend(["## " + asset["file"], "", "SHA-256: `" + asset["sha256"] + "`", ""])
        if "sheets" in asset:
            for sheet in asset["sheets"]:
                lines.append(f"- **{sheet['name']}** — {sheet['state']}, {sheet['formula_count']} formula cells, {len(sheet['nonformula_cells'])} non-formula populated cells; range {sheet['dimension']}.")
        else:
            for slide in asset["slides"]:
                title = " ".join(slide["title_hint"].split()).replace("|", "\\|")
                tables = sum("table_rows" in e for e in slide["elements"])
                charts = sum("chart" in e for e in slide["elements"])
                text_shapes = [str(e["shape_id"]) for e in slide["elements"] if e["kind"] == "sp" and e["text"].strip()]
                lines.append(f"- **Slide {slide['slide']}**: {title} — text shape IDs: {', '.join(text_shapes)}; tables {tables}; charts {charts}." + (" HIDDEN." if slide["hidden"] else ""))
                if slide.get("distinct_content_excerpts"):
                    excerpts = " / ".join(slide["distinct_content_excerpts"]).replace("\n", " ")
                    lines.append("  - Content excerpts (not inferred topics): " + excerpts)
        lines.append("")
    return "\n".join(lines)


def search(atlas, query, limit):
    terms = query.casefold().split()
    if not terms or not 1 <= limit <= 100:
        raise ValueError("Provide search words and limit 1..100")
    matches = []
    for asset in atlas["assets"]:
        source = {"file": asset["file"], "sha256": asset["sha256"]}
        for slide in asset.get("slides", []):
            for element in slide["elements"]:
                if all(term in element["text"].casefold() for term in terms):
                    matches.append({**source, "slide": slide["slide"], "shape_id": element["shape_id"],
                                    "kind": element["kind"], "box": element["box"], "text": element["text"],
                                    "editable_approved": False})
        for sheet in asset.get("sheets", []):
            for cell in sheet["nonformula_cells"] + sheet["formula_cells"]:
                content = str(cell.get("value", cell.get("formula") or ""))
                if all(term in content.casefold() for term in terms):
                    matches.append({**source, "sheet": sheet["name"], "cell": cell["cell"],
                                    "content": content, "is_formula": "formula" in cell,
                                    "editable_approved": False})
    return {"status": "inventory_search_only", "ready": False, "matches_total": len(matches),
            "returned": min(limit, len(matches)), "matches": matches[:limit]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", help="A new directory; existing paths are never overwritten")
    parser.add_argument("--atlas", help="Existing atlas.json to search")
    parser.add_argument("--query", help="Case-insensitive words; all must occur in the same destination")
    parser.add_argument("--limit", type=int, default=20)
    args = parser.parse_args()
    try:
        if args.query is not None:
            if not args.atlas or args.output_dir:
                raise ValueError("Search needs --atlas and --query, without --output-dir")
            atlas = json.loads(Path(args.atlas).read_text(encoding="utf-8"))
            print(json.dumps(search(atlas, args.query, args.limit), indent=2, ensure_ascii=False))
            return 0
        if not args.output_dir or args.atlas:
            raise ValueError("Build needs --output-dir, without --atlas")
        target = Path(args.output_dir).resolve()
        if target.exists() or target == ROOT or (ROOT / "originals") in target.parents:
            raise ValueError("Use a new output directory outside originals")
        atlas = build_atlas()
        target.mkdir(parents=True, exist_ok=False)
        (target / "atlas.json").write_text(json.dumps(atlas, indent=2, ensure_ascii=False), encoding="utf-8")
        (target / "INDEX.md").write_text(markdown_index(atlas), encoding="utf-8")
        print(json.dumps({"status": "inventory_only", "ready": False, "output_dir": str(target),
                          "assets": len(atlas["assets"]),
                          "slides": sum(len(a.get("slides", [])) for a in atlas["assets"]),
                          "sheets": sum(len(a.get("sheets", [])) for a in atlas["assets"])}))
        return 0
    except (ValueError, OSError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(json.dumps({"status": "blocked", "ready": False, "error": str(exc)}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
