"""Plan content against a hash-pinned uploaded template; never generate or certify files."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parent
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class PlanError(ValueError):
    pass


def require(condition, message):
    if not condition:
        raise PlanError(message)


def text(value, location):
    require(isinstance(value, str) and value.strip(), f"{location}: non-empty text required")
    return value


def unique_object(pairs):
    result = {}
    for key, value in pairs:
        require(key not in result, f"Duplicate JSON key: {key}")
        result[key] = value
    return result


def read_json(path):
    def bad_number(value):
        raise PlanError(f"Non-finite JSON number: {value}")
    return json.loads(Path(path).read_text(encoding="utf-8-sig"),
                      object_pairs_hook=unique_object, parse_constant=bad_number)


def keys(obj, expected, location):
    require(isinstance(obj, dict) and set(obj) == set(expected),
            f"{location}: expected exactly {', '.join(sorted(expected))}")


def make_plan(request, root=ROOT):
    keys(request, {"kind", "mode", "template", "template_sha256", "items"}, "request")
    kind = request["kind"]
    require(kind in ("excel", "ppt"), "kind must be excel or ppt")
    require(request["mode"] == "exact_template",
            "Only exact_template is implemented. Universal Excel derivatives and cross-deck assembly require separate reviewed adapters; no generic fallback.")
    catalog = read_json(root / "source-catalog.json")
    approved = {item["file"]: item for item in catalog["files"]}
    name = text(request["template"], "template")
    require(name in approved, "Template must be an uploaded original in source-catalog.json")
    entry = approved[name]
    require(name.endswith(".xlsx" if kind == "excel" else ".pptx"), "Template type mismatch")
    require(request["template_sha256"] == entry["sha256"], "Requested template hash differs from catalog")
    path = root / name
    require(path.is_file(), "Original missing; run python verify_sources.py to reconstruct source chunks")
    actual = hashlib.sha256(path.read_bytes()).hexdigest()
    require(actual == entry["sha256"], "Original bytes differ from catalog hash")
    items = request["items"]
    require(isinstance(items, list) and 0 < len(items) <= 1000, "items must contain 1..1000 mappings")
    destinations, identifiers = set(), set()
    mapped = []
    with zipfile.ZipFile(path) as archive:
        if kind == "excel":
            # Use the existing engine's relationship resolution and protection checks.
            from modelmaker import Package
            workbook = Package(path.read_bytes())
        else:
            presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
            rels = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
            relmap = {rel.get("Id"): rel.get("Target") for rel in rels}
            slides = presentation.findall(P + "sldIdLst/" + P + "sldId")
        for index, item in enumerate(items):
            loc = f"items[{index}]"
            fields = {"id", "value", "source", "units", "period", "destination"}
            keys(item, fields, loc)
            identifier = text(item["id"], loc + ".id")
            require(identifier not in identifiers, f"{loc}: duplicate id")
            identifiers.add(identifier)
            for field in ("source", "units", "period"):
                text(item[field], loc + "." + field)
            value = item["value"]
            require(type(value) in (str, int, float), f"{loc}: value must be text or finite number")
            if isinstance(value, str):
                text(value, loc + ".value")
            if isinstance(value, float):
                import math
                require(math.isfinite(value), f"{loc}: non-finite value")
            dest = item["destination"]
            if kind == "excel":
                keys(dest, {"sheet", "cell"}, loc + ".destination")
                sheet, cell = dest["sheet"], dest["cell"]
                text(sheet, loc + ".sheet")
                require(isinstance(cell, str) and re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", cell), f"{loc}: invalid cell")
                require(sheet in workbook.sheets, f"{loc}: unknown sheet")
                node = next((c for c in workbook.sheets[sheet][1].iter(S + "c") if c.get("r") == cell), None)
                require(node is not None and node.find(S + "f") is None, f"{loc}: destination must be an existing non-formula cell")
                require(node.get("t") in (None, "n") and node.find(S + "v") is not None,
                        f"{loc}: existing numeric input required by current composer")
                require(type(value) in (int, float), f"{loc}: numeric value required")
                target = (sheet, cell)
            else:
                keys(dest, {"slide", "shape_id", "max_characters"}, loc + ".destination")
                slide, shape, budget = dest["slide"], dest["shape_id"], dest["max_characters"]
                require(type(slide) is int and 1 <= slide <= len(slides), f"{loc}: slide outside uploaded deck")
                require(type(shape) is int and shape > 0, f"{loc}: invalid shape id")
                require(type(budget) is int and 1 <= budget <= 10000, f"{loc}: set explicit text budget")
                import posixpath
                target_name = relmap[slides[slide - 1].get(R + "id")]
                part = posixpath.normpath("ppt/" + target_name) if not target_name.startswith("/") else target_name.lstrip("/")
                require(part.startswith("ppt/slides/"), f"{loc}: unsupported slide relationship")
                tree = ET.fromstring(archive.read(part))
                shapes = [sp for sp in tree.iter(P + "sp")
                          if (n := sp.find(P + "nvSpPr/" + P + "cNvPr")) is not None and n.get("id") == str(shape)]
                require(len(shapes) == 1 and shapes[0].find(P + "txBody") is not None,
                        f"{loc}: select an existing text shape; chart/table adapters need separate mappings")
                require(len(str(value)) <= budget, f"{loc}: text exceeds slot budget; edit content or select another uploaded slide")
                target = (slide, shape)
            require(target not in destinations, f"{loc}: duplicate destination")
            destinations.add(target)
            mapped.append(item)
    return {"status": "planned_needs_review", "ready": False, "kind": kind,
            "mode": "exact_template", "template": name, "template_sha256": actual,
            "catalog_approval_status": entry.get("approval_status", "unreviewed"),
            "items": mapped, "required_next_steps": [
                "Review provenance, units, periods and editable input/slot map; planning does not approve a model pack.",
                "Compose a separate candidate using the corresponding reviewed model/deck pack.",
                "Recalculate Excel and check independent financial identities; bind deck figures to that exact workbook hash.",
                "Render and inspect the actual output against its uploaded template, including stale company data and overflow."
            ]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--request", required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(make_plan(read_json(args.request)), indent=2, allow_nan=False))
        return 0
    except (PlanError, ValueError, OSError, KeyError, TypeError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(json.dumps({"status": "blocked", "ready": False, "error": str(exc)}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
