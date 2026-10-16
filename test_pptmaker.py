"""Unit and regression tests for pptmaker.py."""
from __future__ import annotations

from copy import deepcopy
from decimal import Decimal
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from modelmaker import NS as XLSX_NS, sha
from pptmaker import (
    A_NS,
    DeckError,
    P_NS,
    PptxPackage,
    R_NS,
    compose_deck,
    verify_slide_changes,
)
from test_modelmaker import fixture as xlsx_fixture


def pptx_fixture(
    slide1_xml: bytes | None = None,
    slide2_xml: bytes | None = None,
    extra: dict[str, bytes] | None = None,
) -> bytes:
    default_slide1 = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">'
        f'<p:cSld><p:spTree>'
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="10" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="100" y="200"/></a:xfrm></p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>ACME CORP</a:t></a:r></a:p>'
        f'<a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>Ticker: ACME | FY25 Revenue: $100M</a:t></a:r></a:p>'
        f'</p:txBody></p:sp>'
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="11" name="Untouched Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr/>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Static Disclaimer</a:t></a:r></a:p></p:txBody>'
        f'</p:sp>'
        f'<p:graphicFrame>'
        f'<p:nvGraphicFramePr><p:cNvPr id="20" name="Financial Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
        f'<p:xfrm/>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000"/><a:gridCol w="1000"/></a:tblGrid>'
        f'<a:tr h="300">'
        f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Metric</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>'
        f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>FY25</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>'
        f'</a:tr>'
        f'<a:tr h="300">'
        f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>EBITDA</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>'
        f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200" i="1"/><a:t>0.00</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>'
        f'</a:tr>'
        f'</a:tbl></a:graphicData></a:graphic>'
        f'</p:graphicFrame>'
        f'</p:spTree></p:cSld></p:sld>'
    ).encode("utf-8")

    default_slide2 = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">'
        f'<p:cSld><p:spTree>'
        f'<p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide 2 Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Unedited Slide 2</a:t></a:r></a:p></p:txBody></p:sp>'
        f'</p:spTree></p:cSld></p:sld>'
    ).encode("utf-8")

    parts = {
        "[Content_Types].xml": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
            b'<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            b'<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            b'</Types>'
        ),
        "ppt/presentation.xml": (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:presentation xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">'
            f'<p:sldIdLst>'
            f'<p:sldId id="256" r:id="rId1"/>'
            f'<p:sldId id="257" r:id="rId2"/>'
            f'</p:sldIdLst>'
            f'</p:presentation>'
        ).encode("utf-8"),
        "ppt/_rels/presentation.xml.rels": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
            b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>'
            b'</Relationships>'
        ),
        "ppt/slides/slide1.xml": slide1_xml or default_slide1,
        "ppt/slides/slide2.xml": slide2_xml or default_slide2,
        "ppt/slides/_rels/slide1.xml.rels": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>'
            b'</Relationships>'
        ),
        "ppt/slideLayouts/slideLayout1.xml": b'<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        "ppt/slideMasters/slideMaster1.xml": b'<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        "ppt/charts/chart1.xml": b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
        "ppt/media/image1.png": b"\x89PNG\r\n\x1a\nfakeimagebytes",
    }
    if extra:
        parts.update(extra)

    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as zf:
        for name, content in parts.items():
            zf.writestr(name, content)
    return stream.getvalue()


def make_deck_pack(pptx_bytes: bytes) -> dict:
    digest = sha(pptx_bytes)
    return {
        "schema_version": 1,
        "template_sha256": digest,
        "approval": {
            "reviewer": "Unit Test Reviewer",
            "approved_at": "2026-10-16T00:00:00Z",
            "template_sha256": digest,
        },
        "slots": [
            {
                "name": "company_title",
                "slide": 1,
                "target_type": "shape_text",
                "shape_id": 10,
                "shape_name": "Title 1",
                "paragraph_index": 0,
            },
            {
                "name": "ticker_token",
                "slide": 1,
                "target_type": "shape_token",
                "shape_id": 10,
                "shape_name": "Title 1",
                "token": "ACME",
            },
            {
                "name": "ebitda_cell",
                "slide": 1,
                "target_type": "table_cell",
                "shape_id": 20,
                "shape_name": "Financial Table",
                "row": 1,
                "col": 1,
                "workbook_binding": {
                    "sheet": "Forecast",
                    "cell": "C2",
                    "scale": 1.0,
                    "decimals": 2,
                    "abs_tolerance": 0.01,
                },
            },
        ],
    }


def make_deck_inputs(ebitda_val: float | str = 10.0) -> dict:
    return {
        "values": {
            "company_title": {
                "value": "MARKSANS PHARMA LTD.",
                "source": "Annual Report FY25",
            },
            "ticker_token": {
                "value": "MARKSANS",
                "source": "NSE Symbol Directory",
            },
            "ebitda_cell": {
                "value": ebitda_val,
                "source": "Recalculated Workbook Forecast!C2",
            },
        }
    }


class PptMakerTests(unittest.TestCase):
    def setUp(self):
        self.pptx = pptx_fixture()
        self.pack = make_deck_pack(self.pptx)
        self.wb = xlsx_fixture()

    def test_inspect_inventory_and_read_only(self):
        pkg = PptxPackage(self.pptx)
        summary = pkg.summary()
        self.assertEqual(summary["slide_count"], 2)
        self.assertEqual(summary["layout_count"], 1)
        self.assertEqual(summary["master_count"], 1)
        self.assertEqual(summary["chart_count"], 1)
        self.assertEqual(summary["media_count"], 1)
        self.assertEqual(summary["unsupported_features"], [])
        slide1 = summary["slides"][0]
        self.assertEqual(slide1["layout"], "ppt/slideLayouts/slideLayout1.xml")
        self.assertEqual(slide1["charts"], ["ppt/charts/chart1.xml"])
        self.assertEqual(slide1["table_count"], 1)
        self.assertEqual(slide1["tables"][0]["cells"][1][1], "0.00")

    def test_build_preserves_unrelated_parts_and_run_formatting(self):
        output, manifest = compose_deck(
            self.pptx,
            self.pack,
            make_deck_inputs(10.0),
            workbook_bytes=self.wb,
        )
        self.assertEqual(manifest["status"], "composed_pending_render_review")
        self.assertFalse(manifest["ready"])
        self.assertIn("render_review", manifest["pending_gates"])
        self.assertEqual(manifest["changed_parts"], ["ppt/slides/slide1.xml"])

        orig_pkg = PptxPackage(self.pptx)
        out_pkg = PptxPackage(output)
        # All non-edited parts (including slide2.xml, layouts, masters, charts, media) are byte-for-byte identical
        for part_name, orig_bytes in orig_pkg.parts.items():
            if part_name != "ppt/slides/slide1.xml":
                self.assertEqual(out_pkg.parts[part_name], orig_bytes, part_name)

        # Verify <a:rPr> formatting attributes (sz="2400", b="1", solidFill) survived on Slide 1 Shape 10
        s1_tree = out_pkg.slides[0][2]
        sp10 = next(
            sp for sp in s1_tree.findall(f".//{{{P_NS}}}sp")
            if sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr").get("id") == "10"
        )
        paras = sp10.findall(f".//{{{A_NS}}}p")
        self.assertEqual(paras[0].findtext(f".//{{{A_NS}}}t"), "MARKSANS PHARMA LTD.")
        rpr0 = paras[0].find(f".//{{{A_NS}}}r/{{{A_NS}}}rPr")
        self.assertIsNotNone(rpr0)
        self.assertEqual(rpr0.get("sz"), "2400")
        self.assertEqual(rpr0.get("b"), "1")
        self.assertIsNotNone(rpr0.find(f"{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr"))

        # Verify token replacement in paragraph 1 preserved sz="1400"
        self.assertEqual(paras[1].findtext(f".//{{{A_NS}}}t"), "Ticker: MARKSANS | FY25 Revenue: $100M")
        rpr1 = paras[1].find(f".//{{{A_NS}}}r/{{{A_NS}}}rPr")
        self.assertEqual(rpr1.get("sz"), "1400")

        # Verify table cell (1, 1) updated to "10.00" and preserved sz="1200" i="1"
        tbl = s1_tree.find(f".//{{{A_NS}}}tbl")
        tc_1_1 = tbl.findall(f"{{{A_NS}}}tr")[1].findall(f"{{{A_NS}}}tc")[1]
        self.assertEqual(tc_1_1.findtext(f".//{{{A_NS}}}t"), "10.00")
        tc_rpr = tc_1_1.find(f".//{{{A_NS}}}r/{{{A_NS}}}rPr")
        self.assertEqual(tc_rpr.get("sz"), "1200")
        self.assertEqual(tc_rpr.get("i"), "1")

    def test_workbook_binding_rejects_mismatch_or_uncalculated_cell(self):
        # Value mismatch (15.0 vs 10.0 in Forecast!C2)
        with self.assertRaises(DeckError):
            compose_deck(self.pptx, self.pack, make_deck_inputs(15.0), workbook_bytes=self.wb)

        # Missing workbook when workbook_binding is present
        with self.assertRaises(DeckError):
            compose_deck(self.pptx, self.pack, make_deck_inputs(10.0), workbook_bytes=None)

        # Workbook with uncalculated formula cell (missing <v>)
        uncalc_sheet = (
            f'<worksheet xmlns="{XLSX_NS}"><sheetData>'
            f'<row r="2"><c r="B2" t="n"><v>0.10</v></c><c r="C2"><f>B2*100</f></c></row>'
            f'</sheetData></worksheet>'
        ).encode()
        uncalc_wb = xlsx_fixture(sheet=uncalc_sheet)
        with self.assertRaises(DeckError):
            compose_deck(self.pptx, self.pack, make_deck_inputs(10.0), workbook_bytes=uncalc_wb)

    def test_provenance_and_pack_validation_guards(self):
        # Empty source rejected
        bad_inputs = make_deck_inputs(10.0)
        bad_inputs["values"]["company_title"]["source"] = "   "
        with self.assertRaises(DeckError):
            compose_deck(self.pptx, self.pack, bad_inputs, workbook_bytes=self.wb)

        # Wrong template hash rejected
        bad_pack = deepcopy(self.pack)
        bad_pack["template_sha256"] = "0" * 64
        with self.assertRaises(DeckError):
            compose_deck(self.pptx, bad_pack, make_deck_inputs(10.0), workbook_bytes=self.wb)

        # Duplicate slot destination rejected
        dup_pack = deepcopy(self.pack)
        dup_pack["slots"].append(deepcopy(dup_pack["slots"][0]))
        dup_pack["slots"][-1]["name"] = "duplicate_title"
        with self.assertRaises(DeckError):
            compose_deck(self.pptx, dup_pack, make_deck_inputs(10.0), workbook_bytes=self.wb)

    def test_semantic_slide_guard_rejects_unapproved_shape_changes(self):
        orig_slide1 = PptxPackage(self.pptx).parts["ppt/slides/slide1.xml"]
        # Modifying Untouched Box (id="11") when only id=10 is approved must fail
        tampered = orig_slide1.replace(b"Static Disclaimer", b"Tampered Disclaimer")
        with self.assertRaises(DeckError):
            verify_slide_changes(orig_slide1, tampered, {10})
        # Modifying shape coordinates (<a:off x="100"...) even on approved shape 10 must fail
        tampered_geom = orig_slide1.replace(b'x="100"', b'x="999"')
        with self.assertRaises(DeckError):
            verify_slide_changes(orig_slide1, tampered_geom, {10})

    def test_security_and_cli_overwrite_protection(self):
        with self.assertRaises(DeckError):
            PptxPackage(pptx_fixture(extra={"../escape.xml": b"<x/>"}))
        macro_pptx = pptx_fixture(extra={"ppt/vbaProject.bin": b"macro"})
        self.assertTrue(len(PptxPackage(macro_pptx).blockers()) > 0)
        with self.assertRaises(DeckError):
            compose_deck(macro_pptx, make_deck_pack(macro_pptx), make_deck_inputs(10.0), workbook_bytes=self.wb)

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tpl = root / "template.pptx"
            out = root / "candidate.pptx"
            pack_f = root / "pack.json"
            in_f = root / "inputs.json"
            wb_f = root / "wb.xlsx"
            tpl.write_bytes(self.pptx)
            pack_f.write_text(json.dumps(self.pack))
            in_f.write_text(json.dumps(make_deck_inputs(10.0)))
            wb_f.write_bytes(self.wb)

            cmd = [
                sys.executable,
                str(Path(__file__).with_name("pptmaker.py")),
                "build",
                "--template",
                str(tpl),
                "--pack",
                str(pack_f),
                "--inputs",
                str(in_f),
                "--workbook",
                str(wb_f),
                "--output",
                str(tpl),
            ]
            # Refuses to overwrite template
            res_same = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(res_same.returncode, 2)
            self.assertEqual(json.loads(res_same.stdout)["status"], "blocked")

            # Succeeds on fresh output, refuses to overwrite existing output
            cmd[-1] = str(out)
            res_ok = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(res_ok.returncode, 0, res_ok.stdout + res_ok.stderr)
            res_dup = subprocess.run(cmd, capture_output=True, text=True)
            self.assertEqual(res_dup.returncode, 2)

    def test_real_originals_five_decks_and_err4_composition(self):
        root_dir = Path(__file__).resolve().parent
        originals = root_dir / "originals"
        err4_path = originals / "ERR-4.pptx"
        wb_path = originals / "Marksans Keshav.xlsx"
        pack_path = root_dir / "proposals" / "err4.deck-pack.proposal.json"
        inputs_path = root_dir / "proposals" / "err4.inputs.proposal.json"
        recalc_path = root_dir / "candidates" / "Marksans-Recalculated.xlsx"
        if not err4_path.exists() or not wb_path.exists():
            return

        pptx_files = sorted(originals.glob("*.pptx"))
        self.assertEqual(len(pptx_files), 5)
        total_slides = sum(PptxPackage(p.read_bytes()).summary()["slide_count"] for p in pptx_files)
        self.assertEqual(total_slides, 182)

        from modelmaker import read_json
        wb_data = recalc_path.read_bytes() if recalc_path.exists() else wb_path.read_bytes()
        deck_pack = read_json(str(pack_path))
        if not recalc_path.exists():
            wb_digest = sha(wb_data)
            for s in deck_pack.get("slots", []):
                if s.get("workbook_binding"):
                    s["workbook_binding"]["workbook_sha256"] = wb_digest

        out_bytes, manifest = compose_deck(
            err4_path.read_bytes(),
            deck_pack,
            read_json(str(inputs_path)),
            workbook_bytes=wb_data,
            allow_proposal=True,
        )
        self.assertEqual(manifest["status"], "composed_pending_render_review")
        self.assertFalse(manifest["ready"])
        self.assertIn("render_review", manifest["pending_gates"])
        self.assertIn("human_pack_approval", manifest["pending_gates"])
        self.assertEqual(len(manifest["changed_slots"]), len(deck_pack["slots"]))
        self.assertEqual(len(manifest["changed_slots"]), 59)
        self.assertEqual(
            manifest["changed_parts"],
            [
                "ppt/slides/slide1.xml",
                "ppt/slides/slide14.xml",
                "ppt/slides/slide16.xml",
                "ppt/slides/slide2.xml",
                "ppt/slides/slide23.xml",
            ],
        )
        out_summary = PptxPackage(out_bytes).summary()
        self.assertEqual(out_summary["slide_count"], 25)
        self.assertEqual(out_summary["chart_count"], 25)


if __name__ == "__main__":
    unittest.main()
