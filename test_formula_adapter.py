"""Adapter regression tests: well-formed special formulas compose with byte
preservation; malformed structures and protected-zone violations stay blocked.
Synthetic fixtures exercise safeguards; the Marksans integration test checks
the real workbook. These are not Excel certification."""
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
import unittest

from modelmaker import ModelError, Package, compose, sha, NS
from test_modelmaker import fixture, pack, inputs
import formula_adapter


class AdapterTests(unittest.TestCase):
    def build(self, data, pack_data=None, supplied=None):
        return compose(data, pack_data or self.pack_for(data), supplied or inputs())

    def pack_for(self, data, check_cell="B2"):
        """Fixture pack whose check targets a cell that exists in the custom sheet."""
        p = pack(data)
        p["checks"][0]["cell"] = check_cell
        p["checks"][0]["expected"] = 1
        return p

    def sheet(self, body, root_attrs="", tail=""):
        return f'<worksheet xmlns="{NS}"{root_attrs}>{body}{tail}</worksheet>'.encode()

    def fragments_equal(self, source: Package, candidate: Package):
        for name, (part, _) in source.sheets.items():
            self.assertEqual(
                formula_adapter.formula_fragments(source.parts[part]),
                formula_adapter.formula_fragments(candidate.parts[part]),
                f"formula fragments changed on {name}")

    def test_wellformed_shared_group_builds_and_preserves_formula_bytes(self):
        body = ('<sheetData><row r="2"><c r="B2"><v>1</v></c>'
                '<c r="C2"><f t="shared" si="0" ref="C2:C5">B2*100</f><v>100</v></c>'
                '<c r="C3"><f t="shared" si="0"/><v>100</v></c>'
                '<c r="C4"><f t="shared" si="0"/><v>100</v></c>'
                '<c r="C5"><f t="shared" si="0"/><v>100</v></c>'
                '</row></sheetData>')
        data = fixture(sheet=self.sheet(body))
        output, manifest = self.build(data)
        self.assertEqual(manifest["status"], "needs_recalculation")
        self.assertFalse(manifest["ready"])
        source, candidate = Package(data), Package(output)
        self.assertEqual(source.blockers(), [])
        self.assertEqual(candidate.blockers(), [])
        self.assertEqual(candidate.summary()["special_formulas"]["Forecast"], {"shared": 4})
        self.fragments_equal(source, candidate)
        sheet_bytes = candidate.parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b'<f t="shared" si="0" ref="C2:C5">B2*100</f>', sheet_bytes)
        self.assertIn(b'<c r="C3"><f t="shared" si="0"/></c>', sheet_bytes)
        self.assertNotIn(b"<v>100</v>", sheet_bytes)
        self.assertIn(b'<c r="B2"><v>0.2</v></c>', sheet_bytes)

    def test_malformed_shared_variants_are_blocked(self):
        cases = {
            "follower with text": '<c r="C2"><f t="shared" si="0" ref="C2:C5">B2*100</f><v>1</v></c><c r="C3"><f t="shared" si="0">B2*100</f><v>1</v></c>',
            "follower with ref": '<c r="C2"><f t="shared" si="0" ref="C2:C5">B2*100</f><v>1</v></c><c r="C3"><f t="shared" si="0" ref="C3"/><v>1</v></c>',
            "unresolved si": '<c r="C2"><f t="shared" si="0" ref="C2:C5">B2*100</f><v>1</v></c><c r="C3"><f t="shared" si="9"/><v>1</v></c>',
            "master without si": '<c r="C2"><f t="shared" ref="C2:C5">B2*100</f><v>1</v></c>',
            "master without ref": '<c r="C2"><f t="shared" si="0">B2*100</f><v>1</v></c>',
            "follower outside master range": '<c r="C2"><f t="shared" si="0" ref="C2:C5">B2*100</f><v>1</v></c><c r="D9"><f t="shared" si="0"/><v>1</v></c>',
        }
        for label, cells in cases.items():
            data = fixture(sheet=self.sheet(f'<sheetData><row r="2"><c r="B2"><v>1</v></c>{cells}</row></sheetData>'))
            with self.subTest(label=label), self.assertRaises(ModelError) as caught:
                self.build(data)
            self.assertNotIn("Unsupported native/external feature", str(caught.exception))

    def test_wellformed_array_zone_is_preserved_and_protected(self):
        body = ('<sheetData><row r="2"><c r="B2"><v>1</v></c>'
                '<c r="D2"><f t="array" ref="D2:E2">B2*2</f><v>2</v></c>'
                '<c r="E2"><v>2</v></c>'
                '</row></sheetData>')
        data = fixture(sheet=self.sheet(body))
        output, manifest = self.build(data)
        source, candidate = Package(data), Package(output)
        self.assertEqual(candidate.summary()["special_formulas"]["Forecast"], {"array": 1})
        self.assertEqual(candidate.protected_zones(), {"Forecast": ["D2:E2"]})
        self.fragments_equal(source, candidate)
        sheet_bytes = candidate.parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b'<f t="array" ref="D2:E2">B2*2</f>', sheet_bytes)
        self.assertNotIn(b'<v>2</v>', sheet_bytes)
        self.assertIn("E2", manifest["stale_cache_cells_stripped"])
        # mapping an input into the array zone is refused: the master cell is
        # already refused as a formula cell, the cached result cell as a zone member
        p = self.pack_for(data)
        p["inputs"][0]["cell"] = "D2"
        with self.assertRaises(ModelError) as caught:
            self.build(data, pack_data=p)
        self.assertIn("non-formula", str(caught.exception))
        p = self.pack_for(data)
        p["inputs"][0]["cell"] = "E2"
        with self.assertRaises(ModelError) as caught:
            self.build(data, pack_data=p)
        self.assertIn("protected", str(caught.exception))

    def test_data_table_zone_is_preserved_cache_stripped_and_protected(self):
        body = ('<sheetData><row r="2"><c r="B2"><v>1</v></c>'
                '<c r="B3"><v>2</v></c>'
                '<c r="C3"><f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="B3" r2="B2" ca="1"/><v>9</v></c>'
                '<c r="D3"><v>9</v></c>'
                '</row><row r="4"><c r="C4"><v>9</v></c><c r="D4"><v>9</v></c></row></sheetData>')
        data = fixture(sheet=self.sheet(body))
        output, manifest = self.build(data)
        source, candidate = Package(data), Package(output)
        report = candidate.formula_reports["Forecast"]
        self.assertEqual(candidate.summary()["special_formulas"]["Forecast"], {"dataTable": 1})
        self.assertEqual(candidate.protected_zones(), {"Forecast": ["C3:D4"]})
        self.assertEqual(report.table_input_cells, ["B3", "B2"])
        self.fragments_equal(source, candidate)
        sheet_bytes = candidate.parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b'<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="B3" r2="B2" ca="1"/>', sheet_bytes)
        for ref in ("D3", "C4", "D4"):
            self.assertIn(ref, manifest["stale_cache_cells_stripped"])
        self.assertNotIn(b"<v>9</v>", sheet_bytes)
        for ref in ("C3", "D3"):
            p = self.pack_for(data)
            p["inputs"][0]["cell"] = ref
            with self.subTest(cell=ref), self.assertRaises(ModelError):
                self.build(data, pack_data=p)
        # malformed dataTable (missing r1/r2) stays blocked
        bad = body.replace('r1="B3" r2="B2" ', "")
        bad_data = fixture(sheet=self.sheet(bad))
        with self.assertRaises(ModelError) as caught:
            self.build(bad_data)
        self.assertIn("dataTable is missing r1", str(caught.exception))

    def test_sparkline_prefixed_tags_outside_sheetdata_are_preserved(self):
        body = ('<sheetData><row r="2"><c r="B2"><v>1</v></c>'
                '<c r="C2"><f>B2*100</f><v>10</v></c></row></sheetData>')
        root_attrs = (' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
                      ' xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"')
        spark = ('<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">'
                 '<x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline>'
                 '<xm:f>Forecast!B2:C2</xm:f><xm:sqref>C3</xm:sqref>'
                 '</x14:sparkline></x14:sparklines></x14:sparklineGroup>'
                 '</x14:sparklineGroups></ext></extLst>')
        data = fixture(sheet=self.sheet(body, root_attrs=root_attrs, tail=spark))
        output, manifest = self.build(data)
        self.assertEqual(manifest["status"], "needs_recalculation")
        candidate = Package(output)
        self.fragments_equal(Package(data), candidate)
        sheet_bytes = candidate.parts["xl/worksheets/sheet1.xml"]
        self.assertIn(b"<xm:f>Forecast!B2:C2</xm:f>", sheet_bytes)
        self.assertIn(b"<x14:sparkline>", sheet_bytes)

    def test_prefixed_cell_inside_sheetdata_is_still_blocked(self):
        data = fixture(sheet=f'<worksheet xmlns="{NS}" xmlns:xm="{NS}"><sheetData><row r="2"><xm:c r="B2"><v>1</v></xm:c></row></sheetData></worksheet>'.encode())
        with self.assertRaises(ModelError) as caught:
            self.build(data)
        self.assertIn("Nonstandard worksheet XML", str(caught.exception))

    def test_prefixed_calcpr_is_updated_not_duplicated(self):
        data = fixture(calc=f'<ns0:calcPr xmlns:ns0="{NS}" calcId="99" calcMode="manual"/>')
        output, _ = self.build(data)
        workbook = Package(output).parts["xl/workbook.xml"]
        self.assertEqual(workbook.count(b"calcPr"), 1)
        self.assertIn(b'calcId="99"', workbook)
        self.assertIn(b'calcMode="auto"', workbook)
        self.assertIn(b'fullCalcOnLoad="1"', workbook)
        self.assertNotIn(b'calcMode="manual"', workbook)
        # both a prefixed and an unprefixed calcPr are refused as duplicates
        data = fixture(calc=f'<ns0:calcPr xmlns:ns0="{NS}" calcId="99"/><calcPr calcId="100"/>')
        with self.assertRaises(ModelError) as caught:
            self.build(data)
        self.assertIn("Duplicate calculation properties", str(caught.exception))


MARKSANS_SPECIAL = {
    "Main Financials": {"shared": 124},
    "Ratio Analysis": {"shared": 60},
    "Intrinsic Growth": {"shared": 50},
    "Comps": {"shared": 45, "array": 3},
    "WACC-DCF": {"shared": 15},
    "DCF 1st way": {"shared": 27, "dataTable": 1},
    "DCF 2nd way": {"shared": 168, "array": 9},
    "Data Sheet": {"shared": 18},
    "Beta": {"shared": 207},
    "compdata": {"shared": 28},
}


class MarksansIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(__file__).resolve().parent
        cls.original = cls.root / "originals" / "Marksans Keshav.xlsx"
        cls.data = cls.original.read_bytes()
        cls.before = sha(cls.data)

    def real_pack(self):
        return {
            "schema_version": 1,
            "template_sha256": sha(self.data),
            "approval": {"reviewer": "proposal reviewer (fixture)", "approved_at": "2026-10-16T00:00:00Z",
                         "template_sha256": sha(self.data)},
            "inputs": [{"name": "terminal_growth", "sheet": "DCF 2nd way", "cell": "D10",
                        "kind": "number", "unit": "fraction per year", "minimum": 0, "maximum": 0.5}],
            "checks": [{"name": "wacc_cell_exists", "sheet": "DCF 2nd way", "cell": "D9",
                        "expected": 0.121, "abs_tolerance": 0.001}],
        }

    def real_inputs(self, value=Decimal("0.065")):
        return {"values": {"terminal_growth": {"value": value, "source": "Proposed terminal growth assumption (unapproved proposal)"}}}

    def test_marksans_inventory_matches_recorded_structure(self):
        package = Package(self.data)
        self.assertEqual(package.special_formulas(), MARKSANS_SPECIAL)
        totals = {"shared": 0, "array": 0, "dataTable": 0}
        for counts in MARKSANS_SPECIAL.values():
            for kind, count in counts.items():
                totals[kind] = totals.get(kind, 0) + count
        self.assertEqual(totals, {"shared": 742, "array": 12, "dataTable": 1})

    def test_marksans_passes_adapter_gate_without_other_blockers(self):
        self.assertEqual(Package(self.data).blockers(), [])

    def test_marksans_build_preserves_every_formula_fragment(self):
        source = Package(self.data)
        output, manifest = compose(self.data, self.real_pack(), self.real_inputs())
        self.assertEqual(manifest["status"], "needs_recalculation")
        self.assertFalse(manifest["ready"])
        candidate = Package(output)
        self.assertEqual(candidate.blockers(), [])
        self.assertEqual(candidate.special_formulas(), source.special_formulas())
        for name, (part, _) in source.sheets.items():
            self.assertEqual(
                formula_adapter.formula_fragments(source.parts[part]),
                formula_adapter.formula_fragments(candidate.parts[part]),
                f"formula fragments changed on {name}")
        # the mapped input was applied
        sheet8 = candidate.parts[source.sheets["DCF 2nd way"][0]]
        self.assertIn(b'<c r="D10" s="287"><v>0.065</v></c>', sheet8)
        # protected-zone stale caches were invalidated
        self.assertIn("K13", manifest["stale_cache_cells_stripped"])
        self.assertIn("I25", manifest["stale_cache_cells_stripped"])
        self.assertNotIn(b'<v>3695.2', sheet8)
        # original file is untouched
        self.assertEqual(sha(self.original.read_bytes()), self.before)

    def test_marksans_rejects_input_inside_protected_zone(self):
        p = self.real_pack()
        p["inputs"][0]["cell"] = "K13"  # inside array zone J13:N13
        with self.assertRaises(ModelError) as caught:
            compose(self.data, p, self.real_inputs())
        self.assertIn("protected", str(caught.exception))

    def test_marksans_rejects_out_of_bounds_input(self):
        p = self.real_pack()
        p["inputs"][0]["minimum"] = 0.1
        with self.assertRaises(ModelError) as caught:
            compose(self.data, p, self.real_inputs())
        self.assertIn("bounds", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
