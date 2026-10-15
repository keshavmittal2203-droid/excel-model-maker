"""Synthetic OOXML fixtures exercise safeguards; these are not Excel certification."""
from copy import deepcopy
from decimal import Decimal
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from modelmaker import ModelError, Package, compose, coordinate, numeric, sha, NS, REL, read_json, verify_sheet_changes


def fixture(extra=None, sheet=None, calc=''):
    parts = {
        '[Content_Types].xml': b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
        'xl/workbook.xml': f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets><sheet name="Forecast" sheetId="1" r:id="rId1"/></sheets>{calc}<extLst/></workbook>'.encode(),
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/calcChain" Target="calcChain.xml"/></Relationships>'.encode(),
        'xl/worksheets/sheet1.xml': sheet or f'<worksheet xmlns="{NS}"><sheetData><row r="2"><c r="B2" s="1"><v>0.1</v></c><c r="C2" s="2"><f>B2*100</f><v>10</v></c><c r="D2" t="e"><f>1/B2</f><v>#DIV/0!</v></c></row></sheetData><mergeCells><mergeCell ref="A5:C5"/></mergeCells></worksheet>'.encode(),
        'xl/calcChain.xml': f'<calcChain xmlns="{NS}"/>'.encode(),
        'xl/styles.xml': f'<styleSheet xmlns="{NS}"/>'.encode(),
        'xl/charts/chart1.xml': b'<chartSpace xmlns="urn:test">unchanged native part</chartSpace>',
        'docProps/custom.xml': b'<properties/>',
    }
    parts.update(extra or {})
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return stream.getvalue()


def pack(data):
    return {'schema_version': 1, 'template_sha256': sha(data),
            'approval': {'reviewer': 'fixture reviewer', 'approved_at': '2026-10-16T00:00:00Z', 'template_sha256': sha(data)},
            'inputs': [{'name': 'growth', 'sheet': 'Forecast', 'cell': 'B2', 'kind': 'number', 'unit': 'fraction', 'minimum': 0, 'maximum': 1}],
            'checks': [{'name': 'revenue', 'sheet': 'Forecast', 'cell': 'C2', 'expected': 20, 'abs_tolerance': 0.01}]}


def inputs(value=Decimal('0.2')):
    return {'values': {'growth': {'value': value, 'source': 'Explicit user forecast assumption'}}}


class CoreTests(unittest.TestCase):
    def setUp(self):
        self.data = fixture()
        self.pack = pack(self.data)

    def build(self, **kwargs):
        return compose(kwargs.get('data', self.data), kwargs.get('pack', self.pack), kwargs.get('inputs', inputs()))

    def test_inspection_is_read_only_and_detects_cached_error(self):
        result = Package(self.data).summary()
        self.assertEqual(result['sheets'][0]['formulas'], 2)
        self.assertEqual(result['sheets'][0]['cached_errors'][0]['cell'], 'D2')
        self.assertEqual(result['sha256'], sha(self.data))
        self.assertEqual(result['financial_correctness'], 'not_evaluated')

    def test_build_preserves_all_unrelated_parts_and_formula_text(self):
        output, manifest = self.build()
        source, candidate = Package(self.data), Package(output)
        for name in ('xl/styles.xml', 'xl/charts/chart1.xml', 'docProps/custom.xml'):
            self.assertEqual(source.parts[name], candidate.parts[name])
        sheet = candidate.parts['xl/worksheets/sheet1.xml']
        self.assertIn(b'<c r="B2" s="1"><v>0.2</v></c>', sheet)
        self.assertIn(b'<c r="C2" s="2"><f>B2*100</f></c>', sheet)
        self.assertIn(b'<c r="D2"><f>1/B2</f></c>', sheet)
        self.assertNotIn('xl/calcChain.xml', candidate.parts)
        self.assertNotIn(b'calcChain', candidate.parts['xl/_rels/workbook.xml.rels'])
        self.assertNotIn(b'calcChain', candidate.parts['[Content_Types].xml'])
        self.assertIn(b'forceFullCalc="1"', candidate.parts['xl/workbook.xml'])
        self.assertLess(candidate.parts['xl/workbook.xml'].index(b'<calcPr'), candidate.parts['xl/workbook.xml'].index(b'<extLst'))
        self.assertEqual(manifest['output_sha256'], sha(output))
        self.assertEqual(manifest['status'], 'needs_recalculation')
        self.assertFalse(manifest['ready'])
        self.assertEqual(manifest['changed_inputs'][0]['source'], inputs()['values']['growth']['source'])

    def test_self_closing_input_does_not_consume_following_formula(self):
        data = fixture(sheet=f'<worksheet xmlns="{NS}"><sheetData><row r="2"><c s="1" r="B2"/><c r="C2"><f>B2*100</f><v>10</v></c></row></sheetData></worksheet>'.encode())
        output, _ = self.build(data=data, pack=pack(data))
        sheet = Package(output).parts['xl/worksheets/sheet1.xml']
        self.assertIn(b'<c s="1" r="B2"><v>0.2</v></c><c r="C2"><f>B2*100</f></c>', sheet)

    def test_existing_calc_properties_are_updated_without_losing_settings(self):
        data = fixture(calc='<calcPr calcMode="manual" fullCalcOnLoad="0" calcId="123"/>')
        output, _ = self.build(data=data, pack=pack(data))
        workbook = Package(output).parts['xl/workbook.xml']
        self.assertIn(b'calcId="123"', workbook)
        self.assertIn(b'calcMode="auto"', workbook)
        self.assertNotIn(b'calcMode="manual"', workbook)

    def test_wrong_template_or_approval_hash_is_rejected(self):
        for field in ('template', 'approval'):
            p = deepcopy(self.pack)
            if field == 'template': p['template_sha256'] = '0' * 64
            else: p['approval']['template_sha256'] = '0' * 64
            with self.assertRaises(ModelError): self.build(pack=p)

    def test_missing_approval_or_checks_is_rejected(self):
        for field in ('approval', 'checks'):
            p = deepcopy(self.pack)
            del p[field]
            with self.assertRaises(ModelError): self.build(pack=p)

    def test_nonfinite_boolean_string_and_excess_precision_rejected(self):
        for value in (True, '0.2', float('nan'), float('inf'), Decimal('0.1234567890123456'), Decimal('1e308'), 123456789012345678901234567891):
            with self.subTest(value=str(value)), self.assertRaises(ModelError):
                self.build(inputs=inputs(value))

    def test_bounds_and_provenance_are_enforced(self):
        for value in (-1, 2):
            with self.assertRaises(ModelError): self.build(inputs=inputs(value))
        i = inputs()
        i['values']['growth']['source'] = ' '
        with self.assertRaises(ModelError): self.build(inputs=i)

    def test_missing_and_extra_inputs_are_rejected(self):
        for value in ({'values': {}}, {'values': {**inputs()['values'], 'extra': 1}}):
            with self.assertRaises(ModelError): self.build(inputs=value)

    def test_duplicate_names_and_destinations_are_rejected(self):
        for mode in ('name', 'destination'):
            p = deepcopy(self.pack)
            p['inputs'].append(deepcopy(p['inputs'][0]))
            if mode == 'destination': p['inputs'][1]['name'] = 'another'
            with self.assertRaises(ModelError): self.build(pack=p)

    def test_formula_cell_and_missing_cell_are_rejected(self):
        for ref in ('C2', 'B3', 'A5'):
            p = deepcopy(self.pack)
            p['inputs'][0]['cell'] = ref
            with self.assertRaises(ModelError): self.build(pack=p)

    def test_merged_input_is_rejected(self):
        data = fixture(sheet=f'<worksheet xmlns="{NS}"><sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2"><v>2</v></c></row></sheetData><mergeCells><mergeCell ref="A2:B2"/></mergeCells></worksheet>'.encode())
        with self.assertRaises(ModelError): self.build(data=data, pack=pack(data))

    def test_native_features_are_blocked_until_supported(self):
        for part in ('xl/vbaProject.bin', '_xmlsignatures/sig1.xml', 'xl/externalLinks/externalLink1.xml', 'xl/connections.xml', 'xl/queryTables/queryTable1.xml'):
            data = fixture(extra={part: b'<test/>'})
            with self.subTest(part=part), self.assertRaises(ModelError): self.build(data=data, pack=pack(data))

    def test_shared_array_data_table_and_broken_formulas_are_blocked(self):
        for formula in ('<f t="shared" si="0">B2*100</f>', '<f t="array" ref="C2">B2*100</f>', '<f t="dataTable"/>', '<f>#REF!*100</f>'):
            sheet = f'<worksheet xmlns="{NS}"><sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2">{formula}<v>2</v></c></row></sheetData></worksheet>'.encode()
            data = fixture(sheet=sheet)
            with self.subTest(formula=formula), self.assertRaises(ModelError): self.build(data=data, pack=pack(data))

    def test_protection_and_iteration_are_blocked(self):
        data = fixture(calc='<calcPr iterate="1"/>')
        with self.assertRaises(ModelError): self.build(data=data, pack=pack(data))
        sheet = f'<worksheet xmlns="{NS}"><sheetProtection sheet="1"/><sheetData/></worksheet>'.encode()
        data = fixture(sheet=sheet)
        with self.assertRaises(ModelError): self.build(data=data, pack=pack(data))

    def test_zip_traversal_duplicates_and_invalid_zip_are_rejected(self):
        with self.assertRaises(ModelError): Package(b'not a zip')
        data = fixture(extra={'../escape.xml': b'<test/>'})
        with self.assertRaises(ModelError): Package(data)
        stream = io.BytesIO(self.data)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            with zipfile.ZipFile(stream, 'a') as archive:
                archive.writestr('xl/styles.xml', b'<test/>')
        with self.assertRaises(ModelError): Package(stream.getvalue())

    def test_dtd_and_duplicate_cells_are_rejected(self):
        data = fixture(extra={'docProps/custom.xml': b'<!DOCTYPE x [<!ENTITY y "z">]><x>&y;</x>'})
        with self.assertRaises(ModelError): Package(data)
        data = fixture(sheet=f'<worksheet xmlns="{NS}"><sheetData><row r="2"><c r="B2"/><c r="B2"/></row></sheetData></worksheet>'.encode())
        with self.assertRaises(ModelError): Package(data)

    def test_cell_bounds(self):
        self.assertEqual(coordinate('XFD1048576'), (16384, 1048576))
        for ref in ('XFE1', 'A1048577', '$B$2', 'B0', 'B2:C2', 'b2'):
            with self.subTest(ref=ref), self.assertRaises(ModelError): coordinate(ref)

    def test_invalid_approval_timestamp_and_boolean_version_rejected(self):
        for timestamp in ('yesterday', '2026-10-16', '2026-10-16T00:00:00'):
            p = deepcopy(self.pack)
            p['approval']['approved_at'] = timestamp
            with self.assertRaises(ModelError): self.build(pack=p)
        p = deepcopy(self.pack)
        p['schema_version'] = True
        with self.assertRaises(ModelError): self.build(pack=p)

    def test_duplicate_json_keys_and_nonstandard_constants_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'bad.json'
            for text in ('{"x":1,"x":2}', '{"x":NaN}', '{"x":Infinity}'):
                path.write_text(text)
                with self.assertRaises(ModelError): read_json(str(path))

    def test_semantic_guard_rejects_unapproved_formula_and_style_edits(self):
        original = Package(self.data).parts['xl/worksheets/sheet1.xml']
        for changed in (original.replace(b'B2*100', b'B2*200'), original.replace(b's="1"', b's="3"')):
            with self.assertRaises(ModelError): verify_sheet_changes(original, changed, {'B2'})

    def test_pretty_printed_worksheet_is_supported(self):
        sheet = Package(self.data).parts['xl/worksheets/sheet1.xml'].replace(b'><', b'>\n<')
        data = fixture(sheet=sheet)
        output, report = self.build(data=data, pack=pack(data))
        self.assertEqual(report['status'], 'needs_recalculation')
        self.assertEqual(Package(output).summary()['sheets'][0]['formulas'], 2)

    def test_cli_refuses_overwrite_and_emits_blocked_status(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template = root / 'template.xlsx'
            template.write_bytes(self.data)
            (root / 'pack.json').write_text(json.dumps(self.pack))
            (root / 'inputs.json').write_text(json.dumps(inputs(0.2)))
            command = [sys.executable, str(Path(__file__).with_name('modelmaker.py')), 'build', '--template', str(template), '--pack', str(root / 'pack.json'), '--inputs', str(root / 'inputs.json'), '--output', str(template)]
            result = subprocess.run(command, capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            self.assertEqual(json.loads(result.stdout)['status'], 'blocked')
            self.assertEqual(template.read_bytes(), self.data)
            output = root / 'candidate.xlsx'
            command[-1] = str(output)
            first = subprocess.run(command, capture_output=True, text=True)
            self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
            before = output.read_bytes()
            second = subprocess.run(command, capture_output=True, text=True)
            self.assertEqual(second.returncode, 2)
            self.assertEqual(output.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
