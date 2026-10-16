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

    def test_malformed_and_broken_formulas_are_blocked(self):
        for formula in ('<f t="shared" si="0">B2*100</f>', '<f t="array">B2*100</f>', '<f t="dataTable"/>', '<f>#REF!*100</f>', '<f t="unsupported">B2*100</f>'):
            sheet = f'<worksheet xmlns="{NS}"><sheetData><row r="2"><c r="B2"><v>1</v></c><c r="C2">{formula}<v>2</v></c></row></sheetData></worksheet>'.encode()
            data = fixture(sheet=sheet)
            with self.subTest(formula=formula), self.assertRaises(ModelError): self.build(data=data, pack=pack(data))

    def test_special_formulas_are_supported_and_preserved(self):
        sheet = (
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2">'
            f'<c r="B2"><v>1</v></c>'
            f'<c r="C2"><f t="shared" ref="C2:C3" si="0">B2*10</f><v>10</v></c>'
            f'<c r="D2"><f t="array" ref="D2">B2*20</f><v>20</v></c>'
            f'<c r="E2"><f t="dataTable" ref="E2:F3" dt2D="1" dtr="1" r1="B2" r2="B2" ca="1"/><v>30</v></c>'
            f'</row>'
            f'<row r="3">'
            f'<c r="C3"><f t="shared" si="0"/><v>10</v></c>'
            f'<c r="E3"><v>40</v></c>'
            f'</row>'
            f'</sheetData></worksheet>'
        ).encode()
        data = fixture(sheet=sheet)
        output, report = self.build(data=data, pack=pack(data), inputs=inputs(0.5))
        self.assertEqual(report['status'], 'needs_recalculation')
        cand_pkg = Package(output)
        cand_sheet = cand_pkg.parts['xl/worksheets/sheet1.xml']
        self.assertIn(b'<f t="shared" ref="C2:C3" si="0">B2*10</f>', cand_sheet)
        self.assertIn(b'<f t="shared" si="0"/>', cand_sheet)
        self.assertIn(b'<f t="array" ref="D2">B2*20</f>', cand_sheet)
        self.assertIn(b'<f t="dataTable" ref="E2:F3" dt2D="1" dtr="1" r1="B2" r2="B2" ca="1"/>', cand_sheet)
        self.assertNotIn(b'<c r="C2"><f t="shared" ref="C2:C3" si="0">B2*10</f><v>', cand_sheet)
        self.assertNotIn(b'<c r="C3"><f t="shared" si="0"/><v>', cand_sheet)
        self.assertNotIn(b'<c r="D2"><f t="array" ref="D2">B2*20</f><v>', cand_sheet)
        self.assertIn(b'<c r="B2"><v>0.5</v></c>', cand_sheet)

    def test_mapped_inputs_colliding_with_formula_ranges_are_rejected(self):
        sheet = (
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2">'
            f'<c r="B2"><v>1</v></c>'
            f'<c r="C2"><f t="shared" ref="C2:C3" si="0">B2*10</f><v>10</v></c>'
            f'<c r="D2"><f t="array" ref="D2:D3">B2*20</f><v>20</v></c>'
            f'<c r="E2"><f t="dataTable" ref="E2:E3" dt2D="1" dtr="1" r1="B2" r2="B2" ca="1"/><v>30</v></c>'
            f'</row>'
            f'<row r="3">'
            f'<c r="C3"><f t="shared" si="0"/><v>10</v></c>'
            f'<c r="D3"><v>20</v></c>'
            f'<c r="E3"><v>30</v></c>'
            f'</row>'
            f'</sheetData></worksheet>'
        ).encode()
        data = fixture(sheet=sheet)
        for target in ('C2', 'C3', 'D2', 'D3', 'E2', 'E3'):
            p = deepcopy(self.pack)
            p['inputs'][0]['cell'] = target
            with self.subTest(target=target), self.assertRaises(ModelError):
                self.build(data=data, pack=p)

    def test_marksans_workbook_compatibility_and_composition(self):
        marksans_path = Path(__file__).resolve().parent / 'originals' / 'Marksans Keshav.xlsx'
        if not marksans_path.exists():
            return
        data = marksans_path.read_bytes()
        pkg = Package(data)
        self.assertEqual(pkg.blockers(), [])
        summary = pkg.summary()
        self.assertEqual(len(summary['special_formulas']), 10)
        self.assertEqual(summary['unsupported_features'], [])

        digest = sha(data)
        m_pack = {
            'schema_version': 1,
            'template_sha256': digest,
            'approval': {
                'template_sha256': digest,
                'reviewer': 'Alice Independent Reviewer',
                'approved_at': '2026-10-16T01:30:00+05:30'
            },
            'inputs': [
                {
                    'name': 'tax_rate',
                    'sheet': 'Intrinsic Growth',
                    'cell': 'E47',
                    'kind': 'number',
                    'unit': 'ratio',
                    'minimum': 0,
                    'maximum': 1
                }
            ],
            'checks': [
                {
                    'name': 'sales_check',
                    'sheet': 'Data Sheet',
                    'cell': 'B8',
                    'expected': Decimal('325.6'),
                    'abs_tolerance': Decimal('0.01')
                }
            ]
        }
        m_inputs = {
            'values': {
                'tax_rate': {
                    'value': Decimal('0.28'),
                    'source': 'Revised tax assumption'
                }
            }
        }
        output, manifest = compose(data, m_pack, m_inputs)
        self.assertEqual(manifest['status'], 'needs_recalculation')
        self.assertFalse(manifest['ready'])
        out_pkg = Package(output)
        self.assertEqual(out_pkg.blockers(), [])

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

    def test_marksans_full_93_input_pack_and_identities(self):
        root_dir = Path(__file__).resolve().parent
        marksans_path = root_dir / 'originals' / 'Marksans Keshav.xlsx'
        pack_path = root_dir / 'marksans.model-pack.json'
        inputs_path = root_dir / 'marksans.inputs.json'
        if not marksans_path.exists() or not pack_path.exists():
            return
        from validate_workbook import (
            evaluate_pack_checks,
            run_equity_research_identities,
            scan_formula_errors_and_missing_caches,
        )
        data = marksans_path.read_bytes()
        m_pack = read_json(str(pack_path))
        m_inputs = read_json(str(inputs_path))
        self.assertEqual(len(m_pack['inputs']), 93)
        self.assertEqual(len(m_pack['checks']), 19)

        # Verify all 29 independent identities and 19 pack checks hold on baseline cached values
        orig_pkg = Package(data)
        errors, missing = scan_formula_errors_and_missing_caches(orig_pkg)
        self.assertEqual(errors, [])
        self.assertEqual(missing, [])
        p_checks = evaluate_pack_checks(orig_pkg, m_pack)
        self.assertTrue(all(c['passed'] for c in p_checks), p_checks)
        identities = run_equity_research_identities(orig_pkg)
        self.assertEqual(len(identities), 29)
        self.assertTrue(all(i['passed'] for i in identities), identities)

        # Verify compose strips formula caches so un-recalculated output is not prematurely validated
        cand_bytes, manifest = compose(data, m_pack, m_inputs)
        self.assertEqual(manifest['status'], 'needs_recalculation')
        self.assertFalse(manifest['ready'])
        cand_pkg = Package(cand_bytes)
        c_errors, c_missing = scan_formula_errors_and_missing_caches(cand_pkg)
        self.assertEqual(c_errors, [])
        self.assertGreater(len(c_missing), 100)

    def test_validate_candidate_recalculation_and_perturbation_gates(self):
        from validate_workbook import validate_candidate
        clean_sheet = (
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="B2" s="1"><v>0.1</v></c><c r="C2" s="2"><f>B2*100</f><v>10</v></c></row>'
            f'</sheetData></worksheet>'
        ).encode()
        clean_data = fixture(sheet=clean_sheet)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template = root / 'template.xlsx'
            cand = root / 'candidate.xlsx'
            recalc = root / 'recalculated.xlsx'
            pack_file = root / 'pack.json'
            inputs_file = root / 'inputs.json'

            template.write_bytes(clean_data)
            p = pack(clean_data)
            p['checks'][0]['expected'] = 15
            pack_file.write_text(json.dumps(p, default=float))
            inputs_file.write_text(json.dumps(inputs(0.15), default=float))

            cand_bytes, _ = self.build(data=clean_data, pack=p, inputs=inputs(0.15))
            cand.write_bytes(cand_bytes)

            # Without recalculation, candidate remains needs_recalculation (ready=False)
            unrecalc_report = validate_candidate(cand, pack_file, recalculate=False)
            self.assertEqual(unrecalc_report['status'], 'needs_recalculation')
            self.assertFalse(unrecalc_report['ready'])
            self.assertGreater(unrecalc_report['uncalculated_formula_count'], 0)

            # With recalculation and perturbation check, candidate transitions to validated (ready=True)
            valid_report = validate_candidate(
                cand,
                pack_file,
                recalculate=True,
                recalculated_output=recalc,
                perturb_check=True,
                template_path=template,
                inputs_path=inputs_file,
                use_python_recalc_fallback=True,
            )
            self.assertEqual(valid_report['status'], 'recalculated_checks_passed')
            self.assertFalse(valid_report['ready'])
            self.assertTrue(valid_report['checks_passed'])
            self.assertIn('render_review', valid_report['pending_gates'])
            self.assertEqual(valid_report['formula_error_count'], 0)
            self.assertEqual(valid_report['uncalculated_formula_count'], 0)
            self.assertTrue(valid_report['perturbation_and_restoration']['passed'])
            self.assertTrue(valid_report['perturbation_and_restoration']['restoration_byte_identical'])

    def test_shared_formula_validation_rejects_malformed_structures(self):
        # 1. Follower referencing non-existent master si
        bad_follower_si = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData><row r="2">'
            f'<c r="B2"><v>10</v></c><c r="C2"><f t="shared" si="99"/></c>'
            f'</row></sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=bad_follower_si, pack=pack(bad_follower_si))
        self.assertIn("Shared formula follower references missing group", str(cm.exception))

        # 2. Follower outside master ref range
        follower_outside_ref = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="shared" si="0" ref="C2:C3">B2*10</f></c></row>'
            f'<row r="4"><c r="C4"><f t="shared" si="0"/></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=follower_outside_ref, pack=pack(follower_outside_ref))
        self.assertIn("outside master ref range", str(cm.exception))

        # 3. Follower with non-empty formula text
        follower_with_text = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="shared" si="0" ref="C2:C3">B2*10</f></c></row>'
            f'<row r="3"><c r="C3"><f t="shared" si="0">B3*10</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=follower_with_text, pack=pack(follower_with_text))
        self.assertIn("Shared formula follower contains unexpected formula text", str(cm.exception))

        # 4. Master with reversed range
        reversed_range = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="3"><c r="C3"><f t="shared" si="0" ref="C3:B2">B3*10</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=reversed_range, pack=pack(reversed_range))
        self.assertIn("Invalid shared formula range", str(cm.exception))

        # 5. Master not at top-left
        master_not_top_left = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="3"><c r="C3"><f t="shared" si="0" ref="C2:C4">B3*10</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=master_not_top_left, pack=pack(master_not_top_left))
        self.assertIn("Shared formula master not at top-left of ref range", str(cm.exception))

        # 6. Duplicate master si
        dup_master = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="shared" si="0" ref="C2:C3">B2*10</f></c></row>'
            f'<row r="5"><c r="C5"><f t="shared" si="0" ref="C5:C6">B5*10</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=dup_master, pack=pack(dup_master))
        self.assertIn("Duplicate shared formula group index", str(cm.exception))

        # 7. Shared master with unexpected attributes
        unexpected_attr = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="shared" si="0" ref="C2:C3" bogus="1">B2*10</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=unexpected_attr, pack=pack(unexpected_attr))
        self.assertIn("Unsupported attributes on shared master", str(cm.exception))

    def test_array_and_data_table_formula_validation_rejects_malformed_structures(self):
        # 1. Array formula master not at top-left
        array_not_top_left = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="3"><c r="C3"><f t="array" ref="C2:C4">B2:B4*2</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=array_not_top_left, pack=pack(array_not_top_left))
        self.assertIn("Array formula master not at top-left of ref range", str(cm.exception))

        # 2. Array formula with unexpected attributes
        array_bad_attr = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="array" ref="C2:C4" si="0">B2:B4*2</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=array_bad_attr, pack=pack(array_bad_attr))
        self.assertIn("Unsupported attributes on array formula", str(cm.exception))

        # 3. Overlapping array formulas
        overlap_array = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="array" ref="C2:C4">B2:B4*2</f></c></row>'
            f'<row r="3"><c r="C3"><f t="array" ref="C3:D5">B3:B5*3</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=overlap_array, pack=pack(overlap_array))
        self.assertIn("Overlapping special formula range", str(cm.exception))

        # 4. DataTable formula missing required attributes
        dt_missing_attrs = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="dataTable" ref="C2:D4"/></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=dt_missing_attrs, pack=pack(dt_missing_attrs))
        self.assertIn("Malformed dataTable attributes", str(cm.exception))

        # 5. DataTable formula with non-empty formula text
        dt_non_empty = fixture(sheet=(
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="C2"><f t="dataTable" ref="C2:D4" dt2D="1" dtr="1" r1="A1" r2="A2">SUM(A1:A2)</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        ))
        with self.assertRaises(ModelError) as cm:
            self.build(data=dt_non_empty, pack=pack(dt_non_empty))
        self.assertIn("dataTable formula must not contain text", str(cm.exception))

    def test_pack_approval_disallowed_reviewers_and_proposals(self):
        # 1. Disallowed reviewer token: 'antigravity'
        bad_reviewer_pack = deepcopy(self.pack)
        bad_reviewer_pack["approval"]["reviewer"] = "Antigravity AI Agent"
        with self.assertRaises(ModelError) as cm:
            self.build(pack=bad_reviewer_pack)
        self.assertIn("Pack approval reviewer cannot be an agent or placeholder", str(cm.exception))

        # 2. Proposal status without allow_proposal flag
        proposal_pack = deepcopy(self.pack)
        proposal_pack["status"] = "proposal"
        proposal_pack["approval"] = None
        with self.assertRaises(ModelError) as cm:
            self.build(pack=proposal_pack)
        self.assertIn("Pack is an unapproved proposal", str(cm.exception))

        # 3. Proposal status with allow_proposal=True
        out, manifest = compose(self.data, proposal_pack, inputs(), allow_proposal=True)
        self.assertEqual(manifest["status"], "needs_recalculation")
        self.assertFalse(manifest["ready"])

    def test_python_evaluator_rejects_unsupported_and_special_formulas(self):
        from validate_workbook import recalculate_package_python
        # Unsupported function FORECAST
        unsupported_fn_sheet = (
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="B2"><v>10</v></c><c r="C2"><f>FORECAST(B2, 1, 2)</f></c></row>'
            f'</sheetData></worksheet>'.encode()
        )
        with self.assertRaises(ModelError) as cm:
            recalculate_package_python(fixture(sheet=unsupported_fn_sheet))
        self.assertIn("Unsupported function in Python evaluator: FORECAST", str(cm.exception))

        # Special formula (e.g. shared) cannot be evaluated by partial Python evaluator
        shared_sheet = (
            f'<worksheet xmlns="{NS}"><sheetData>'
            f'<row r="2"><c r="B2"><v>10</v></c><c r="C2"><f t="shared" si="0" ref="C2:C3">B2*10</f></c></row>'
            f'<row r="3"><c r="B3"><v>20</v></c><c r="C3"><f t="shared" si="0"/></c></row>'
            f'</sheetData></worksheet>'.encode()
        )
        with self.assertRaises(ModelError) as cm:
            recalculate_package_python(fixture(sheet=shared_sheet))
        self.assertIn("rejects workbooks with shared, array, or dataTable formulas", str(cm.exception))

    def test_tolerance_rejects_hidden_relative_error_fallback(self):
        from validate_workbook import _compare_metric
        # Exceeds abs_tolerance (0.5), relative difference is ~1%. Must NOT pass.
        metric = _compare_metric(
            name="test_strict_tol",
            computed=Decimal("100.0"),
            expected=Decimal("101.0"),
            abs_tolerance=Decimal("0.5"),
            rel_tolerance=None,
        )
        self.assertFalse(metric["passed"])
        self.assertEqual(metric["abs_diff"], "1.0")


if __name__ == '__main__':
    unittest.main()

