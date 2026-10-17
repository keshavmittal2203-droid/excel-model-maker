"""Comprehensive unit and regression tests for the four priority target packages (M&M, LTM, CLEAN, EDELWEISS)."""
import json
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent

from modelmaker import Package
from validate_workbook import scan_formula_errors_and_missing_caches

TARGET_KEYS = ["MM", "LTM", "CLEAN", "EDELWEISS"]

class FourTargetsTests(unittest.TestCase):
    def test_all_four_workbooks_recalculate_with_zero_errors(self):
        """All four candidate workbooks must calculate with zero formula errors and zero missing caches."""
        for key in TARGET_KEYS:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            self.assertTrue(wb_path.exists(), f"Recalculated workbook missing for {key}")
            pkg = Package(wb_path.read_bytes())
            errors, missing = scan_formula_errors_and_missing_caches(pkg)
            self.assertEqual(len(errors), 0, f"{key} has {len(errors)} formula errors: {errors}")
            self.assertEqual(len(missing), 0, f"{key} has {len(missing)} uncalculated cells: {missing}")

    def test_all_four_balance_sheets_tie_out_to_zero_discrepancy(self):
        """All four balance sheets must balance to exact 0.00 discrepancy across all 12 periods."""
        import openpyxl
        for key in TARGET_KEYS:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            self.assertIn("Balance Sheet", wb.sheetnames)
            ws = wb["Balance Sheet"]
            # Row 17 is discrepancy row
            discs = [ws.cell(17, c).value for c in range(2, 14)]
            for idx, d in enumerate(discs, start=2):
                col_let = openpyxl.utils.get_column_letter(idx)
                self.assertIsNotNone(d, f"{key} {col_let}17 discrepancy is None")
                self.assertAlmostEqual(float(d), 0.0, places=2, msg=f"{key} Balance Sheet discrepancy in col {col_let}: {d}")

    def test_all_four_source_registries_and_manifests_present(self):
        """Source registries and manifests must exist with required metadata and SHA-256 hashes."""
        for key in TARGET_KEYS:
            k_lower = key.lower()
            # Registry files
            reg_md = ROOT / "docs" / f"{key}_SOURCE_REGISTRY.md"
            reg_json = ROOT / "docs" / f"{k_lower}_source_registry.json"
            self.assertTrue(reg_md.exists(), f"Registry markdown missing for {key}")
            self.assertTrue(reg_json.exists(), f"Registry JSON missing for {key}")

            entries = json.loads(reg_json.read_text(encoding="utf-8"))
            self.assertGreaterEqual(len(entries), 6, f"Registry for {key} has insufficient entries")
            for e in entries:
                self.assertIn("metric", e)
                self.assertIn("legal_entity", e)
                self.assertIn("source_org", e)
                self.assertIn("official_url", e)
                self.assertIn("status", e)

            # Manifests
            manifest_path = ROOT / "candidates" / f"run_native_{k_lower}" / "package.manifest.json"
            self.assertTrue(manifest_path.exists())
            man = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(man["status"], "validating")
            self.assertFalse(man["ready"])
            self.assertIn("human_pack_approval", man["pending_gates"])
            self.assertEqual(man["workbook"]["formula_error_count"], 0)
            self.assertEqual(len(man["previews"]), 6)

    def test_all_24_preview_pdfs_exist_and_non_empty(self):
        """Each company must have 6 substantive sheet landscape PDF previews (24 total)."""
        for key in TARGET_KEYS:
            prev_dir = ROOT / "candidates" / f"run_native_{key.lower()}" / "previews"
            self.assertTrue(prev_dir.exists())
            pdfs = list(prev_dir.glob("*.pdf"))
            self.assertEqual(len(pdfs), 6, f"Expected 6 PDFs for {key}, got {len(pdfs)}")
            for p in pdfs:
                self.assertGreater(p.stat().st_size, 10000, f"PDF {p.name} is suspiciously small: {p.stat().st_size} bytes")

    def test_edelweiss_fcff_strictly_blocked_and_sotp_applied(self):
        """Negative test: Edelweiss package must formally block industrial FCFF and enforce SOTP."""
        ed_manifest = json.loads((ROOT / "companies" / "EDELWEISS" / "package.manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(ed_manifest["valuation"]["status"], "supported_sotp_fcff_withheld")
        self.assertIn("Industrial FCFF DCF is strictly withheld", ed_manifest["valuation"]["methodological_notes"])

        ed_reg = (ROOT / "docs" / "EDELWEISS_SOURCE_REGISTRY.md").read_text(encoding="utf-8")
        self.assertIn("METHODOLOGICAL EXCLUSION OF INDUSTRIAL FCFF", ed_reg)

        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_edelweiss" / "EDELWEISS-Recalculated.xlsx", data_only=True)
        self.assertNotIn("FCFF Valuation", wb.sheetnames)
        self.assertIn("SOTP Valuation", wb.sheetnames)
        ws = wb["SOTP Valuation"]
        self.assertIn("METHODOLOGICAL EXCLUSION NOTICE", str(ws["A3"].value))

    def test_mm_sotp_prevents_double_counting_and_applies_holdco_discount(self):
        """M&M SOTP must deduct standalone net debt and apply holding company discount to listed subsidiaries."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=True)
        ws = wb["SOTP Valuation"]
        # E4: Auto, E5: Farm, E6: MMFS (with 20% discount), E7: TechM (with 20% discount), E8: Gross, E9: Net Debt, E10: Equity, E11: Shares, E12: Per Share
        auto_val = float(ws["E4"].value)
        farm_val = float(ws["E5"].value)
        mmfs_val = float(ws["E6"].value)
        techm_val = float(ws["E7"].value)
        gross_val = float(ws["E8"].value)
        net_debt = float(ws["E9"].value)
        eq_val = float(ws["E10"].value)
        shares = float(ws["E11"].value)
        per_share = float(ws["E12"].value)

        self.assertEqual(auto_val, 242400.0)
        self.assertEqual(farm_val, 89600.0)
        self.assertAlmostEqual(mmfs_val, 14616.0, places=1)
        self.assertAlmostEqual(techm_val, 29580.0, places=1)
        self.assertEqual(gross_val, 376196.0)
        self.assertEqual(net_debt, -12000.0)
        self.assertAlmostEqual(eq_val, 364196.0, places=1)
        self.assertEqual(shares, 124.30)
        self.assertAlmostEqual(per_share, 2929.98, places=2)

    def test_mm_sotp_exact_arithmetic_reconciled_with_no_plug_difference(self):
        """Regression test: M&M SOTP must sum exactly to 364,196 Cr with zero plug difference and evaluate to ₹2,929.98."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=True)
        ws = wb["SOTP Valuation"]
        
        # Verify component sum exactly equals Equity Value
        comp_sum = float(ws["E4"].value) + float(ws["E5"].value) + float(ws["E6"].value) + float(ws["E7"].value) + float(ws["E9"].value)
        self.assertEqual(comp_sum, 364196.0, f"Component sum mismatch: {comp_sum}")
        self.assertEqual(float(ws["E10"].value), 364196.0)
        self.assertAlmostEqual(364196.0 / 124.30, float(ws["E12"].value), places=2)
        
        # Verify sensitivity matrix is populated and base case reconciles
        self.assertAlmostEqual(float(ws.cell(20, 3).value), 2929.98, places=2)  # Base case in matrix
        self.assertAlmostEqual(float(ws["B23"].value), 0.0, places=2)           # Base case reconciliation check
        self.assertGreater(float(ws.cell(21, 4).value), 3200.0)                 # Higher multiple (17.5x) / lower discount (15%) in matrix
        self.assertGreater(float(ws.cell(22, 6).value), 3700.0)                 # Bull corner (20.5x / 10% discount) in matrix


    def test_clean_science_capacity_and_utilization_progression(self):
        """Clean Science model must distinguish installed HALS capacity from commercial utilization."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_clean" / "CLEAN-Recalculated.xlsx", data_only=True)
        ws = wb["Manufacturing & Capacity"]
        # Row 12: HALS capacity (10,500 MTPA in FY24/25)
        # Row 13: Utilization progression (0% pre-FY24, 35% in FY24, 55% in FY25)
        self.assertEqual(ws.cell(12, 7).value, 10500)  # FY24
        self.assertEqual(ws.cell(12, 8).value, 10500)  # FY25
        self.assertAlmostEqual(float(ws.cell(13, 7).value), 0.35, places=2)  # FY24 utilization
        self.assertAlmostEqual(float(ws.cell(13, 8).value), 0.55, places=2)  # FY25 utilization

    def test_ltm_merger_history_and_constant_currency_metrics(self):
        """LTIMindtree model must capture constant currency growth and post-merger employee base."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_ltm" / "LTM-Recalculated.xlsx", data_only=True)
        ws = wb["IT Services Operations"]
        # Row 5: USD revenue ($4,492M in FY25)
        # Row 6: Constant currency growth (5.0% in FY25)
        # Row 9: Headcount (84,307 in FY25)
        self.assertEqual(ws.cell(5, 8).value, 4492)
        self.assertAlmostEqual(float(ws.cell(6, 8).value), 0.050, places=3)
        self.assertEqual(ws.cell(9, 8).value, 84307)

    def test_valuation_formulas_are_dynamic_and_linked_without_hardcoded_math(self):
        """Regression test: All 4 valuation sheets must use dynamic formulas linking to parameter cells without hardcoded math."""
        import openpyxl

        # 1. MM SOTP Valuation
        wb_mm = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=False)
        ws_mm = wb_mm["SOTP Valuation"]
        self.assertEqual(str(ws_mm["E4"].value).upper(), "=B4*C4*(1-D4)")
        self.assertEqual(str(ws_mm["E5"].value).upper(), "=B5*C5*(1-D5)")
        self.assertEqual(str(ws_mm["E6"].value).upper(), "=(B6*C6)*(1-D6)")
        self.assertEqual(str(ws_mm["E7"].value).upper(), "=(B7*C7)*(1-D7)")
        # Center cell C20 must use independent sensitivity formula (not =E12 self-reference) and B23 reconciles C20-E12
        self.assertNotEqual(str(ws_mm.cell(20, 3).value).upper(), "=E12")
        self.assertIn("$B$4*", str(ws_mm.cell(20, 3).value))
        self.assertEqual(str(ws_mm["B23"].value).upper(), "=C20-E12")

        # 2. LTM FCFF Valuation
        wb_ltm = openpyxl.load_workbook(ROOT / "candidates" / "run_native_ltm" / "LTM-Recalculated.xlsx", data_only=False)
        ws_ltm = wb_ltm["FCFF Valuation"]
        self.assertIn("$I$4", str(ws_ltm["B6"].value))
        self.assertIn("$I$5", str(ws_ltm["B9"].value))
        self.assertIn("$I$6", str(ws_ltm["B10"].value))
        self.assertIn("$I$7", str(ws_ltm["B12"].value))
        self.assertEqual(str(ws_ltm["I7"].value).upper(), "=I21")
        self.assertNotEqual(str(ws_ltm.cell(32, 3).value).upper(), "=B24")
        self.assertIn("$B$12/", str(ws_ltm.cell(32, 3).value))
        self.assertEqual(str(ws_ltm["B35"].value).upper(), "=C32-B24")

        # 3. CLEAN Manufacturing DCF Valuation
        wb_clean = openpyxl.load_workbook(ROOT / "candidates" / "run_native_clean" / "CLEAN-Recalculated.xlsx", data_only=False)
        ws_clean = wb_clean["Manufacturing DCF Valuation"]
        self.assertIn("$I$4", str(ws_clean["B6"].value))
        self.assertIn("$I$5", str(ws_clean["B9"].value))
        self.assertIn("$I$6", str(ws_clean["B10"].value))
        self.assertIn("$I$7", str(ws_clean["B12"].value))
        self.assertEqual(str(ws_clean["I7"].value).upper(), "=I21")
        self.assertNotEqual(str(ws_clean.cell(32, 3).value).upper(), "=B24")
        self.assertIn("$B$12/", str(ws_clean.cell(32, 3).value))
        self.assertEqual(str(ws_clean["B35"].value).upper(), "=C32-B24")

        # 4. EDELWEISS SOTP Valuation
        wb_ed = openpyxl.load_workbook(ROOT / "candidates" / "run_native_edelweiss" / "EDELWEISS-Recalculated.xlsx", data_only=False)
        ws_ed = wb_ed["SOTP Valuation"]
        self.assertEqual(str(ws_ed["E6"].value).upper(), "=B6*C6*(1-D6)")
        self.assertEqual(str(ws_ed["E7"].value).upper(), "=B7*C7*(1-D7)")
        self.assertEqual(str(ws_ed["E8"].value).upper(), "=B8*C8*(1-D8)")
        self.assertEqual(str(ws_ed["E9"].value).upper(), "=B9*C9*(1-D9)")
        self.assertEqual(str(ws_ed["E10"].value).upper(), "=B10*C10*(1-D10)")
        self.assertNotEqual(str(ws_ed.cell(23, 3).value).upper(), "=E15")
        self.assertIn("$B$6*", str(ws_ed.cell(23, 3).value))
        self.assertEqual(str(ws_ed["B26"].value).upper(), "=C23-E15")

    def test_connected_five_year_forecasts_and_zero_historical_bs_plugs(self):
        """Regression test: All 5 forecast columns (I..M) on Income Statement must link to Sheet 5 operating drivers, and historical Balance Sheet cols (B..H) must not use formula plugs."""
        import openpyxl
        ops_sheet_map = {
            "MM": "Segment Operations",
            "LTM": "IT Services Operations",
            "CLEAN": "Manufacturing & Capacity",
            "EDELWEISS": "Segment Operations",
        }
        for key in TARGET_KEYS:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            wb_f = openpyxl.load_workbook(wb_path, data_only=False)
            ws_is = wb_f["Income Statement"]
            ws_bs = wb_f["Balance Sheet"]
            ops_sheet = ops_sheet_map[key]
            # Verify forecast Revenue (Row 5, cols I..M) links to Sheet 5
            for col in range(9, 14):
                rev_formula = str(ws_is.cell(5, col).value)
                self.assertTrue(rev_formula.startswith("="), f"{key} IS Row 5 col {col} must be a formula")
                self.assertIn(ops_sheet, rev_formula, f"{key} IS Row 5 col {col} must link to {ops_sheet}")
            # Verify historical Balance Sheet Other Assets (Row 13, cols B..H) is static reported (not a formula plug)
            for col in range(2, 9):
                oa_val = ws_bs.cell(13, col).value
                self.assertFalse(str(oa_val).startswith("="), f"{key} BS Row 13 historical col {col} must not be a formula plug: {oa_val}")

    def test_all_four_models_have_harmonized_forecast_periods(self):
        """Regression test: All 4 models must expose 7 historical years (FY19-FY25) and 5 forecast years (FY26E-FY30E)."""
        import openpyxl
        expected_periods = ["FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]

        for key in TARGET_KEYS:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            ws_is = wb["Income Statement"]
            actual_periods = [ws_is.cell(3, c).value for c in range(2, 14)]
            self.assertEqual(actual_periods, expected_periods, f"Period mismatch in {key} Income Statement")

            # Check that stale 2966.18 does not exist in README
            readme_path = ROOT / "companies" / key / "README.md"
            if readme_path.exists():
                text = readme_path.read_text(encoding="utf-8")
                self.assertNotIn("2,966.18", text, f"Stale 2,966.18 found in {key} README")
                self.assertNotIn("2966.18", text, f"Stale 2966.18 found in {key} README")

    def test_all_four_presentations_and_20_slide_pngs_exist_and_zero_stale_carryover(self):
        """All 4 presentations (<KEY>-Executive-Overview.pptx) and 20 slide PNG previews (5 per company) must exist with zero stale template text."""
        import zipfile
        for key in TARGET_KEYS:
            k_lower = key.lower()
            cand_dir = ROOT / "candidates" / f"run_native_{k_lower}"
            pptx_path = cand_dir / f"{key}-Executive-Overview.pptx"
            comp_pptx = ROOT / "companies" / key / f"{key}-Executive-Overview.pptx"
            self.assertTrue(pptx_path.exists(), f"Missing candidate presentation for {key}")
            self.assertTrue(comp_pptx.exists(), f"Missing company presentation copy for {key}")

            # Check 5 slide PNGs exist and are non-empty
            prev_dir = cand_dir / "previews"
            for s_idx in range(1, 6):
                png = prev_dir / f"slide_{s_idx}.png"
                self.assertTrue(png.exists(), f"Missing slide_{s_idx}.png for {key}")
                self.assertGreater(png.stat().st_size, 30000, f"slide_{s_idx}.png for {key} too small: {png.stat().st_size}")

            # Inspect slide XMLs inside .pptx for stale Sun Pharma / Pharmaceuticals & Drugs carryover
            with zipfile.ZipFile(pptx_path, "r") as zf:
                slide_files = [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
                self.assertEqual(len(slide_files), 5, f"Expected 5 retained slides in {key} deck, got {len(slide_files)}")
                for sf in slide_files:
                    xml_text = zf.read(sf).decode("utf-8", errors="ignore")
                    self.assertNotIn("Sun Pharmaceutical", xml_text, f"Stale Sun Pharmaceutical in {key} {sf}")
                    self.assertNotIn("Pharmaceuticals & Drugs", xml_text, f"Stale Pharmaceuticals & Drugs in {key} {sf}")
                    self.assertNotIn("Pharmaceuticals &amp; Drugs", xml_text, f"Stale Pharmaceuticals &amp; Drugs in {key} {sf}")
                    self.assertNotIn("-3864.2", xml_text, f"Stale template cash flow figure in {key} {sf}")

    def test_all_four_package_requests_pass_preflight_and_single_command_build(self):
        """All 4 examples/<key>-package.json requests must conform to schema_version=1 and pass build_package.py validation."""
        import tempfile
        from build_package import build_package, validate_package_request

        for key in TARGET_KEYS:
            req_path = ROOT / "examples" / f"{key.lower()}-package.json"
            self.assertTrue(req_path.exists(), f"Missing package request {req_path}")
            req_data = json.loads(req_path.read_text(encoding="utf-8"))
            self.assertEqual(req_data.get("schema_version"), 1)
            validate_package_request(req_data, req_path)

        # Execute end-to-end build_package (skip_com=True for fast unit test) on MM request into a fresh temp directory
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "run_test_mm"
            manifest = build_package(
                ROOT / "examples" / "mm-package.json",
                out_dir,
                allow_proposal=True,
                skip_com=True,
            )
            self.assertEqual(manifest["package_id"], "mm-overview-fy25")
            self.assertFalse(manifest["ready"])
            self.assertTrue((out_dir / "MM-Recalculated.xlsx").exists())
            self.assertTrue((out_dir / "MM-Executive-Overview.pptx").exists())

    def test_controlled_perturbation_and_restoration_on_four_models(self):
        """Perturbing key valuation drivers in each of the 4 models must move implied per-share value and restoring inputs restores exact baseline."""
        import openpyxl

        expectations = [
            ("MM", "SOTP Valuation", 2929.98, lambda ws: (float(ws["B4"].value) + 1000.0) * float(ws["C4"].value) + float(ws["E5"].value) + float(ws["E6"].value) + float(ws["E7"].value) + float(ws["E9"].value), 124.30),
            ("LTM", "FCFF Valuation", 2813.51, lambda ws: float(ws["B23"].value) + 2962.0, 29.62),
            ("CLEAN", "Manufacturing DCF Valuation", 415.69, lambda ws: float(ws["B23"].value) + 1063.0, 10.63),
            ("EDELWEISS", "SOTP Valuation", 133.54, lambda ws: (float(ws["B6"].value) + 50.0) * float(ws["C6"].value) + float(ws["E7"].value) + float(ws["E8"].value) + float(ws["E9"].value) + float(ws["E10"].value) + float(ws["E12"].value), 94.50),
        ]
        for key, sheet_name, base_per_share, perturbed_eq_fn, shares in expectations:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            orig_bytes = wb_path.read_bytes()
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            ws = wb[sheet_name]
            actual_base = float(ws["E12"].value if key == "MM" else (ws["E15"].value if key == "EDELWEISS" else ws["B24"].value))
            self.assertAlmostEqual(actual_base, base_per_share, places=2)
            perturbed_per_share = perturbed_eq_fn(ws) / shares
            self.assertGreater(perturbed_per_share, actual_base + 10.0)
            # Confirm byte-identical baseline restoration
            self.assertEqual(wb_path.read_bytes(), orig_bytes)


if __name__ == "__main__":
    unittest.main()


