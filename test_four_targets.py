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
        # E4: Auto, E5: Farm, E6: MMFS (with 20% discount), E7: TechM (with 20% discount), E10: Net Debt deduction
        auto_val = float(ws["E4"].value)
        farm_val = float(ws["E5"].value)
        mmfs_val = float(ws["E6"].value)
        techm_val = float(ws["E7"].value)
        net_debt = float(ws["E10"].value)
        eq_val = float(ws["E11"].value)

        self.assertEqual(auto_val, 242400.0)
        self.assertEqual(farm_val, 89600.0)
        self.assertAlmostEqual(mmfs_val, 14616.0, places=1)
        self.assertAlmostEqual(techm_val, 29580.0, places=1)
        self.assertEqual(net_debt, -12000.0)
        self.assertAlmostEqual(eq_val, 368696.0, places=1)

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


if __name__ == "__main__":
    unittest.main()
