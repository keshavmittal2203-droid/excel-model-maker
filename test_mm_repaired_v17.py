"""Regression test suite for single-target Mahindra & Mahindra V17 model repair and packaging."""
from __future__ import annotations

import json
from pathlib import Path
import unittest
import zipfile

import openpyxl

ROOT = Path(__file__).resolve().parent
MM_DIR = ROOT / "companies" / "MM" / "template_v17"
DELIVERABLES_DIR = ROOT / "deliverables" / "V16_user_handoff"
MASTER_WB = ROOT / "originals" / "Marksans Keshav.xlsx"
BETA_SOURCE = ROOT / "work" / "sources" / "weekly_market_returns_104w_v2.json"
ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}


class TestMMRepairedV17(unittest.TestCase):
    def setUp(self):
        self.wb_path = MM_DIR / "MM-FullTemplate.xlsx"
        self.assertTrue(self.wb_path.exists(), f"MM V17 workbook missing: {self.wb_path}")
        self.handoff_wb = DELIVERABLES_DIR / "workbooks" / "MM-FullTemplate.xlsx"
        self.assertTrue(self.handoff_wb.exists(), f"Handoff MM workbook missing: {self.handoff_wb}")

    def test_master_template_unmodified(self):
        """The 16-sheet master template must remain byte-identical to expected SHA-256."""
        import hashlib
        expected_sha = "4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0"
        actual_sha = hashlib.sha256(MASTER_WB.read_bytes()).hexdigest()
        self.assertEqual(actual_sha, expected_sha, "Master template modified!")

    def test_preserves_all_16_sheets_in_order(self):
        """Workbook must retain all 16 sheets from master template in identical order."""
        master = openpyxl.load_workbook(MASTER_WB, read_only=True)
        wb = openpyxl.load_workbook(self.wb_path, read_only=True)
        self.assertEqual(wb.sheetnames, master.sheetnames)
        self.assertEqual(len(wb.sheetnames), 16)
        master.close()
        wb.close()

    def test_zero_formula_errors_across_all_sheets(self):
        """Values workbook must contain zero formula errors (#REF!, #DIV/0!, #VALUE!, etc.)."""
        wb = openpyxl.load_workbook(self.wb_path, data_only=True)
        errors_found = []
        for name in wb.sheetnames:
            ws = wb[name]
            for row in ws.iter_rows(values_only=True):
                for val in row:
                    if isinstance(val, str) and val in ERRORS:
                        errors_found.append((name, val))
        wb.close()
        self.assertEqual(len(errors_found), 0, f"Found formula errors in MM workbook: {errors_found}")

    def test_historical_balance_sheet_ties_and_blocked_forecast(self):
        """BS!C40:M40 must evaluate to 0.00 Cr and N40 must evaluate to BLOCKED."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=True)
        wb_forms = openpyxl.load_workbook(self.wb_path, data_only=False)
        ws_vals = wb_vals["BS"]
        ws_forms = wb_forms["BS"]

        # Verify historical periods C40:M40
        for col_idx in range(3, 14):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            val = ws_vals.cell(40, col_idx).value
            form = ws_forms.cell(40, col_idx).value
            self.assertEqual(form, f"={col_letter}38-{col_letter}16", f"BS!{col_letter}40 formula wrong")
            self.assertIsNotNone(val, f"BS!{col_letter}40 value is None")
            self.assertAlmostEqual(float(val), 0.0, places=2, msg=f"BS!{col_letter}40 did not tie: {val}")

        # Verify FY26E column N
        n40_val = ws_vals.cell(40, 14).value
        n40_form = ws_forms.cell(40, 14).value
        self.assertIn("BLOCKED", str(n40_val).upper())
        self.assertIn("BLOCKED", str(n40_form).upper())

        wb_vals.close()
        wb_forms.close()

    def test_valuation_and_wacc_explicitly_marked_unavailable(self):
        """DCF fair value, cover page fair value, and WACC must be marked 'Unavailable'."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=True)

        cover_val = wb_vals["Model ->"]["C9"].value
        self.assertIn("UNAVAILABLE", str(cover_val).upper())

        dcf_val = wb_vals["DCF 1st way"]["E42"].value
        self.assertIn("UNAVAILABLE", str(dcf_val).upper())

        wacc_val = wb_vals["WACC-DCF"]["J39"].value
        self.assertIn("UNAVAILABLE", str(wacc_val).upper())

        wb_vals.close()

    def test_compdata_circular_loops_eradicated(self):
        """No formula reverse-solving in compdata: K4:K10 must not form circular loops with F19:F25."""
        wb_forms = openpyxl.load_workbook(self.wb_path, data_only=False)
        ws = wb_forms["compdata"]
        # Check K4 to K10 do not reference F19:F25
        for r in (4, 5, 6, 7, 8, 9, 10):
            val = str(ws.cell(r, 11).value or "")
            self.assertFalse(any(f"F{x}" in val for x in (19, 20, 21, 22, 23, 24, 25)),
                             f"compdata K{r} references F19:F25: {val}")
        # Check F19 to F25
        for r in (19, 20, 21, 22, 23, 24, 25):
            val = str(ws.cell(r, 6).value or "")
            self.assertFalse(any(f"K{x}" in val for x in (4, 5, 6, 7, 8, 9, 10)),
                             f"compdata F{r} references K4:K10: {val}")
        wb_forms.close()

    def test_beta_calculation_independent_reproduction(self):
        """Independent OLS beta from 104-week returns must match 1.4746 raw and 1.3164 Blume."""
        self.assertTrue(BETA_SOURCE.exists())
        source = json.loads(BETA_SOURCE.read_text(encoding="utf-8"))
        sec = source["securities"]["M&M.NS"]["records"]
        bm = source["benchmark"]["records"]
        bm_map = {r["date"]: r["adjclose"] for r in bm}
        sec_map = {r["date"]: r["adjclose"] for r in sec}
        dates = sorted(set(sec_map) & set(bm_map))
        self.assertEqual(len(dates), 105)

        stock_returns = [sec_map[b] / sec_map[a] - 1 for a, b in zip(dates, dates[1:])]
        market_returns = [bm_map[b] / bm_map[a] - 1 for a, b in zip(dates, dates[1:])]
        self.assertEqual(len(stock_returns), 104)
        self.assertEqual(len(market_returns), 104)

        mean_x = sum(market_returns) / len(market_returns)
        mean_y = sum(stock_returns) / len(stock_returns)
        cov = sum((x - mean_x) * (y - mean_y) for x, y in zip(market_returns, stock_returns)) / (len(stock_returns) - 1)
        var = sum((x - mean_x) ** 2 for x in market_returns) / (len(market_returns) - 1)
        raw_beta = cov / var
        blume_beta = (2.0 / 3.0) * raw_beta + (1.0 / 3.0) * 1.0

        self.assertAlmostEqual(raw_beta, 1.4746023202, places=4)
        self.assertAlmostEqual(blume_beta, 1.3164015468, places=4)

    def test_native_sparklines_restored_in_ratio_analysis(self):
        """Ratio Analysis worksheet XML must contain <x14:sparklineGroups> with valid ranges."""
        with zipfile.ZipFile(self.wb_path, "r") as z:
            sheet3_xml = z.read("xl/worksheets/sheet3.xml").decode("utf-8")
            self.assertIn("x14:sparklineGroups", sheet3_xml)
            self.assertIn("C6:G6", sheet3_xml)
            self.assertIn("C26:G26", sheet3_xml)

    def test_deliverables_contain_only_mm(self):
        """User handoff deliverables must contain only MM workbook and previews."""
        workbooks = [p for p in (DELIVERABLES_DIR / "workbooks").glob("*.xlsx") if not p.name.startswith("~$")]
        self.assertEqual(len(workbooks), 1)
        self.assertEqual(workbooks[0].name, "MM-FullTemplate.xlsx")

        preview_dirs = [d.name for d in (DELIVERABLES_DIR / "previews").iterdir() if d.is_dir()]
        self.assertEqual(preview_dirs, ["MM"])

        mm_previews = list((DELIVERABLES_DIR / "previews" / "MM").glob("*.pdf"))
        self.assertEqual(len(mm_previews), 15)

    def test_build_report_and_manifest_validity(self):
        """Build report and package manifest in companies/MM/template_v17 must verify cleanly."""
        report_file = MM_DIR / "build_report.json"
        manifest_file = MM_DIR / "package.manifest.json"
        self.assertTrue(report_file.exists())
        self.assertTrue(manifest_file.exists())

        report = json.loads(report_file.read_text(encoding="utf-8"))
        manifest = json.loads(manifest_file.read_text(encoding="utf-8"))

        self.assertEqual(report["target"], "MM")
        self.assertEqual(report["version"], "V17")
        self.assertFalse(report["ready"])
        self.assertEqual(report["status"], "blocked_forecast_and_valuation")
        self.assertEqual(report["engine"]["circular_references"], 0)
        self.assertEqual(report["checks"]["formula_errors"], 0)
        self.assertEqual(report["checks"]["historical_balance_checks_passed"], 11)
        self.assertEqual(report["checks"]["fy26_balance"], "BLOCKED")
        self.assertEqual(report["checks"]["valuation"], "Unavailable")

        self.assertEqual(manifest["target"], "MM")
        self.assertEqual(manifest["version_tag"], "template_v17")
        self.assertFalse(manifest["governance"]["ready"])
        self.assertEqual(manifest["governance"]["human_pack_approval"], "pending")
        self.assertEqual(manifest["recalculation_audit"]["circular_references"], 0)
        self.assertEqual(manifest["recalculation_audit"]["formula_error_count"], 0)

    def test_main_financials_non_current_assets_restored(self):
        """Main Financials row 62 (Other Assets) and 63 (Total Non-Current Assets) must not be None."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=True)
        mf = wb_vals["Main Financials"]
        # Column I is FY25
        other_assets = mf["I62"].value
        non_current_assets = mf["I63"].value
        total_assets = mf["I70"].value
        total_liab = mf["I57"].value
        self.assertIsNotNone(other_assets, "Main Financials I62 (Other Assets) is None")
        self.assertIsNotNone(non_current_assets, "Main Financials I63 (Total Non-Current Assets) is None")
        self.assertEqual(other_assets, 191771, "Main Financials FY25 Other Assets mismatch")
        self.assertEqual(non_current_assets, 276014, "Main Financials FY25 Non-Current Assets mismatch")
        self.assertEqual(total_assets, 276013, "Main Financials FY25 Total Assets mismatch")
        self.assertEqual(total_liab, 276013, "Main Financials FY25 Total Liabilities mismatch")
        wb_vals.close()

    def test_intrinsic_growth_gross_block_and_no_errors(self):
        """Intrinsic Growth must evaluate gross block from reported fixed assets and have 0 formula errors."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=True)
        ig = wb_vals["Intrinsic Growth"]
        # H is FY25
        gross_block_fy25 = ig["H32"].value
        net_non_curr_fy25 = ig["H34"].value
        net_capex_fy25 = ig["H43"].value
        roic_fy25 = ig["H39"].value
        self.assertIsNotNone(gross_block_fy25)
        self.assertGreater(gross_block_fy25, 0, "Gross block must not be zero")
        self.assertEqual(gross_block_fy25, 34636, "Gross block FY25 should match reported net block 34636")
        self.assertEqual(net_non_curr_fy25, 34636, "Net non current assets FY25 should match 34636")
        self.assertGreater(net_capex_fy25, 0, "Net capex should be positive (3735 Cr)")
        self.assertAlmostEqual(roic_fy25, 0.109364, places=3)
        # Check FY26E blocked headers
        self.assertEqual(ig["I4"].value, "FY26E (Blocked)")
        self.assertEqual(ig["I41"].value, "FY26E (Blocked)")
        self.assertEqual(ig["I57"].value, "FY26E (Blocked)")
        wb_vals.close()

    def test_clean_ascii_headers_and_strings(self):
        """Ensure headers do not contain replacement characters \ufffd or raw \u2014 em-dashes."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=False)
        targets = [
            ("Main Financials", "J8"),
            ("BS", "N2"),
            ("Cash Flow", "N4"),
            ("Ratio Analysis", "H4"),
            ("Data Sheet", "K16"),
            ("Comps", "B4"),
        ]
        for sname, coord in targets:
            val = str(wb_vals[sname][coord].value or "")
            self.assertNotIn("\ufffd", val, f"{sname}!{coord} contains replacement character")
            self.assertNotIn("\u2014", val, f"{sname}!{coord} contains em-dash")
        wb_vals.close()


if __name__ == "__main__":
    unittest.main()
