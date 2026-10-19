"""Regression test suite for single-target Mahindra & Mahindra V17 model repair and packaging."""
from __future__ import annotations

import json
from pathlib import Path
import unittest
import zipfile

import openpyxl
from work.build_template_v17_mm import _validate_output_dir

ROOT = Path(__file__).resolve().parent
MM_DIR = ROOT / "companies" / "MM" / "template_v17"
DELIVERABLES_DIR = ROOT / "deliverables" / "MM_V17_user_handoff"
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
        import hashlib
        workbooks = [p for p in (DELIVERABLES_DIR / "workbooks").glob("*.xlsx") if not p.name.startswith("~$")]
        self.assertEqual(len(workbooks), 1)
        self.assertEqual(workbooks[0].name, "MM-FullTemplate.xlsx")
        self.assertEqual(hashlib.sha256(workbooks[0].read_bytes()).hexdigest(), hashlib.sha256(self.wb_path.read_bytes()).hexdigest())

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

    def test_main_financials_asset_subtotal_and_rounding_delta(self):
        """Source totals remain intact and the one-crore component-rounding delta is visible."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=True)
        wb_forms = openpyxl.load_workbook(self.wb_path, data_only=False)
        mf = wb_vals["Main Financials"]
        mf_forms = wb_forms["Main Financials"]
        # Column I is FY25. Components are rounded by the source, so expose the delta.
        other_assets = mf["I62"].value
        asset_subtotal = mf["I63"].value
        rounding_delta = mf["I69"].value
        total_assets = mf["I70"].value
        total_liab = mf["I57"].value
        self.assertIsNotNone(other_assets, "Main Financials I62 (Other Assets) is None")
        self.assertEqual(other_assets, 191771, "Main Financials FY25 Other Assets mismatch")
        self.assertEqual(asset_subtotal, 276014, "Main Financials asset component subtotal mismatch")
        self.assertEqual(total_assets, 276013, "Main Financials FY25 Total Assets mismatch")
        self.assertEqual(total_liab, 276013, "Main Financials FY25 Total Liabilities mismatch")
        self.assertEqual(rounding_delta, -1, "Source rounding difference must remain visible, not plugged")
        self.assertEqual(mf_forms["I69"].value, "=I70-I63")
        self.assertEqual(mf["I72"].value, "TIED")
        self.assertTrue(all(mf[f"I{r}"].value is None for r in (65,66,67,68)))
        wb_vals.close()
        wb_forms.close()

    def test_intrinsic_growth_blocks_unsupported_asset_classification(self):
        """Keep the reported net block but do not calculate ROIC from undisclosed asset splits."""
        wb_vals = openpyxl.load_workbook(self.wb_path, data_only=True)
        wb_forms = openpyxl.load_workbook(self.wb_path, data_only=False)
        ig = wb_vals["Intrinsic Growth"]
        self.assertEqual(ig["H32"].value, 34636, "FY25 reported net block must remain visible")
        self.assertEqual(wb_forms["Intrinsic Growth"]["H32"].value, "=BS!M28")
        for coord in ("H7","H8","H9","H11","H14","H15","H17","H19","H22","H23","H24","H25","H26","H27","H28","H29","H30","H33","H34","H36","H39","H43","H44","H50","H52","H59","H60","H62"):
            self.assertIsNone(ig[coord].value, f"Unsupported split or derived metric must not display zero: {coord}")
        self.assertIn("unavailable", str(ig["B3"].value).lower())
        self.assertIn("not disclosed", str(ig["B6"].value).lower())
        self.assertEqual(ig["H10"].value, 191771)
        self.assertEqual(ig["H16"].value, 69950)
        self.assertIsNotNone(ig["H47"].value)
        # Check FY26E blocked headers
        self.assertEqual(ig["I4"].value, "FY26E (Blocked)")
        self.assertEqual(ig["I41"].value, "FY26E (Blocked)")
        self.assertEqual(ig["I57"].value, "FY26E (Blocked)")
        wb_vals.close()
        wb_forms.close()

    def test_unverified_share_count_and_market_cap_are_not_used(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=True)
        formulas = openpyxl.load_workbook(self.wb_path, data_only=False)
        ds = wb["Data Sheet"]
        self.assertIsNone(ds["B6"].value)
        self.assertIsNone(ds["B9"].value)
        self.assertIn("unavailable",str(ds["A6"].value).lower())
        self.assertIn("unavailable",str(ds["A9"].value).lower())
        comps = wb["Comps"]
        for coord in ("E8","F8","G8","H8"):
            self.assertIsNone(comps[coord].value, f"Unsupported market field {coord} must be blank, not zero/N/A")
        self.assertAlmostEqual(float(comps["P8"].value), 3497.2 / 103.97, places=3)
        self.assertIn("Data Sheet", formulas["Comps"]["P8"].value)
        wb.close()
        formulas.close()

    def test_cash_flow_labels_are_short_and_reported_detail_stays_blank(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=True)
        ws = wb["Cash Flow"]
        self.assertEqual(ws["B5"].value, "Operating cash flow (Screener)")
        self.assertEqual(ws["B14"].value, "Investing cash flow (Screener)")
        self.assertEqual(ws["B23"].value, "Financing cash flow (Screener)")
        source=json.loads((ROOT/"work"/"sources"/"MM_consolidated.json").read_text(encoding="utf-8"))
        table=source["sections"]["cash-flow"]["tables"][0]
        years={int(h[-4:]):i for i,h in enumerate(table["headers"]) if h[-4:].isdigit()}
        for year in range(2015,2026):
            col=year-2012
            idx=years[year]
            for row,key in ((5,"Cash from Operating Activity"),(14,"Cash from Investing Activity"),(23,"Cash from Financing Activity"),(32,"Free Cash Flow")):
                self.assertEqual(ws.cell(row,col).value,table["rows"][key][idx],f"Cash Flow!{ws.cell(row,col).coordinate} does not match Screener source")
            self.assertEqual(ws.cell(31,col).value,sum(table["rows"][key][idx] for key in ("Cash from Operating Activity","Cash from Investing Activity","Cash from Financing Activity")))
        for row in list(range(6,14)) + list(range(15,23)) + list(range(24,31)):
            for col in range(3,14):
                self.assertIsNone(ws.cell(row,col).value, f"Unsupported cash-flow detail {ws.cell(row,col).coordinate} must be blank")
        wb.close()

    def test_comps_empty_peer_rows_have_no_grid_clutter(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=False)
        ws = wb["Comps"]
        self.assertIn("peer ev multiples unavailable", str(ws["B10"].value).lower())
        self.assertIn("$B$3:$P$10", str(ws.print_area))
        for row in list(range(9,10)) + list(range(11,35)):
            for col in range(2,17):
                cell=ws.cell(row,col)
                self.assertIsNone(cell.value, f"Unused Comps cell {cell.coordinate} must be blank")
                self.assertTrue(all(getattr(cell.border,side).style is None for side in ("left","right","top","bottom")), f"Unused Comps cell {cell.coordinate} has border clutter")
        wb.close()

    def test_dcf_has_no_stale_sotp_component_claims(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=False)
        ws = wb["DCF 1st way"]
        self.assertIn("sotp unavailable", str(ws["B47"].value).lower())
        stale_terms = ("core automotive", "core farm", "standalone net cash bridge", "listed stake in mmfs", "listed stake in tech mahindra")
        for row in ws.iter_rows(min_row=47,max_row=61,min_col=2,max_col=13):
            for cell in row:
                val=str(cell.value or "").lower()
                self.assertFalse(any(term in val for term in stale_terms), f"Stale SOTP text remains in {cell.coordinate}: {val}")
        wb.close()

    def test_data_sheet_banner_is_not_an_internal_warning(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=False)
        text = str(wb["Data Sheet"]["E2"].value or "")
        self.assertNotIn("please do not make any changes", text.lower())
        self.assertIn("historical sourced inputs", text.lower())
        wb.close()

    def test_cover_business_description_includes_financial_services(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=False)
        cover = wb["Model ->"]
        self.assertEqual(cover["C5"].value, "Automotive, Farm Equipment & Financial Services")
        self.assertEqual(cover["B11"].value, "Historical close (15-Oct-2025)")
        self.assertEqual(cover["C11"].value, "='Data Sheet'!B8")
        self.assertIn("0.00", cover["C11"].number_format)
        self.assertGreaterEqual(cover.column_dimensions["B"].width, 32)
        wb.close()

    def test_unused_peer_data_and_unverified_rm_series_are_removed(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=False)
        compdata = wb["compdata"]
        self.assertEqual(compdata["C10"].value, "M&M")
        for row in range(18, 26):
            for col in range(2, 14):
                self.assertIsNone(compdata.cell(row, col).value)
        for row in range(17, 26):
            self.assertTrue(compdata.row_dimensions[row].hidden)
        self.assertIn("$B$2:$M$10", str(compdata.print_area))

        rm = wb["RM"]
        self.assertIn("not used in v17", str(rm["C8"].value).lower())
        self.assertIn("no approved-source evidence", str(rm["E8"].value).lower())
        for row in range(9, 36):
            for col in range(2, 10):
                self.assertIsNone(rm.cell(row, col).value, f"Unverified inherited RM data remains in {rm.cell(row,col).coordinate}")
            self.assertTrue(rm.row_dimensions[row].hidden)
        self.assertIn("$B$8:$I$8", str(rm.print_area))
        wb.close()

    def test_business_description_source_registry_is_screener(self):
        entries = json.loads((ROOT / "docs" / "mm_source_registry.json").read_text(encoding="utf-8"))
        entry = next(e for e in entries if e["metric"] == "Business profile summary")
        self.assertEqual(entry["source_org"], "Screener")
        self.assertEqual(entry["workbook_cell"], "'Model ->'!C5")
        self.assertEqual(entry["status"], "profile_sourced")
        self.assertIn("screener.in/company/M%26M", entry["official_url"])
        local = ROOT / entry["local_source_file"]
        self.assertTrue(local.exists())
        import hashlib
        self.assertEqual(hashlib.sha256(local.read_bytes()).hexdigest(), entry["local_source_sha256"])

    def test_source_registry_matches_v17_used_cells_and_unavailable_data(self):
        entries = json.loads((ROOT / "docs" / "mm_source_registry.json").read_text(encoding="utf-8"))
        by_metric = {entry["metric"]: entry for entry in entries}
        capex = by_metric["CapEx / fixed asset purchases (FY15-FY25)"]
        self.assertEqual(capex["status"], "unavailable_on_allowed_sources")
        self.assertIsNone(capex["workbook_cell"])
        shares = by_metric["Shares outstanding at 2025-10-15 cutoff"]
        self.assertEqual(shares["status"], "unavailable_on_allowed_sources")
        self.assertIn("'Data Sheet'!B6", shares["workbook_cell"])
        sales = by_metric["Revenue from Operations (FY17-FY25)"]
        self.assertEqual(sales["workbook_cell"], "'Data Sheet'!B17:J17")
        self.assertEqual(sales["source_org"], "Screener")
        market_cap = by_metric["Market capitalization / equity value"]
        self.assertEqual(market_cap["status"], "unavailable_on_allowed_sources")
        self.assertIsNone(market_cap["reported_value"])
        self.assertFalse(any("bseindia.com" in str(e.get("official_url")) for e in entries))
        text = (ROOT / "docs" / "MM_SOURCE_REGISTRY.md").read_text(encoding="utf-8").lower()
        self.assertIn("capex detail is unavailable", text)
        self.assertNotIn("bse statutory sub-schedule", text)

    def test_comps_uses_only_supported_market_multiples(self):
        wb_values = openpyxl.load_workbook(self.wb_path, data_only=True)
        wb_formulas = openpyxl.load_workbook(self.wb_path, data_only=False)
        comps = wb_values["Comps"]
        for coord in ("E8","F8","G8","H8"):
            self.assertIsNone(comps[coord].value)
        self.assertAlmostEqual(float(comps["P8"].value), 3497.2 / 103.97, places=3)
        self.assertIn("Data Sheet", wb_formulas["Comps"]["P8"].value)
        wb_values.close()
        wb_formulas.close()

    def test_print_preview_does_not_force_every_sheet_to_one_tiny_page(self):
        wb = openpyxl.load_workbook(self.wb_path, data_only=False)
        for name in ("Main Financials","Cash Flow","Data Sheet"):
            ws=wb[name]
            self.assertEqual(ws.page_setup.fitToHeight,0)
            self.assertTrue(ws.page_setup.scale is None or ws.page_setup.scale >= 50, f"{name} preview is scaled too small")
        self.assertEqual(wb["Cash Flow"].column_dimensions["B"].width,42.0)
        self.assertIn(wb["Intrinsic Growth"].page_setup.fitToHeight,(0,None))
        self.assertTrue(wb["Intrinsic Growth"].page_setup.scale is None or wb["Intrinsic Growth"].page_setup.scale>=50)
        self.assertIn("$B$3:$H$40",str(wb["Intrinsic Growth"].print_area))
        self.assertTrue(all(wb["DCF 2nd way"].row_dimensions[r].hidden for r in range(5,101)))
        self.assertIn("$B$3:$N$4",str(wb["DCF 2nd way"].print_area))
        wb.close()

    def test_template_debug_markers_removed_from_visible_sheets(self):
        wb=openpyxl.load_workbook(self.wb_path,data_only=False)
        mf=wb["Main Financials"]
        self.assertFalse(any(mf.cell(r,1).value=="#" for r in range(1,mf.max_row+1)))
        dcf2=wb["DCF 2nd way"]
        self.assertIsNone(dcf2["C4"].value)
        self.assertFalse(any(dcf2.cell(r,1).value=="x" for r in range(1,dcf2.max_row+1)))
        self.assertIn("unavailable",str(dcf2["B4"].value).lower())
        wb.close()

    def test_no_generated_powerpoint_previews_remain_active(self):
        self.assertFalse((ROOT / "companies" / "MM" / "previews").exists())

    def test_build_output_path_is_scoped_to_generated_output_roots(self):
        self.assertEqual(_validate_output_dir(ROOT / "companies" / "MM" / "test-output"), (ROOT / "companies" / "MM" / "test-output").resolve())
        for unsafe in (ROOT, ROOT / "originals" / "unsafe", Path("C:/outside/MM-output")):
            with self.assertRaises(ValueError):
                _validate_output_dir(unsafe)

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
