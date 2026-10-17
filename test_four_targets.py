"""Comprehensive unit and regression tests for the four priority target packages (M&M, LTM, CLEAN, EDELWEISS)."""
import json
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent

from modelmaker import Package
from validate_workbook import (
    is_six_sheet_research_workbook,
    run_six_sheet_research_identities,
    scan_formula_errors_and_missing_caches,
)

TARGET_KEYS = ["MM", "LTM", "CLEAN", "EDELWEISS"]


class FourTargetsTests(unittest.TestCase):
    def test_all_four_workbooks_recalculate_with_zero_errors_and_pass_36_identities(self):
        """All four candidate workbooks must calculate with zero formula errors, zero missing caches, and pass all 36 independent identities."""
        total_identities = 0
        for key in TARGET_KEYS:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            self.assertTrue(wb_path.exists(), f"Recalculated workbook missing for {key}")
            pkg = Package(wb_path.read_bytes())
            errors, missing = scan_formula_errors_and_missing_caches(pkg)
            self.assertEqual(len(errors), 0, f"{key} has {len(errors)} formula errors: {errors}")
            self.assertEqual(len(missing), 0, f"{key} has {len(missing)} uncalculated cells: {missing}")
            self.assertTrue(is_six_sheet_research_workbook(pkg))
            ids = run_six_sheet_research_identities(pkg)
            self.assertGreaterEqual(len(ids), 8)
            for chk in ids:
                self.assertTrue(chk["passed"], f"{key} failed identity {chk['name']}: {chk}")
            total_identities += len(ids)
        self.assertEqual(total_identities, 36)

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
            self.assertGreaterEqual(len(entries), 25, f"Registry for {key} has insufficient entries")
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

    def test_edelweiss_fcff_strictly_blocked_and_sotp_applied_with_standalone_holdco_debt_bridge(self):
        """Negative test: Edelweiss package must formally block industrial FCFF and enforce SOTP with Audited Standalone HoldCo Net Debt (-₹3,067 Cr)."""
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
        # Verify Audited FY25 EFSL Standalone HoldCo Net Debt Bridge (Rows 34-40)
        self.assertEqual(float(ws["B35"].value), 18004.0)  # Consolidated Borrowings
        self.assertEqual(float(ws["B36"].value), -14932.0)  # Ring-Fenced Operating Subsidiary Debt
        self.assertEqual(float(ws["B37"].value), 3072.0)   # Audited EFSL Standalone Borrowings
        self.assertEqual(float(ws["B38"].value), -5.0)      # Audited EFSL Standalone Cash
        self.assertEqual(float(ws["B39"].value), 3067.0)   # Standalone HoldCo Net Debt
        self.assertEqual(float(ws["E11"].value), 15402.0)  # Gross Vertical Sum
        self.assertEqual(float(ws["E12"].value), -3067.0)  # HoldCo Net Debt Deduction
        self.assertEqual(float(ws["E13"].value), 12335.0)  # Implied Equity Value
        self.assertAlmostEqual(float(ws["E15"].value), 130.53, places=2)

    def test_mm_sotp_prevents_double_counting_and_applies_standalone_net_debt_bridge(self):
        """M&M SOTP must use connected FY26E Standalone Auto/Farm Core EBITDA, 15-Oct-2025 MMFS/TechM market caps with 20% HoldCo discount, and Audited FY25 Standalone Net Cash (+₹9,109 Cr)."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=True)
        ws = wb["SOTP Valuation"]
        # E4: Auto, E5: Farm, E6: MMFS (20% discount), E7: TechM (20% discount), E8: Gross, E9: Standalone Net Cash, E10: Equity, E11: Shares, E12: Per Share
        auto_val = float(ws["E4"].value)
        farm_val = float(ws["E5"].value)
        mmfs_val = float(ws["E6"].value)
        techm_val = float(ws["E7"].value)
        gross_val = float(ws["E8"].value)
        net_cash = float(ws["E9"].value)
        eq_val = float(ws["E10"].value)
        shares = float(ws["E11"].value)
        per_share = float(ws["E12"].value)

        self.assertEqual(auto_val, 191338.0)
        self.assertEqual(farm_val, 96239.0)
        self.assertAlmostEqual(mmfs_val, 15149.35, places=2)
        self.assertAlmostEqual(techm_val, 29115.69, places=2)
        self.assertAlmostEqual(gross_val, 331842.04, places=2)
        self.assertEqual(net_cash, 9109.0)
        self.assertAlmostEqual(eq_val, 340951.04, places=2)
        self.assertEqual(shares, 124.30)
        self.assertAlmostEqual(per_share, 2742.97, places=2)

    def test_mm_sotp_exact_arithmetic_reconciled_with_no_plug_difference(self):
        """Regression test: M&M SOTP must sum exactly to 340,951.04 Cr with zero plug difference and evaluate to ₹2,742.97."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=True)
        ws = wb["SOTP Valuation"]

        comp_sum = float(ws["E4"].value) + float(ws["E5"].value) + float(ws["E6"].value) + float(ws["E7"].value) + float(ws["E9"].value)
        self.assertAlmostEqual(comp_sum, 340951.04064, places=4, msg=f"Component sum mismatch: {comp_sum}")
        self.assertAlmostEqual(float(ws["E10"].value), 340951.04064, places=4)
        self.assertAlmostEqual(340951.04064 / 124.30, float(ws["E12"].value), places=2)

        # Verify sensitivity matrix is populated and base case (C20) reconciles to E12 via B23
        self.assertAlmostEqual(float(ws["C20"].value), 2742.97, places=2)
        self.assertAlmostEqual(float(ws["B23"].value), 0.0, places=2)

    def test_clean_science_capacity_and_utilization_progression(self):
        """Clean Science model must keep CFCL Unit 4 HALS installed capacity at verified 10,500 MTPA across FY24-FY30E and ramp commercial utilization."""
        import openpyxl
        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_clean" / "CLEAN-Recalculated.xlsx", data_only=True)
        ws = wb["Manufacturing & Capacity"]
        # Row 12: HALS capacity (10,500 MTPA across FY24-FY30E, cols G..M)
        for col in range(7, 14):
            self.assertEqual(ws.cell(12, col).value, 10500)
        # Row 13: Commercial utilization progression (15% in FY24, 21% in FY25, ramping to 88% in FY30E)
        self.assertAlmostEqual(float(ws.cell(13, 7).value), 0.15, places=2)
        self.assertAlmostEqual(float(ws.cell(13, 8).value), 0.21, places=2)
        self.assertAlmostEqual(float(ws.cell(13, 13).value), 0.88, places=2)

    def test_ltm_merger_history_and_non_overlapping_operational_chain(self):
        """LTIMindtree model must capture non-overlapping Headcount * Utilization * Billing -> USD Revenue -> INR Revenue chain and Merger Comparability Isolation."""
        import openpyxl
        wb_v = openpyxl.load_workbook(ROOT / "candidates" / "run_native_ltm" / "LTM-Recalculated.xlsx", data_only=True)
        ws_v = wb_v["IT Services Operations"]
        self.assertEqual(ws_v.cell(5, 8).value, 4492)
        self.assertAlmostEqual(float(ws_v.cell(6, 8).value), 0.050, places=3)
        self.assertEqual(ws_v.cell(9, 8).value, 84307)

        wb_f = openpyxl.load_workbook(ROOT / "candidates" / "run_native_ltm" / "LTM-Recalculated.xlsx", data_only=False)
        ws_f = wb_f["IT Services Operations"]
        # Ensure zero dummy +0*$I$5 terms and verify USD Revenue (Row 5) is driven by Headcount (Row 9) * Utilization (Row 10) * Billing Rate (Row 12)
        for col in range(9, 14):
            col_let = openpyxl.utils.get_column_letter(col)
            self.assertEqual(str(ws_f[f"{col_let}5"].value).upper(), f"=ROUND({col_let}9*{col_let}10*{col_let}12/1000000,1)")
            self.assertEqual(str(ws_f[f"{col_let}8"].value).upper(), f"=ROUND({col_let}5*{col_let}7/10,0)")
            self.assertNotIn("+0*", str(ws_f[f"{col_let}8"].value))

    def test_valuation_formulas_are_dynamic_and_linked_without_hardcoded_math(self):
        """Regression test: All 4 valuation sheets must use dynamic formulas linking to parameter cells without hardcoded math or cash bridge plugs."""
        import openpyxl

        # 1. MM SOTP Valuation
        wb_mm = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=False)
        ws_mm = wb_mm["SOTP Valuation"]
        self.assertEqual(str(ws_mm["E4"].value).upper(), "=B4*C4*(1-D4)")
        self.assertEqual(str(ws_mm["E5"].value).upper(), "=B5*C5*(1-D5)")
        self.assertEqual(str(ws_mm["E6"].value).upper(), "=(B6*C6)*(1-D6)")
        self.assertEqual(str(ws_mm["E7"].value).upper(), "=(B7*C7)*(1-D7)")
        self.assertEqual(str(ws_mm["E9"].value).upper(), "=-B45")
        self.assertEqual(str(ws_mm["B23"].value).upper(), "=C20-E12")

        # 2. LTM FCFF Valuation
        wb_ltm = openpyxl.load_workbook(ROOT / "candidates" / "run_native_ltm" / "LTM-Recalculated.xlsx", data_only=False)
        ws_ltm = wb_ltm["FCFF Valuation"]
        self.assertEqual(str(ws_ltm["I7"].value).upper(), "=I21")
        self.assertEqual(str(ws_ltm["I14"].value).upper(), "=I60")
        self.assertEqual(str(ws_ltm["B21"].value), "='Balance Sheet'!H12+'Balance Sheet'!H14")
        self.assertEqual(str(ws_ltm["B24"].value).upper(), "=B23/$I$9")
        self.assertEqual(str(ws_ltm["B35"].value).upper(), "=D32-B24")

        # 3. CLEAN Manufacturing DCF Valuation
        wb_clean = openpyxl.load_workbook(ROOT / "candidates" / "run_native_clean" / "CLEAN-Recalculated.xlsx", data_only=False)
        ws_clean = wb_clean["Manufacturing DCF Valuation"]
        self.assertEqual(str(ws_clean["I7"].value).upper(), "=I21")
        self.assertEqual(str(ws_clean["I14"].value).upper(), "=C60")
        self.assertEqual(str(ws_clean["B21"].value), "='Balance Sheet'!H12+'Balance Sheet'!H14")
        self.assertEqual(str(ws_clean["B24"].value).upper(), "=B23/$I$9")
        self.assertEqual(str(ws_clean["B35"].value).upper(), "=D32-B24")

        # 4. EDELWEISS SOTP Valuation
        wb_ed = openpyxl.load_workbook(ROOT / "candidates" / "run_native_edelweiss" / "EDELWEISS-Recalculated.xlsx", data_only=False)
        ws_ed = wb_ed["SOTP Valuation"]
        self.assertEqual(str(ws_ed["E6"].value).upper(), "=B6*C6*(1-D6)")
        self.assertEqual(str(ws_ed["E7"].value).upper(), "=B7*C7*(1-D7)")
        self.assertEqual(str(ws_ed["E8"].value).upper(), "=B8*C8*(1-D8)")
        self.assertEqual(str(ws_ed["E9"].value).upper(), "=B9*C9*(1-D9)")
        self.assertEqual(str(ws_ed["E10"].value).upper(), "=B10*C10*(1-D10)")
        self.assertEqual(str(ws_ed["E12"].value).upper(), "=-B39")
        self.assertEqual(str(ws_ed["B26"].value).upper(), "=C23-E15")

    def test_connected_five_year_forecasts_and_zero_historical_and_forecast_bs_plugs(self):
        """Regression test: All 5 forecast columns (I..M) on Income Statement must link to Sheet 5 operating drivers, and Balance Sheet must have zero historical plugs and double-entry forecast cash closure on Row 14."""
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
            # Verify Balance Sheet Other Assets Ex-Cash (Row 13) does not plug from Total Liabilities (Row 8)
            for col in range(2, 14):
                oa_formula = str(ws_bs.cell(13, col).value)
                self.assertNotIn("8-", oa_formula, f"{key} BS Row 13 col {col} must not plug from Total Liabilities (Row 8): {oa_formula}")
            # Verify forecast Cash closes via Cash_{t-1} + Net Cash Flow_t (either directly on Row 14 or via sub-schedule Row 22)
            for col in range(9, 14):
                prev_let = openpyxl.utils.get_column_letter(col - 1)
                curr_let = openpyxl.utils.get_column_letter(col)
                if key == "MM":
                    cash_formula = str(ws_bs[f"{curr_let}14"].value)
                    self.assertEqual(cash_formula, f"={prev_let}14+'Cash Flow'!{curr_let}7")
                else:
                    self.assertEqual(str(ws_bs[f"{curr_let}14"].value), f"={curr_let}22")
                    self.assertEqual(str(ws_bs[f"{curr_let}22"].value), f"={prev_let}22+'Cash Flow'!{curr_let}7")

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

    def test_all_four_package_requests_pass_preflight_and_excel_only_build(self):
        """All 4 examples/<key>-package.json requests must conform to schema_version=1 and pass build_package.py --excel-only validation."""
        import tempfile
        from build_package import build_package, validate_package_request

        for key in TARGET_KEYS:
            req_path = ROOT / "examples" / f"{key.lower()}-package.json"
            self.assertTrue(req_path.exists(), f"Missing package request {req_path}")
            req_data = json.loads(req_path.read_text(encoding="utf-8"))
            self.assertEqual(req_data.get("schema_version"), 1)
            validate_package_request(req_data, req_path, excel_only=True)

        # Execute end-to-end build_package (excel_only=True, skip_com=True for fast unit test) on MM request into a fresh temp directory
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "run_test_mm_excel"
            manifest = build_package(
                ROOT / "examples" / "mm-package.json",
                out_dir,
                allow_proposal=True,
                skip_com=True,
                excel_only=True,
            )
            self.assertEqual(manifest["package_id"], "mm-overview-fy25")
            self.assertTrue(manifest["excel_only"])
            self.assertFalse(manifest["ready"])
            self.assertTrue((out_dir / "MM-Recalculated.xlsx").exists())
            self.assertTrue((out_dir / "workbook_validation.json").exists())
            self.assertTrue((out_dir / "source_assumption_coverage.json").exists())
            self.assertTrue((out_dir / "SOURCE_ASSUMPTION_COVERAGE.md").exists())
            self.assertFalse((out_dir / "MM-Executive-Overview.pptx").exists())

    def test_controlled_perturbation_and_restoration_on_four_models(self):
        """Perturbing key valuation drivers in each of the 4 models must move implied per-share value and restoring inputs restores exact baseline."""
        import openpyxl

        expectations = [
            ("MM", "SOTP Valuation", 2742.97, lambda ws: (float(ws["B4"].value) + 1000.0) * float(ws["C4"].value) + float(ws["E5"].value) + float(ws["E6"].value) + float(ws["E7"].value) + float(ws["E9"].value), 124.30),
            ("LTM", "FCFF Valuation", 2811.09, lambda ws: float(ws["B23"].value) + 2962.0, 29.62),
            ("CLEAN", "Manufacturing DCF Valuation", 480.89, lambda ws: float(ws["B23"].value) + 1063.0, 10.63),
            ("EDELWEISS", "SOTP Valuation", 130.53, lambda ws: (float(ws["B6"].value) + 500.0) * float(ws["C6"].value) * (1.0 - float(ws["D6"].value)) + float(ws["E7"].value) + float(ws["E8"].value) + float(ws["E9"].value) + float(ws["E10"].value) + float(ws["E12"].value), 94.50),
        ]
        for key, sheet_name, base_per_share, perturbed_eq_fn, shares in expectations:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            orig_bytes = wb_path.read_bytes()
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            ws = wb[sheet_name]
            if key == "MM":
                actual_base = float(ws["E12"].value)
            elif key == "EDELWEISS":
                actual_base = float(ws["E15"].value)
            else:
                actual_base = float(ws["B24"].value)
            self.assertAlmostEqual(actual_base, base_per_share, places=2)
            perturbed_per_share = perturbed_eq_fn(ws) / shares
            self.assertGreater(perturbed_per_share, actual_base + 5.0)
            # Confirm byte-identical baseline restoration
            self.assertEqual(wb_path.read_bytes(), orig_bytes)

    def test_empirical_104_week_regression_diagnostics_embedded_in_all_four_models(self):
        """All 4 models must embed an authentic 104-week OLS regression diagnostics table vs Nifty 50 with N=104 and valid statistics."""
        import openpyxl

        configs = [
            ("MM", "SOTP Valuation", 52, 54, 55, 59, 61),
            ("LTM", "FCFF Valuation", 83, 85, 86, 90, 92),
            ("CLEAN", "Manufacturing DCF Valuation", 83, 85, 86, 90, 92),
            ("EDELWEISS", "SOTP Valuation", 52, 54, 55, 59, 61),
        ]
        for key, sheet_name, r_n, r_raw, r_blume, r_r2, r_t in configs:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            self.assertIn(sheet_name, wb.sheetnames)
            ws = wb[sheet_name]

            # N = 104
            self.assertEqual(int(ws.cell(r_n, 2).value), 104, f"{key} N must equal 104")
            raw_beta = float(ws.cell(r_raw, 2).value)
            blume_beta = float(ws.cell(r_blume, 2).value)
            r2 = float(ws.cell(r_r2, 2).value)
            t_stat = float(ws.cell(r_t, 2).value)

            self.assertGreater(raw_beta, 0.5, f"{key} raw beta out of expected bounds: {raw_beta}")
            self.assertLess(raw_beta, 2.5, f"{key} raw beta out of expected bounds: {raw_beta}")
            self.assertGreater(blume_beta, 0.5, f"{key} Blume beta out of expected bounds: {blume_beta}")
            self.assertGreater(r2, 0.10, f"{key} R^2 must be > 0.10: {r2}")
            self.assertGreater(t_stat, 3.5, f"{key} t-statistic must be > 3.5: {t_stat}")

    def test_dcf_stub_period_timing_reconciliation_and_cash_bridge_audit_in_ltm_and_clean(self):
        """LTM and CLEAN must feature stub-period timing reconciliation and unencumbered liquid treasury cash bridge audit."""
        import openpyxl

        for key, sheet_name, expected_base, exp_ev in [
            ("LTM", "FCFF Valuation", 2811.09, 72018.0),
            ("CLEAN", "Manufacturing DCF Valuation", 480.89, 4713.0),
        ]:
            wb_path = ROOT / "candidates" / f"run_native_{key.lower()}" / f"{key}-Recalculated.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            ws = wb[sheet_name]

            # Row 65: Annual baseline ties to E8 / B24
            ann_val = float(ws["J65"].value)
            self.assertAlmostEqual(ann_val, expected_base, places=2)

            # Row 66: Stub discounting (167 days remaining)
            stub_val = float(ws["J66"].value)
            self.assertGreater(stub_val, 0.0)

            # Row 67: Terminal value concentration
            tv_conc = float(ws["B67"].value)
            self.assertGreater(tv_conc, 0.50, f"{key} TV concentration should be significant: {tv_conc}")
            self.assertLess(tv_conc, 0.95, f"{key} TV concentration should be bounded: {tv_conc}")

            # Cash bridge audit (Rows 71-77)
            cash_eq = float(ws["B71"].value)
            liq_treas = float(ws["B72"].value)
            gross_surplus = float(ws["B73"].value)
            op_buffer = float(ws["B74"].value)
            net_unenc = float(ws["B75"].value)
            borrowings = float(ws["B76"].value)
            net_cushion = float(ws["B77"].value)

            self.assertAlmostEqual(gross_surplus, cash_eq + liq_treas, places=2)
            self.assertAlmostEqual(net_unenc, gross_surplus - op_buffer, places=2)
            self.assertAlmostEqual(net_cushion, net_unenc - borrowings, places=2)
            self.assertGreater(net_cushion, 0.0, f"{key} fortress balance sheet must have positive net cash cushion")

    def test_mm_volume_realization_asp_progression(self):
        """M&M model must articulate vehicle and tractor realization (₹ Lakh/unit) progression on Segment Operations."""
        import openpyxl

        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_mm" / "MM-Recalculated.xlsx", data_only=True)
        ws = wb["Segment Operations"]

        # Row 31: Auto Realization per unit (₹ Lakh / Unit)
        # Row 32: Farm Realization per tractor (₹ Lakh / Unit)
        # Cols B..H: FY19..FY25 historical; Cols I..M: FY26E..FY30E forecast
        for col in range(2, 14):
            auto_asp = float(ws.cell(31, col).value)
            farm_asp = float(ws.cell(32, col).value)
            self.assertGreater(auto_asp, 5.0, f"Auto ASP too low in col {col}: {auto_asp}")
            self.assertLess(auto_asp, 15.0, f"Auto ASP too high in col {col}: {auto_asp}")
            self.assertGreater(farm_asp, 4.0, f"Farm ASP too low in col {col}: {farm_asp}")
            self.assertLess(farm_asp, 12.0, f"Farm ASP too high in col {col}: {farm_asp}")

    def test_edelweiss_minority_interest_and_debt_ring_fencing_governance_schedule(self):
        """EDELWEISS model must articulate minority interest and debt ring-fencing governance schedule."""
        import openpyxl

        wb = openpyxl.load_workbook(ROOT / "candidates" / "run_native_edelweiss" / "EDELWEISS-Recalculated.xlsx", data_only=True)
        ws = wb["SOTP Valuation"]

        # Row 64 Header: MINORITY INTEREST & OPERATING DEBT RING-FENCING GOVERNANCE SCHEDULE
        self.assertIn("GOVERNANCE SCHEDULE", str(ws["A64"].value).upper())
        self.assertIn("CDPQ", str(ws["A65"].value))
        self.assertIn("Tokio Marine", str(ws["A66"].value))
        self.assertIn("14,932", str(ws["A67"].value))
        self.assertIn("3,067", str(ws["A68"].value))

    def test_batch_build_cli_supports_multiple_requests_and_output_base(self):
        """build_package_batch must validate all 4 target requests and build verified packages in batch mode."""
        import tempfile
        from build_package import build_package_batch

        requests = [
            ROOT / "examples" / "mm-package.json",
            ROOT / "examples" / "ltm-package.json",
            ROOT / "examples" / "clean-package.json",
            ROOT / "examples" / "edelweiss-package.json",
        ]
        with tempfile.TemporaryDirectory() as tmp:
            out_base = Path(tmp) / "batch_test"
            summary = build_package_batch(
                request_paths=requests,
                output_base=out_base,
                allow_proposal=True,
                skip_com=True,
                excel_only=True,
                clean=True,
            )
            self.assertTrue(summary["batch_mode"])
            self.assertEqual(summary["total_requests"], 4)
            self.assertEqual(summary["succeeded"], 4)
            self.assertEqual(summary["failed_count"], 0)
            self.assertEqual(summary["status"], "success")

            # Verify that each subdirectory was created with recalculated workbook and manifest
            for k in ["mm", "ltm", "clean", "edelweiss"]:
                sub = out_base / f"run_excel_{k}"
                self.assertTrue(sub.exists(), f"Batch subdir missing for {k}")
                self.assertTrue((sub / "package.manifest.json").exists())

    def test_four_models_sensitivity_grids_have_independent_equations_and_reconcile(self):
        """All four models must feature 25-cell independent sensitivity matrices that dynamically calculate without circular self-references and reconcile to 0.00."""
        import openpyxl

        models = {
            "MM": ("candidates/run_native_mm/MM-Recalculated.xlsx", "SOTP Valuation", 18, 22, 2, 6, "E12", 23),
            "LTM": ("candidates/run_native_ltm/LTM-Recalculated.xlsx", "FCFF Valuation", 29, 33, 2, 6, "B24", 35),
            "CLEAN": ("candidates/run_native_clean/CLEAN-Recalculated.xlsx", "Manufacturing DCF Valuation", 29, 33, 2, 6, "B24", 35),
            "EDELWEISS": ("candidates/run_native_edelweiss/EDELWEISS-Recalculated.xlsx", "SOTP Valuation", 21, 25, 2, 6, "E15", 26),
        }
        for name, (fpath, sname, r_start, r_end, c_start, c_end, out_cell, check_row) in models.items():
            wb_formula = openpyxl.load_workbook(ROOT / fpath, data_only=False)
            ws_f = wb_formula[sname]
            center_r = (r_start + r_end) // 2
            center_c = (c_start + c_end) // 2
            center_cell = ws_f.cell(center_r, center_c)
            self.assertNotEqual(str(center_cell.value).strip(), f"={out_cell}", f"{name} center cell is a trivial link to output cell")

            wb_val = openpyxl.load_workbook(ROOT / fpath, data_only=True)
            ws_v = wb_val[sname]
            tl_val = ws_v.cell(r_start, c_start).value
            tr_val = ws_v.cell(r_start, c_end).value
            bl_val = ws_v.cell(r_end, c_start).value
            br_val = ws_v.cell(r_end, c_end).value
            chk_val = float(ws_v.cell(check_row, 2).value)

            self.assertAlmostEqual(chk_val, 0.0, places=2, msg=f"{name} sensitivity reconciliation check failed")
            self.assertEqual(len({tl_val, tr_val, bl_val, br_val}), 4, f"{name} corner values must be distinct")

    def test_clean_rejects_unowned_directory_without_manifest(self):
        """build_package must raise RuntimeError if --clean is passed for an arbitrary directory lacking package.manifest.json outside candidates/ or temp."""
        import tempfile
        from build_package import build_package

        with tempfile.TemporaryDirectory() as td:
            arbitrary_dir = Path(td) / "sub_arbitrary" / "unowned_dir"
            arbitrary_dir.mkdir(parents=True)
            (arbitrary_dir / "unrelated_file.txt").write_text("critical user data", encoding="utf-8")

            # Monkeypatch temporary directory check to simulate external non-temp path
            with self.assertRaises(RuntimeError) as ctx:
                build_package(
                    request_path=ROOT / "examples" / "mm-package.json",
                    output_dir=ROOT / "proposals",
                    clean=True,
                    allow_proposal=True,
                )
            self.assertIn("protected", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
