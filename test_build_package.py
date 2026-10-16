"""Unit and regression tests for build_package.py orchestrator."""
from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from build_package import (
    build_package,
    check_com_capabilities,
    compute_file_sha,
    get_presentation_dimensions,
    verify_output_directory,
)
from modelmaker import sha

ROOT = Path(__file__).resolve().parent


class BuildPackageTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_forbidden_directories_rejected(self):
        # Repositories roots and protected source trees must be rejected
        with self.assertRaises(RuntimeError):
            verify_output_directory(ROOT)
        with self.assertRaises(RuntimeError):
            verify_output_directory(ROOT / "originals")
        with self.assertRaises(RuntimeError):
            verify_output_directory(ROOT / "proposals")
        with self.assertRaises(RuntimeError):
            verify_output_directory(ROOT / "docs")
        with self.assertRaises(RuntimeError):
            verify_output_directory(ROOT / "work")
        with self.assertRaises(RuntimeError):
            verify_output_directory(ROOT / ".git")

    def test_overwrite_rejected_without_force(self):
        out_dir = self.test_dir / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        # Create an existing file inside output directory
        (out_dir / "dummy.txt").write_text("existing content", encoding="utf-8")

        req_path = ROOT / "examples" / "marksans-package.json"
        with self.assertRaises(RuntimeError) as cm:
            build_package(
                request_path=req_path,
                output_dir=out_dir,
                allow_proposal=True,
                force=False,
                skip_com=True,
            )
        self.assertIn("already exists and is not empty", str(cm.exception))
        self.assertIn("--force", str(cm.exception))

    def test_force_rejected_on_non_package_arbitrary_dir(self):
        # If directory is not a known run dir and lacks package.manifest.json, force must fail
        # Create a mock arbitrary dir outside temp
        arbitrary_dir = ROOT / "test_scratch_nonpackage"
        try:
            arbitrary_dir.mkdir(parents=True, exist_ok=True)
            (arbitrary_dir / "important_file.txt").write_text("critical data", encoding="utf-8")
            with self.assertRaises(RuntimeError) as cm:
                verify_output_directory(arbitrary_dir, force=True)
            self.assertTrue(
                "does not contain a package.manifest.json" in str(cm.exception)
                or "fresh output directory" in str(cm.exception)
            )
        finally:
            import shutil
            shutil.rmtree(arbitrary_dir, ignore_errors=True)

    def test_com_capabilities_structure(self):
        caps = check_com_capabilities()
        self.assertIn("excel_com", caps)
        self.assertIn("powerpoint_com", caps)
        self.assertIn("platform", caps)
        self.assertIsInstance(caps["excel_com"], bool)
        self.assertIsInstance(caps["powerpoint_com"], bool)

    def test_dynamic_presentation_dimensions(self):
        w, h = get_presentation_dimensions(ROOT / "originals" / "ERR-4.pptx")
        self.assertEqual(w, 1080)
        self.assertEqual(h, 1440)

    def test_source_sha_mismatch_rejected(self):
        out_dir = self.test_dir / "output"
        req_text = (ROOT / "examples" / "marksans-package.json").read_text(encoding="utf-8")
        req = json.loads(req_text)
        # Corrupt expected source SHA
        req["workbook"]["source_sha256"] = "0" * 64
        corrupted_req = self.test_dir / "bad_req.json"
        corrupted_req.write_text(json.dumps(req), encoding="utf-8")

        with self.assertRaises(ValueError) as cm:
            build_package(
                request_path=corrupted_req,
                output_dir=out_dir,
                allow_proposal=True,
                force=True,
                skip_com=True,
            )
        self.assertIn("Workbook source SHA mismatch", str(cm.exception))

    def test_template_sha_mismatch_rejected(self):
        out_dir = self.test_dir / "output"
        req_text = (ROOT / "examples" / "marksans-package.json").read_text(encoding="utf-8")
        req = json.loads(req_text)
        # Corrupt expected template SHA
        req["presentation"]["template_sha256"] = "f" * 64
        corrupted_req = self.test_dir / "bad_tpl_req.json"
        corrupted_req.write_text(json.dumps(req), encoding="utf-8")

        with self.assertRaises(ValueError) as cm:
            build_package(
                request_path=corrupted_req,
                output_dir=out_dir,
                allow_proposal=True,
                force=True,
                skip_com=True,
            )
        self.assertIn("Presentation template SHA mismatch", str(cm.exception))

    def test_unapproved_pack_rejected_without_allow_proposal(self):
        out_dir = self.test_dir / "output"
        req_path = ROOT / "examples" / "marksans-package.json"

        with self.assertRaises(Exception) as cm:
            build_package(
                request_path=req_path,
                output_dir=out_dir,
                allow_proposal=False,
                force=True,
                skip_com=True,
            )
        err_msg = str(cm.exception).lower()
        self.assertTrue("proposal" in err_msg or "approval" in err_msg or "blocked" in err_msg)

    def test_build_package_baseline_end_to_end_skip_com(self):
        out_dir = self.test_dir / "output_baseline"
        req_path = ROOT / "examples" / "marksans-package.json"

        manifest = build_package(
            request_path=req_path,
            output_dir=out_dir,
            allow_proposal=True,
            force=True,
            skip_com=True,
        )

        self.assertEqual(manifest["package_id"], "marksans-curated-overview-2025")
        self.assertEqual(manifest["status"], "proposal")
        self.assertFalse(manifest["ready"])
        self.assertIn("human_pack_approval", manifest["pending_gates"])
        self.assertIn("render_review", manifest["pending_gates"])
        self.assertIn("com_capabilities", manifest)

        # Check output artifacts
        wb_file = out_dir / "Marksans-Recalculated.xlsx"
        pres_file = out_dir / "Marksans-Pharma-Overview.pptx"
        manifest_file = out_dir / "package.manifest.json"

        self.assertTrue(wb_file.exists())
        self.assertTrue(pres_file.exists())
        self.assertTrue(manifest_file.exists())

        # Verify no self-referential manifest hash
        artifact_files = [a["file"] for a in manifest["output_artifacts"]]
        self.assertNotIn("package.manifest.json", artifact_files)

        # Check reconciliation table entries
        reconciled_metrics = {m["metric"]: m for m in manifest["reconciliation_summary"]}
        self.assertEqual(reconciled_metrics["Risk-Free Rate (Rf)"]["displayed"], "7.30%")
        self.assertEqual(reconciled_metrics["Pre-tax Cost of Debt"]["displayed"], "8.00%")
        self.assertEqual(reconciled_metrics["Post-tax Cost of Debt (Kd)"]["displayed"], "5.60%")
        self.assertEqual(reconciled_metrics["Cost of Equity (Ke)"]["displayed"], "12.26%")
        self.assertEqual(reconciled_metrics["WACC (Target-Weighted)"]["displayed"], "12.11%")
        self.assertEqual(reconciled_metrics["FY25 Current Ratio"]["displayed"], "4.95")
        self.assertEqual(reconciled_metrics["FY25 Quick Ratio"]["displayed"], "3.07")

    def test_build_package_scenario_b_directional_effects(self):
        out_dir = self.test_dir / "output_scenario_b"
        req_path = ROOT / "examples" / "scenario-b-package.json"

        manifest = build_package(
            request_path=req_path,
            output_dir=out_dir,
            allow_proposal=True,
            force=True,
            skip_com=True,
        )

        self.assertEqual(manifest["package_id"], "marksans-scenario-b-stress-2025")
        self.assertFalse(manifest["ready"])

        reconciled_metrics = {m["metric"]: m for m in manifest["reconciliation_summary"]}
        # Directional verification: Rf increases to 8.00%
        self.assertEqual(reconciled_metrics["Risk-Free Rate (Rf)"]["displayed"], "8.00%")
        # Pre-tax Kd increases to 9.00%, post-tax Kd increases to 6.30%
        self.assertEqual(reconciled_metrics["Pre-tax Cost of Debt"]["displayed"], "9.00%")
        self.assertEqual(reconciled_metrics["Post-tax Cost of Debt (Kd)"]["displayed"], "6.30%")
        # Ke increases from 12.26% to 12.57%
        self.assertEqual(reconciled_metrics["Cost of Equity (Ke)"]["displayed"], "12.57%")
        # WACC increases from 12.11% to 12.43%
        self.assertEqual(reconciled_metrics["WACC (Target-Weighted)"]["displayed"], "12.43%")
        # Historical ratio metrics remain 100% identical
        self.assertEqual(reconciled_metrics["FY25 Current Ratio"]["displayed"], "4.95")
        self.assertEqual(reconciled_metrics["FY25 Quick Ratio"]["displayed"], "3.07")

    def test_workbook_inventory_matches_actual_sheets(self):
        """Regression test: verify candidate workbook has exactly 17 sheets and no hallucinated names."""
        import openpyxl
        wb_path = ROOT / "originals" / "Marksans Keshav.xlsx"
        wb = openpyxl.load_workbook(str(wb_path), read_only=True)
        actual_sheets = wb.sheetnames
        expected_sheets = [
            "Model ->",
            "Main Financials",
            "Ratio Analysis",
            "Intrinsic Growth",
            "Comps",
            "WACC-DCF",
            "DCF 1st way",
            "DCF 2nd way",
            "Data ->",
            "compdata",
            "Data Sheet",
            "BS",
            "Beta",
            "RM",
            "peerdata",
            "Cash Flow",
            "sectordata",
        ]
        self.assertEqual(actual_sheets, expected_sheets)
        # Ensure none of the 7 hallucinated sheet names exist
        forbidden_hallucinated = [
            "PL",
            "common size statement",
            "historic forecast",
            "FCFF",
            "SOP",
            "cost of capital",
            "dividend discount model",
        ]
        for bad_sheet in forbidden_hallucinated:
            self.assertNotIn(bad_sheet, actual_sheets)

    def test_clean_checkout_guide_hashes_and_request_schema(self):
        """Regression test: verify AI_CLEAN_CHECKOUT_GUIDE.md has verified source hashes and valid request JSON."""
        import jsonschema
        guide_text = (ROOT / "docs" / "AI_CLEAN_CHECKOUT_GUIDE.md").read_text(encoding="utf-8")
        catalog = json.loads((ROOT / "source-catalog.json").read_text(encoding="utf-8"))
        for item in catalog["files"]:
            self.assertIn(
                item["sha256"],
                guide_text,
                f"Source hash for {item['file']} missing from AI_CLEAN_CHECKOUT_GUIDE.md",
            )

        # Verify example JSON request in guide validates against schema
        import re
        m = re.search(r"```json\s*(\{[\s\S]*?\})\s*```", guide_text)
        self.assertIsNotNone(m, "Could not find example JSON block in clean-checkout guide")
        example_json = json.loads(m.group(1))
        schema = json.loads((ROOT / "package-request.schema.json").read_text(encoding="utf-8"))
        # Should validate without raising jsonschema.ValidationError
        jsonschema.validate(instance=example_json, schema=schema)

    def test_reconciliation_summary_verified_identities_and_statuses(self):
        """Regression test: verify reconciliation summary evaluates mathematical identities and honest status labels."""
        from build_package import evaluate_reconciliation_summary
        wb_path = ROOT / "candidates" / "run_native_baseline_v5" / "Marksans-Recalculated.xlsx"
        if not wb_path.exists():
            wb_path = ROOT / "candidates" / "run_native_baseline_v4" / "Marksans-Recalculated.xlsx"
        if not wb_path.exists():
            wb_path = ROOT / "originals" / "Marksans Keshav.xlsx"
        summary = evaluate_reconciliation_summary(wb_path)
        metrics = {m["metric"]: m for m in summary}
        self.assertEqual(metrics["Market Return (Rm)"]["status"], "verified_identity")
        self.assertEqual(metrics["Cost of Equity (Ke)"]["status"], "verified_identity")
        self.assertEqual(metrics["Post-tax Cost of Debt (Kd)"]["status"], "verified_identity")
        self.assertEqual(metrics["WACC (Target-Weighted)"]["status"], "verified_identity")
        self.assertEqual(metrics["Target Weight of Equity (Peer Median)"]["status"], "verified_identity")
        self.assertEqual(metrics["Target Weight of Debt (Peer Median)"]["status"], "verified_identity")
        self.assertEqual(metrics["FY25 Current Ratio"]["status"], "tied_to_presentation")
        self.assertEqual(metrics["FY25 Quick Ratio"]["status"], "tied_to_presentation")
        self.assertEqual(metrics["Risk-Free Rate (Rf)"]["status"], "sourced_assumption")
        self.assertEqual(metrics["Pre-tax Cost of Debt"]["status"], "sourced_assumption")
        self.assertEqual(metrics["Effective Tax Rate"]["status"], "statutory_assumption")

    def test_non_circular_rm_verification_and_strict_numeric_parsing(self):
        """Regression test: verify non-circular Rm evaluation and strict rejection of None/non-numeric."""
        from build_package import evaluate_reconciliation_summary, _parse_finite_float
        wb_path = ROOT / "candidates" / "run_native_baseline_v5" / "Marksans-Recalculated.xlsx"
        if not wb_path.exists():
            wb_path = ROOT / "candidates" / "Marksans-Recalculated.xlsx"
        if not wb_path.exists():
            wb_path = ROOT / "originals" / "Marksans Keshav.xlsx"

        summary = evaluate_reconciliation_summary(wb_path)
        metrics = {m["metric"]: m for m in summary}
        # Non-circular check: exact_value should match RM!F12 (0.163692)
        self.assertEqual(metrics["Market Return (Rm)"]["exact_value"], "0.163692")
        self.assertEqual(metrics["Market Return (Rm)"]["status"], "verified_identity")

        # Test strict parsing rejects None, bool, non-numeric, or NaN/Inf
        with self.assertRaises(ValueError):
            _parse_finite_float(None, "Test!A1")
        with self.assertRaises(ValueError):
            _parse_finite_float(True, "Test!A1")
        with self.assertRaises(ValueError):
            _parse_finite_float("not_a_number", "Test!A1")
        with self.assertRaises(ValueError):
            _parse_finite_float(float("nan"), "Test!A1")
        with self.assertRaises(ValueError):
            _parse_finite_float(float("inf"), "Test!A1")

    def test_audit_retained_content_passes_with_zero_stale_elements(self):
        """Regression test: audit_retained_content passes with 0 unresolved stale elements."""
        from work.audit_retained_content import audit_presentation
        tpl_path = ROOT / "originals" / "ERR-4.pptx"
        # Compose in-memory candidate deck
        from pptmaker import compose_deck
        tpl_bytes = tpl_path.read_bytes()
        deck_pack = json.loads((ROOT / "proposals" / "err4.deck-pack.proposal.json").read_text(encoding="utf-8"))
        deck_inputs = json.loads((ROOT / "proposals" / "err4.inputs.proposal.json").read_text(encoding="utf-8"))
        wb_bytes = (ROOT / "candidates" / "Marksans-Recalculated.xlsx").read_bytes()
        cand_bytes, _ = compose_deck(
            template=tpl_bytes,
            pack=deck_pack,
            supplied=deck_inputs,
            workbook_bytes=wb_bytes,
            allow_proposal=True,
        )
        audit_data = audit_presentation(tpl_path, candidate_bytes=cand_bytes)
        self.assertTrue(audit_data["audit_passed"])
        self.assertEqual(audit_data["unresolved_stale_elements_count"], 0)
        self.assertGreater(audit_data["total_elements_audited"], 200)


if __name__ == "__main__":
    unittest.main()
