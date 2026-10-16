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
            self.assertIn("does not contain a package.manifest.json", str(cm.exception))
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


if __name__ == "__main__":
    unittest.main()
