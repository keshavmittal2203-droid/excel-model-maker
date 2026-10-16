"""Unit and regression tests for build_package.py orchestrator."""
from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from build_package import build_package, compute_file_sha
from modelmaker import sha

ROOT = Path(__file__).resolve().parent


class BuildPackageTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_overwrite_rejected_without_force(self):
        out_dir = self.test_dir / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        # Create a file inside output directory
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

    def test_build_package_end_to_end_skip_com(self):
        out_dir = self.test_dir / "output"
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

        # Check output artifacts
        wb_file = out_dir / "Marksans-Recalculated.xlsx"
        pres_file = out_dir / "Marksans-Pharma-Overview.pptx"
        manifest_file = out_dir / "package.manifest.json"

        self.assertTrue(wb_file.exists())
        self.assertTrue(pres_file.exists())
        self.assertTrue(manifest_file.exists())

        # Check reconciliation table entries
        reconciled_metrics = {m["metric"]: m for m in manifest["reconciliation_summary"]}
        self.assertIn("Cost of Equity (Ke)", reconciled_metrics)
        self.assertEqual(reconciled_metrics["Cost of Equity (Ke)"]["displayed"], "12.26%")
        self.assertIn("WACC (Target-Weighted)", reconciled_metrics)
        self.assertEqual(reconciled_metrics["WACC (Target-Weighted)"]["displayed"], "12.11%")
        self.assertIn("FY25 Current Ratio", reconciled_metrics)
        self.assertEqual(reconciled_metrics["FY25 Current Ratio"]["displayed"], "4.95")
        self.assertIn("FY25 Quick Ratio", reconciled_metrics)
        self.assertEqual(reconciled_metrics["FY25 Quick Ratio"]["displayed"], "3.07")


if __name__ == "__main__":
    unittest.main()
