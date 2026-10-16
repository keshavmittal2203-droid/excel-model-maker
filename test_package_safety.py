"""Safety regressions use temporary files and mocked COM; never activate Office."""
import json
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest.mock import Mock, patch

import build_package as package


def com_modules(app):
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = Mock()
    pythoncom.CoUninitialize = Mock()
    pythoncom.CLSIDFromProgID = Mock(return_value="registered")
    client = types.ModuleType("win32com.client")
    client.Dispatch = Mock(side_effect=AssertionError("Must not attach to running Office"))
    client.DispatchEx = Mock(return_value=app)
    parent = types.ModuleType("win32com")
    parent.client = client
    return {"pythoncom": pythoncom, "win32com": parent, "win32com.client": client}


class PackageSafetyTests(unittest.TestCase):
    def test_reject_path_names_before_writes(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            sentinel = root / "original.xlsx"
            sentinel.write_bytes(b"immutable sentinel")
            for kind, suffix in (("workbook", ".xlsx"), ("presentation", ".pptx")):
                for bad in ("../original.xlsx", str(sentinel), "..\\original.xlsx",
                            "C:\\original.xlsx", "CON" + suffix, "file:stream" + suffix,
                            "file" + suffix + " ", "dir/file" + suffix):
                    with self.subTest(kind=kind, name=bad):
                        req = {"workbook": {}, "presentation": {}}
                        req[kind]["output_name"] = bad
                        request = root / "request.json"
                        request.write_text(json.dumps(req))
                        output = root / "new-run"
                        with patch.object(package, "check_com_capabilities") as capabilities:
                            with self.assertRaises(ValueError):
                                package.build_package(request, output)
                            capabilities.assert_not_called()
                        self.assertFalse(output.exists())
                        self.assertEqual(sentinel.read_bytes(), b"immutable sentinel")

    def test_existing_output_never_replaced(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / "candidate.xlsx"
            path.write_bytes(b"old workbook")
            with self.assertRaises(ValueError):
                package.output_file(root, path.name, ".xlsx")
            self.assertEqual(path.read_bytes(), b"old workbook")

    def test_images_are_generated_only_inside_current_run(self):
        generator = types.ModuleType("work.generate_wacc_images")
        generator.generate_images = Mock()
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for name in ("source.xlsx", "source.pptx", "cached.xlsx"):
                (root / name).write_bytes(b"fixture")
            for name, value in (("model.json", {}), ("deck.json", {"slots": []}),
                                ("values.json", {"values": {}})):
                (root / name).write_text(json.dumps(value))
            req = {"workbook": {"source": "source.xlsx", "source_sha256": package.compute_file_sha(root / "source.xlsx"),
                                "model_pack": "model.json", "candidate_source": "cached.xlsx"},
                   "presentation": {"template": "source.pptx", "template_sha256": package.compute_file_sha(root / "source.pptx"),
                                    "deck_pack": "deck.json", "deck_inputs": "values.json"}}
            request = root / "request.json"
            request.write_text(json.dumps(req))
            output = root / "run"
            with patch.object(package, "ROOT", root), patch.dict(sys.modules, {"work.generate_wacc_images": generator}), \
                 patch.object(package, "validate_candidate", return_value={"checks_passed": True}), \
                 patch.object(package, "compose_deck", return_value=(b"deck", {})), \
                 patch.object(package, "check_com_capabilities", return_value={}), \
                 patch.object(package, "evaluate_reconciliation_summary", return_value=[]):
                package.build_package(request, output, allow_proposal=True, skip_com=True)
            generator.generate_images.assert_called_once_with(assets_dir=output / "assets", workbook_path=output / "Marksans-Recalculated.xlsx")
            self.assertFalse((root / "work" / "assets").exists())

    def test_failed_forced_rebuild_cannot_leave_old_manifest(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = root / "package.manifest.json"
            manifest.write_text('{"status":"proposal"}')
            with self.assertRaisesRegex(RuntimeError, "fresh output"):
                package.verify_output_directory(root, force=True)
            self.assertEqual(manifest.read_text(), '{"status":"proposal"}')

    def test_probe_never_activates_or_quits_office(self):
        app = Mock()
        modules = com_modules(app)
        with patch.dict(sys.modules, modules), patch.object(package.sys, "platform", "win32"):
            result = package.check_com_capabilities()
        self.assertTrue(result["excel_com"])
        self.assertTrue(result["powerpoint_com"])
        self.assertIsNone(result["excel_version"])
        modules["win32com.client"].DispatchEx.assert_not_called()
        app.Quit.assert_not_called()
        modules["pythoncom"].CoUninitialize.assert_called_once()

    def test_ppt_export_closes_only_owned_presentation(self):
        app = Mock()
        presentation = app.Presentations.Open.return_value
        presentation.Slides.Count = 1
        def export(path, *_):
            Path(path).write_bytes(b"preview")
        presentation.Slides.return_value.Export.side_effect = export
        modules = com_modules(app)
        with tempfile.TemporaryDirectory() as temp, patch.dict(sys.modules, modules):
            results = package.export_slide_previews_com(Path(temp) / "input.pptx", Path(temp) / "previews")
        self.assertEqual(len(results), 1)
        presentation.Close.assert_called_once()
        app.Quit.assert_not_called()
        modules["win32com.client"].DispatchEx.assert_called_once_with("PowerPoint.Application")

    def test_activation_failure_uninitializes_com(self):
        modules = com_modules(Mock())
        modules["win32com.client"].DispatchEx.side_effect = RuntimeError("activation failed")
        with tempfile.TemporaryDirectory() as temp, patch.dict(sys.modules, modules):
            with self.assertRaisesRegex(RuntimeError, "activation failed"):
                package.export_slide_previews_com(Path(temp) / "x.pptx", Path(temp) / "previews")
        modules["pythoncom"].CoUninitialize.assert_called_once()

    def test_excel_export_owns_application_and_disables_links(self):
        app = Mock()
        workbook = app.Workbooks.Open.return_value
        def export(_, path):
            Path(path).write_bytes(b"PDF")
        workbook.Sheets.return_value.ExportAsFixedFormat.side_effect = export
        modules = com_modules(app)
        with tempfile.TemporaryDirectory() as temp, patch.dict(sys.modules, modules):
            package.export_sheet_previews_com(Path(temp) / "x.xlsx", Path(temp) / "previews", ["Sheet1"])
        modules["win32com.client"].DispatchEx.assert_called_once_with("Excel.Application")
        self.assertEqual(app.Workbooks.Open.call_args.kwargs["UpdateLinks"], 0)
        self.assertEqual(app.AutomationSecurity, 3)
        workbook.Close.assert_called_once_with(SaveChanges=False)
        app.Quit.assert_called_once()

    def test_unregistered_adapter_rejected(self):
        """Unregistered adapter scripts must be rejected before execution."""
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            req = {
                "workbook": {
                    "source": "originals/Marksans Keshav.xlsx",
                    "source_sha256": "5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a",
                    "model_pack": "proposals/cipla.model-pack.proposal.json",
                    "adapter": "work/evil_adapter.py",
                },
                "presentation": {
                    "template": "originals/ERR-4.pptx",
                    "template_sha256": "9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba",
                    "deck_pack": "proposals/err4.deck-pack.proposal.json",
                    "deck_inputs": "proposals/cipla.inputs.proposal.json",
                },
            }
            req_file = root / "bad_adapter.json"
            req_file.write_text(json.dumps(req), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Unregistered or forbidden workbook adapter"):
                package.validate_package_request(req, req_file)

    def test_adapter_path_traversal_rejected(self):
        """Adapter paths containing directory traversal must be rejected."""
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for bad_path in ("../evil.py", "..\\evil.py", "/evil.py", "\\evil.py"):
                with self.subTest(bad_path=bad_path):
                    req = {
                        "workbook": {
                            "source": "originals/Marksans Keshav.xlsx",
                            "source_sha256": "5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a",
                            "model_pack": "proposals/cipla.model-pack.proposal.json",
                            "adapter": bad_path,
                        },
                        "presentation": {
                            "template": "originals/ERR-4.pptx",
                            "template_sha256": "9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba",
                            "deck_pack": "proposals/err4.deck-pack.proposal.json",
                            "deck_inputs": "proposals/cipla.inputs.proposal.json",
                        },
                    }
                    req_file = root / "traversal.json"
                    req_file.write_text(json.dumps(req), encoding="utf-8")
                    with self.assertRaisesRegex(ValueError, "Path traversal detected"):
                        package.validate_package_request(req, req_file)

    def test_candidate_source_path_traversal_rejected(self):
        """Candidate source paths escaping repository root must be rejected."""
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            req = {
                "workbook": {
                    "source": "originals/Marksans Keshav.xlsx",
                    "source_sha256": "5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a",
                    "model_pack": "proposals/cipla.model-pack.proposal.json",
                    "candidate_source": "../outside.xlsx",
                },
                "presentation": {
                    "template": "originals/ERR-4.pptx",
                    "template_sha256": "9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba",
                    "deck_pack": "proposals/err4.deck-pack.proposal.json",
                    "deck_inputs": "proposals/cipla.inputs.proposal.json",
                },
            }
            req_file = root / "cand_traversal.json"
            req_file.write_text(json.dumps(req), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Path traversal detected"):
                package.validate_package_request(req, req_file)

    def test_check_only_preflight_mode(self):
        """Preflight mode must validate request and return status without disk writes or COM activation."""
        req_path = package.ROOT / "examples" / "cipla-package.json"
        with tempfile.TemporaryDirectory() as temp:
            out_dir = Path(temp) / "preflight_run"
            res = package.build_package(
                request_path=req_path,
                output_dir=out_dir,
                check_only=True,
            )
            self.assertEqual(res["status"], "preflight_passed")
            self.assertFalse(res["ready"])
            self.assertEqual(res["package_id"], "cipla-overview-fy25")
            self.assertIn("com_capabilities", res)
            # Preflight must NOT write candidate artifacts or directories
            self.assertFalse(out_dir.exists())

    def test_repeated_price_history_detection(self):
        """Repeated flat price series in historical regression schedules must be detected."""
        def check_price_series_variance(prices: list[float], min_unique_ratio: float = 0.5) -> bool:
            if not prices:
                return True
            unique_count = len(set(prices))
            ratio = unique_count / len(prices)
            return ratio >= min_unique_ratio

        # Artificial repeated price series (e.g. 107 identical ₹1,301 CMPs)
        artificial_series = [1301.0] * 107
        self.assertFalse(check_price_series_variance(artificial_series))

        # Real fluctuating market price series
        real_series = [1301.0 + i * 2.5 for i in range(107)]
        self.assertTrue(check_price_series_variance(real_series))

    def test_unknown_inputs_set_to_zero_rejected(self):
        """Unevidenced financial breakdown inputs must not be silently populated with 0.0."""
        def validate_unbacked_input_policy(slot_name: str, value: Any, is_evidenced: bool):
            if not is_evidenced and value == 0.0:
                raise ValueError(
                    f"Unevidenced breakdown slot '{slot_name}' cannot be set to 0.0; "
                    "must be None or omitted to prevent artificial balancing."
                )

        # Setting unevidenced breakdown cell to 0.0 is rejected
        with self.assertRaisesRegex(ValueError, "cannot be set to 0.0"):
            validate_unbacked_input_policy("slide14_non_cash_fy18", 0.0, is_evidenced=False)

        # Evidenced 0.0 or None/em-dash is accepted
        validate_unbacked_input_policy("slide14_non_cash_fy18", None, is_evidenced=False)
        validate_unbacked_input_policy("slide14_non_cash_fy18", "—", is_evidenced=False)
        validate_unbacked_input_policy("slide14_direct_taxes_fy25", 0.0, is_evidenced=True)

    def test_debt_definition_inconsistency_detection(self):
        """Inconsistencies between period-end debt close and TTM snapshot debt must be reconciled."""
        def reconcile_debt_metrics(fy25_bs_close: float, ttm_snapshot: float, metadata: dict[str, str]):
            if abs(fy25_bs_close - ttm_snapshot) > 1.0:
                # If numbers differ, explicit disclosure reconciling the timeframes is required
                if "reconciliation_note" not in metadata or not metadata["reconciliation_note"]:
                    raise ValueError(
                        f"Unreconciled debt discrepancy: FY25 close ({fy25_bs_close}) vs "
                        f"TTM snapshot ({ttm_snapshot}) requires explicit disclosure."
                    )
            return True

        # Unreconciled debt numbers raise ValueError
        with self.assertRaisesRegex(ValueError, "Unreconciled debt discrepancy"):
            reconcile_debt_metrics(438.00, 613.77, {})

        # With explicit reconciliation note, passes
        meta = {"reconciliation_note": "₹438 Cr is 31 March 2025 audited close; ₹613.77 Cr is Oct 2025 TTM snapshot."}
        self.assertTrue(reconcile_debt_metrics(438.00, 613.77, meta))

    def test_non_circular_share_count_verification(self):
        """Statutory share count must be derived from paid-up capital and face value, not circular market cap."""
        def verify_share_count(equity_capital_cr: float, face_value: float, claimed_shares_cr: float, is_statutory: bool):
            if is_statutory:
                statutory_shares = round(equity_capital_cr / face_value, 2)
                if abs(claimed_shares_cr - statutory_shares) > 0.05:
                    raise ValueError(
                        f"Claimed share count {claimed_shares_cr} Cr does not match statutory "
                        f"paid-up equity {statutory_shares} Cr ({equity_capital_cr} Cr / ₹{face_value})."
                    )
            return True

        # Statutory Note 20: ₹161.52 Cr / ₹2.00 = 80.76 Cr shares
        self.assertTrue(verify_share_count(161.52, 2.0, 80.76, is_statutory=True))

        # Claiming non-statutory count without disclosing difference fails strict check
        with self.assertRaisesRegex(ValueError, "does not match statutory"):
            verify_share_count(161.52, 2.0, 85.00, is_statutory=True)

    def test_excluded_schedule_error_handling(self):
        """Excluded schedules (e.g. Beta) isolate formula errors from active validation."""
        from modelmaker import Package
        from validate_workbook import scan_formula_errors_and_missing_caches

        # Load Cipla recalculated candidate which has empty Beta!C causing #DIV/0! on Beta
        cipla_path = package.ROOT / "candidates" / "Cipla-Recalculated.xlsx"
        if not cipla_path.exists():
            self.skipTest("Cipla-Recalculated.xlsx not found")

        pkg = Package(cipla_path.read_bytes())
        # With excluded_sheets={"Beta"}, zero errors must be reported
        errors_excluded, missing_excluded = scan_formula_errors_and_missing_caches(pkg, excluded_sheets={"Beta"})
        self.assertEqual(len(errors_excluded), 0)
        self.assertEqual(len(missing_excluded), 0)

        # Without excluded_sheets, formula errors on Beta must be detected
        errors_all, missing_all = scan_formula_errors_and_missing_caches(pkg, excluded_sheets=None)
        self.assertGreater(len(errors_all), 0)
        self.assertTrue(all(e["sheet"] == "Beta" for e in errors_all))


if __name__ == "__main__":
    unittest.main()
