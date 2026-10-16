## Summary of Implementation & Curated Package Delivery (v6+ & Cipla Packaging v3)

Substantiates and hardens PR #3 as the implementation owner, reconciling `main`'s standalone styled report builder, integrating Codex PR #4 (`agent/codex-template-first`) and Codex Critical Safety PR #7 (`1395b247...`), eradicating stale template rows, implementing a complete financial audit and primary statutory source verification for Cipla Limited (v3), reconciling all report contradictions, completely eradicating artificial data used to make checks pass (Beta flat prices and Cash Flow zeros replaced with None), implementing excluded schedules governance (`excluded_schedules: ["Beta"]`), adding 17 negative regression safety tests, and delivering a clean-checkout reproducible guide with paste-ready AI instructions on `agent/argon-continue`.

### 1. Repeatable Build Orchestrator (`build_package.py`)
- Created unified build orchestrator with formal request schema (`package-request.schema.json`) and shipped package requests (`examples/marksans-package.json`, `examples/scenario-b-package.json`, `examples/cipla-package.json`).
- Baseline Command (requires a fresh, non-existent output directory):
  ```powershell
  python build_package.py --request examples/marksans-package.json --output-dir candidates/run_native_baseline_v6 --allow-proposal
  ```
- Scenario B (Stress Forecast) Command:
  ```powershell
  python build_package.py --request examples/scenario-b-package.json --output-dir candidates/run_native_scenario_b_v6 --allow-proposal
  ```
- Cipla (Second Company Demonstration v3) Command:
  ```powershell
  python build_package.py --request examples/cipla-package.json --output-dir candidates/run_native_cipla_v3 --allow-proposal
  ```
- Fast Preflight Dry-Run Mode (`--check-only` / `--preflight-only`):
  ```powershell
  python build_package.py --request examples/cipla-package.json --output-dir candidates/preflight_run --check-only
  ```
  Validates request schema, template SHA-256 integrity, pack presence, and COM registration in milliseconds with zero disk writes and zero COM application activation.
- **Strict Destination Overwrite Protection**: Refuses to clobber existing output files or reuse non-empty directories, even with `--force`.
- **Pre-Write Path & Adapter Safety**:
  - Validates request-controlled output filenames via `output_file()` before any source file reads, directory creation, or Office activation. Rejects traversal (`../`), absolute paths, Windows device names (`CON`), and alternate data streams (`file:stream`).
  - Restricts workbook adapters to a strict whitelist (`REGISTERED_ADAPTERS = {"work/adapt_cipla_workbook.py": "adapt_workbook_for_cipla"}`). Unregistered scripts or functions raise an immediate `ValueError`.
  - Enforces containment within `work/` for adapters, repository root for candidate sources, and run-local `assets_dir` for dynamic images.
- **Pre-Flight Request Validation**: Validates `schema_version == 1`, template SHA-256 integrity, company metadata, reporting units, company cross-matching, empty value rejection, and complete slot mapping before directory creation or Office COM activation. Missing inputs **never silently become zero or retain prior company data**.
- **Office Process Isolation**: Excel recalculation uses dedicated `DispatchEx` instance with links and macros disabled (`AutomationSecurity = 3`, `AskToUpdateLinks = False`). PowerPoint export closes only its owned presentation without calling application `Quit()`. COM probes check Windows registry without activating user Office sessions.
- **Run-Local Visual Asset Isolation**: Dynamic replacement table images generate exclusively inside `candidates/<run_dir>/assets/` with zero writes to `work/assets/`.
- **Automated Office COM Previews**: Exports 1080x1440 PNG slide previews via PowerPoint COM, composite slide contact sheets (`contact_sheet.png`), and landscape `FitToPagesWide = 1` multi-column schedule PDFs via Excel COM.
- **Package Manifest (`package.manifest.json`)**: Records execution timestamp, COM capabilities, source hashes, validation results, mathematical identity verification table, and generated artifact SHA-256 hashes (zero self-referential manifest hash loops).

### 2. Primary Statutory Source Verification & Dedicated Evidence Registry
- Authored [`docs/PRIMARY_SOURCE_REGISTRY.md`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/PRIMARY_SOURCE_REGISTRY.md) and [`docs/primary_source_registry.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/primary_source_registry.json).
- Segregated figures into two transparent tiers:
  - **FY21 to FY25 (5 Years)**: 100% sourced from primary BSE/NSE statutory filings and Walker Chandiok & Co LLP independent audited reports (Cipla 89th, 88th, 87th Annual Reports).
  - **FY17 to FY20 (4 Years)**: Sourced from verified secondary financial databases (Screener.in / Capitaline) due to earlier Ind AS standard transitions.
- Line items mapped with exact URLs, printed vs PDF pages, statement and note citations, reported values, units, and mathematical transformations.

### 3. Resolution of All Report Contradictions
- **Total Borrowings Reconciled (₹438 Cr vs ₹613.77 Cr)**:
  - **₹438.00 Cr**: Represents Cipla's audited statutory balance sheet close as of 31 March 2025 (Cipla 89th Annual Report Note 23: Non-current borrowings ₹11.98 Cr, Current facilities ₹185.53 Cr, Lease liabilities under Ind AS 116 ₹240.49 Cr). Resides in `Data Sheet!J59` and is cited in the presentation narrative.
  - **₹613.77 Cr**: Represents the October 2025 trailing twelve-month (TTM) market snapshot debt from `compdata!G8` and placed in `Data Sheet!K59`, directly feeding `WACC-DCF!C30`.
  - Both figures are authentic and belong to different time horizons; fully reconciled and disclosed.
- **Cash Flow Reconciliation Claim Reconciled**:
  - Across all 8 historical years (FY18–FY25), CFO + CFI + CFF = Net Cash Flow ties out to $\Delta = ₹0.00$ Cr tolerance in both `Data Sheet` and `Cash Flow`. FY25 ties out: $+5,005 - 3,682 - 1,293 = +30.00$ Cr.
  - Unbacked granular sub-schedule intermediate breakdown rows are withheld as transparent em-dashes (`"—"`) on Slide 14 and empty cells (`None`) in the workbook, completely eradicating synthetic balancing plugs and zero loops.
- **Market Price & Non-Circular Statutory Share Count**:
  - CMP ₹1,301.00 is an inherited market snapshot from `compdata!D8`.
  - Statutory share count is derived directly from Note 20 paid-up equity: ₹161.52 Cr / ₹2.00 face value = **80.76 crore statutory shares**.
  - Template implied share count `=B9/B8` evaluates to **80.79 crore shares** (0.04% variance representing post-FY25 ESOP allotments and market capitalisation rounding, explicitly disclosed).

### 4. Complete Eradication of Artificial Data Used to Make Checks Pass
- **`Beta` Sheet**:
  - Eradicated 107 artificial repeated ₹1,301 flat prices in `Beta!C8:C114` by setting them to `None` (empty cells).
  - Updated header in `Beta!B4` to mark weekly regression series unavailable and schedule excluded.
  - Added `"excluded_schedules": ["Beta"]` to `model-pack.schema.json` and `proposals/cipla.model-pack.proposal.json`.
  - Updated `validate_workbook.py` to route and isolate formula errors on excluded schedules while enforcing 0 errors on all active schedules.
- **`Cash Flow` Detail Rows**:
  - Replaced loop setting unbacked sub-schedule rows to `0.0` with `None` (empty cells). Supported statutory lines (PBT, taxes, CFO, CapEx, CFI, interest, dividends, CFF, Net CF) are preserved intact.
- **`Data Sheet` Column B (FY17) & Quarterly Rows**:
  - Populated Column B with authentic Cipla FY17 consolidated financials (Revenue ₹14,630 Cr, PBT ₹1,222 Cr, PAT ₹1,006 Cr, Assets ₹20,094 Cr, Equity ₹161.47 Cr, Borrowings ₹4,113 Cr, Cash ₹1,328 Cr).
  - Cleared Marksans quarterly rows 42–50 across all columns B to K to `None`. 100% of Marksans carryover eradicated.
- **`BS` Columns J to N (FY22 to FY26)**:
  - Formally documented that Column N represents **FY26 forward estimate / TTM snapshot**, not an audited historical year.

### 5. Valuation Methodology Disclosures (DCF 1st Way vs. DCF 2nd Way)
- Thoroughly audited and documented both DCF models:
  - **`DCF 2nd way` (10-Year Fundamental ROIC Model)**: Uses Cipla's authentic balance sheet reinvestment ($4.21\%$) and ROIC ($29.13\%$) drivers, yielding an intrinsic value of **₹1,217.21 per share** (within 6.4% of CMP ₹1,301.00).
  - **`DCF 1st way` (5-Year Revenue/Margin FCFF Model)**: Evaluates to **₹319.37 per share** because the template hardcoded an arbitrary **45% reinvestment rate** assumption in `F11`, artificially penalizing Cipla with massive unevidenced CapEx deductions.

### 6. Focused Negative Regression Test Suite (17 Tests)
- Implemented targeted OOXML part mutation and validation testing proving safety guards fail when defects are reintroduced:
  - `test_repeated_price_history_detection`: Detects repetitive flat price histories.
  - `test_unknown_inputs_set_to_zero_rejected`: Rejects unbacked breakdown inputs silently set to 0.0.
  - `test_debt_definition_inconsistency_detection`: Enforces reconciliation of differing debt definitions.
  - `test_non_circular_share_count_verification`: Derives statutory share count from paid-up capital and face value.
  - `test_excluded_schedule_error_handling`: Verifies formula errors on excluded schedules are isolated from active calculation validation.
  - `test_unregistered_adapter_rejected`
  - `test_adapter_path_traversal_rejected`
  - `test_candidate_source_path_traversal_rejected`
  - `test_check_only_preflight_mode`
  - `test_audit_fails_when_original_template_numeric_row_reintroduced`
  - `test_audit_fails_when_chart_cache_or_embedded_workbook_is_stale`
  - `test_audit_fails_when_destination_disagrees_with_workbook_binding`
  - `test_audit_fails_when_unsupported_company_specific_claim_remains`
  - `test_audit_fails_when_slide23_raster_screenshot_is_unreplaced`
  - `test_company_mismatch_rejected`
  - `test_unmapped_or_empty_deck_input_rejected`
  - `test_invalid_reporting_unit_rejected`

### 7. Second Company Demonstration: Cipla Limited (All 5 Pipeline Stages Completed)
- **Stage 1 (Specifications Prepared)**: `examples/cipla-package.json` with audited Ind AS facts and accounting conventions.
- **Stage 2 (Data Mapped)**: `work/adapt_cipla_workbook.py` adapted 752 cells across sheets; `proposals/cipla.inputs.proposal.json` populated with 219 slots (zero Marksans carryover, 48 granular slots marked `"—"`).
- **Stage 3 (Build Completed)**: `python build_package.py --request examples/cipla-package.json --output-dir candidates/run_native_cipla_v3 --allow-proposal`. Produced 11 artifacts including `Cipla-Recalculated.xlsx`, `Cipla-Executive-Overview.pptx`, `package.manifest.json`, 9 replacement images, 5 slide previews, and 4 sheet PDFs.
- **Stage 4 (Financial Validation Passed)**: Recalculated workbook SHA: `3e1bbbd5...`. Balance sheet balances across all 10 years ($\Delta = 0.00$), Cash Flow ties across all 8 years ($\Delta = 0.00$), CAPM $K_e = 12.26\%$, $K_d = 5.25\%$, $WACC = 12.10\%$. 19/19 pack checks and 28/28 independent identities verified.
- **Stage 5 (Outputs Visually Inspected)**: Visually reviewed `slide_1.png` to `slide_5.png` and sheet PDFs via tool inspection. Confirmed layout, typography, charts, and zero text collisions.
- **Documentation**: Sourced citations in `docs/PRIMARY_SOURCE_REGISTRY.md` and `docs/SECOND_COMPANY_EVIDENCE_REPORT.md`.

### 8. Full Test Suite & Originals Integrity
- `python verify_sources.py`: Verified all 6 original template SHA-256 hashes byte-for-byte against `source-catalog.json`.
- `python -m unittest discover -v`: **95/95 tests passing** cleanly across:
  - `test_package_safety.py` (17 tests)
  - `test_build_package.py` (25 tests)
  - `test_modelmaker.py` (33 tests)
  - `test_pptmaker.py` (13 tests)
  - `test_reportmaker.py` (7 tests)
- Manifest readiness: `ready: false`, pending gates: `["human_pack_approval", "render_review"]`.
