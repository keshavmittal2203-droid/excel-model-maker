## Summary of Implementation & Curated Package Delivery (v6+ & Cipla Packaging)

Substantiates and hardens PR #3 as the implementation owner, reconciling `main`'s standalone styled report builder, integrating Codex PR #4 (`agent/codex-template-first`) and Codex Critical Safety PR #7 (`1395b247...`), eradicating stale Sun Pharma detail rows from the Slide 14 cash flow schedule, implementing and hardening a machine-readable content audit across all retained slides, enforcing non-circular reconciliation checks with strict numeric parsing, adding 8 negative regression safety tests, introducing canonical company-request intake with pre-flight validation, providing a complete 5-stage second-company adaptation with Cipla Limited, and delivering a clean-checkout reproducible guide with paste-ready AI instructions on `agent/argon-continue`.

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
- Cipla (Second Company Demonstration) Command:
  ```powershell
  python build_package.py --request examples/cipla-package.json --output-dir candidates/run_native_cipla_v1 --allow-proposal
  ```
- **Strict Destination Overwrite Protection**: Refuses to clobber existing output files or reuse non-empty directories, even with `--force`.
- **Pre-Write Path Safety**: Validates all request-controlled output filenames before writes; rejects traversal (`../`), absolute paths, Windows device names (`CON`), and alternate data streams (`file:stream`).
- **Pre-Flight Request Validation**: Validates `schema_version == 1`, template SHA-256 integrity, company metadata, reporting units, company cross-matching, empty value rejection, and complete slot mapping before directory creation or Office COM activation. Missing inputs **never silently become zero or retain prior company data**.
- **Office Process Isolation**: Excel recalculation uses dedicated `DispatchEx` instance with links and macros disabled (`AutomationSecurity = 3`, `AskToUpdateLinks = False`). PowerPoint export closes only its owned presentation without calling application `Quit()`. COM probes check Windows registry without activating user Office sessions.
- **Run-Local Visual Asset Isolation**: Dynamic replacement table images generate exclusively inside `candidates/<run_dir>/assets/` with zero writes to `work/assets/`.
- **Automated Office COM Previews**: Exports 1080x1440 PNG slide previews via PowerPoint COM, composite slide contact sheets (`contact_sheet.png`), and landscape `FitToPagesWide = 1` multi-column schedule PDFs via Excel COM.
- **Package Manifest (`package.manifest.json`)**: Records execution timestamp, COM capabilities, source hashes, validation results, mathematical identity verification table, and generated artifact SHA-256 hashes (zero self-referential manifest hash loops).

### 2. Complete Slide 14 Cash Flow Reconciliation
- **All 15 Active Rows Replaced (120 Data Cells)**: Replaced every single cell in Slide 14 Table 0 across FY18–FY25 using reconciled Marksans data from `candidates/Marksans-Recalculated.xlsx` (sheets `Cash Flow` and `Data Sheet`):
  - Operating Activities: PBT, Non-cash adjustments (net), Working capital changes (net), Direct taxes paid, Cash From Operating Activities.
  - Investing Activities: Capital expenditure (net), Net purchase/(sale) of investments, Other investing items (net), Cash From Investing Activities.
  - Financing Activities: Net borrowings (proceeds/repay.), Finance costs paid, Dividends paid, Other financing items (net), Cash From Financing Activities.
  - Net Cash Flow: Cash from Operating + Investing + Financing.
- **Zero Sun Pharma Numbers Retained**: Completely eradicated all 10 stale Sun Pharma intermediate detail rows (`1,120.4`, `-323.6`, `-476.8`, `-2,067.6`, `-3,864.2`, `625.6`, `-1,164.9`, `-223.8`, `-3,617.3`, `-2,899.9`).
- **100% Mathematical Tie-Out**: Operating, Investing, and Financing subtotals tie out with 0.00 difference across all 8 historical years (e.g. FY25 Operating: 503.96 + 69.70 - 241.00 - 126.00 = 206.66; Investing: -172.00 + 27.00 + 95.99 = -49.01; Financing: -6.01 - 2.39 - 27.19 - 29.56 = -65.15).

### 3. Hardened Machine-Readable Content Audit (`work/audit_retained_content.py`)
- Programmatic audit tool compares candidate presentation against `originals/ERR-4.pptx` across every retained slide (Slides 1, 2, 14, 16, 23):
  - Audits visible and hidden text shapes, tables, charts, raster pictures, and suppressed logos.
  - Verifies mathematical balance across all 8 years within 0.05 Cr.
  - Detects reintroduced template rows at row level (preventing false positives from coincidental single-cell collisions).
  - Inspects native chart caches (`<c:numCache>`) against template caches and verifies sync with embedded workbooks (`Microsoft_Excel_WorksheetN.xlsx`).
  - Verifies Slide 23 replacement images via cryptographic SHA-256 comparison against original screenshots.
  - Generates machine-readable ledger: `docs/retained_content_audit.json`.
  - Generates comprehensive markdown report: `docs/RETAINED_CONTENT_AUDIT.md`.
  - Result: **274 elements audited, 0 unresolved stale elements (`audit_passed: True`)**.

### 4. Focused Negative Regression Test Suite
- Implemented targeted OOXML part mutation and validation testing proving safety guards fail when defects are reintroduced:
  - `test_audit_fails_when_original_template_numeric_row_reintroduced`
  - `test_audit_fails_when_chart_cache_or_embedded_workbook_is_stale`
  - `test_audit_fails_when_destination_disagrees_with_workbook_binding`
  - `test_audit_fails_when_unsupported_company_specific_claim_remains`
  - `test_audit_fails_when_slide23_raster_screenshot_is_unreplaced`
  - `test_company_mismatch_rejected`
  - `test_unmapped_or_empty_deck_input_rejected`
  - `test_invalid_reporting_unit_rejected`
  - `test_build_package_cipla_end_to_end_skip_com`

### 5. Second Company Demonstration: Cipla Limited (All 5 Pipeline Stages Completed)
- **Stage 1 (Specifications Prepared)**: `examples/cipla-package.json` with audited Ind AS facts and accounting conventions.
- **Stage 2 (Data Mapped)**: `work/adapt_cipla_workbook.py` adapted 359 cells across 8 sheets; `proposals/cipla.inputs.proposal.json` populated with 219 slots (zero Marksans carryover).
- **Stage 3 (Build Completed)**: `python build_package.py --request examples/cipla-package.json --output-dir candidates/run_native_cipla_v1 --allow-proposal`. Produced 11 artifacts including `Cipla-Recalculated.xlsx`, `Cipla-Executive-Overview.pptx`, `package.manifest.json`, 9 replacement images, 5 slide previews, and 4 sheet PDFs.
- **Stage 4 (Financial Validation Passed)**: Recalculated workbook SHA: `050e13b2...`. Balance sheet balances across all 8 years ($\Delta = 0.00$), Cash Flow ties across all 8 years ($\Delta = 0.00$), CAPM $K_e = 12.26\%$, $K_d = 5.25\%$, $WACC = 12.10\%$. 19/19 pack checks and 28/28 independent identities verified.
- **Stage 5 (Outputs Visually Inspected)**: Visually reviewed `slide_1.png` to `slide_5.png` and sheet PDFs via tool inspection. Confirmed layout, typography, charts, and zero text collisions.
- **Documentation**: Sourced citations in `docs/SECOND_COMPANY_EVIDENCE_REPORT.md`; complete cell-level ledger in `docs/CIPLA_WORKBOOK_COVERAGE.md` and `docs/cipla_workbook_coverage.json`.

### 6. Capabilities Matrix

| Scope / Dimension | Classification | Description & Evidence |
|---|---|---|
| **Excel 17-Sheet XML Injection** | **Implemented** | Openpyxl-based targeted XML editing preserving styles, formulas, follower ranges, and data tables. |
| **PowerPoint Deck Assembly** | **Implemented** | Curated slide retention (1, 2, 14, 16, 23), table cell updates, chart cache sync, picture replacement, and logo suppression. |
| **Independent Formula & Identity Checks** | **Implemented** | Recalculation verification, zero `#REF!`/#DIV/0!, $WACC = W_e K_e + W_d K_d (1-t)$, $K_e = R_f + \beta(R_m-R_f)$, balance sheet balance ($\Delta = 0.00$). |
| **Retained-Content Audit** | **Implemented** | 274 elements checked across tables, shapes, charts, embedded workbooks, raster screenshots; zero stale content. |
| **Pre-Flight Request Validation** | **Implemented** | Rejects unmapped slots, company mismatches, empty values, bad hashes, traversal paths, and invalid reporting units before directory writes. |
| **Marksans Pharma Package (Baseline & Scenario B)** | **Demonstrated** | Fully recalculated v6 models (`ef5d6d5a...`), 5-slide decks, 1080x1440 PNG previews, landscape PDFs. |
| **Cipla Limited Second-Company Packaging** | **Demonstrated (All 5 Stages)** | Stage 1 (Specs prepared), Stage 2 (Data mapped), Stage 3 (Build completed), Stage 4 (Financial validation passed), Stage 5 (Outputs visually inspected). |
| **Non-Pharma / Non-Manufacturing Sectors (BFSI / Tech)** | **Manual Template Required** | Banking models (NIM/NPA) and SaaS (ARR/NRR) require fundamentally different 3-statement structures created by human analysts. |
| **Human Pack Approval & Promotion** | **Manual Review Required** | Transition from `status: "proposal"` to `status: "approved"` requires independent human financial review. |
| **Autonomous Agent Self-Approval** | **Unsupported / Prohibited** | Agent self-approval is blocked by reviewer guards and safety invariants. |
| **Fabrication of Missing Evidence** | **Unsupported / Prohibited** | Missing primary source data is cataloged under `unresolved_inputs`; inventing balancing numbers is strictly forbidden. |

### 7. Full Test Suite & Originals Integrity
- `python verify_sources.py`: Verified all 6 original template SHA-256 hashes byte-for-byte against `source-catalog.json`.
- `python -m unittest discover -v`: **86/86 tests passing** cleanly across:
  - `test_package_safety.py` (8)
  - `test_build_package.py` (25, including 8 negative regressions)
  - `test_modelmaker.py` (33)
  - `test_pptmaker.py` (13)
  - `test_reportmaker.py` (7)
- Manifest readiness: `ready: false`, pending gates: `["human_pack_approval", "render_review"]`.
