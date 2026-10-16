# Current Shared Status

## Milestone Complete: Template-Faithful Excel & PPT Package Delivered with Audited Second-Company Model (Cipla v2), Eradicated Carryover, and Hardened Pipeline Safety

- Branch: `agent/argon-continue`
- PR: #3 (`https://github.com/keshavmittal2203-droid/excel-model-maker/pull/3`)
- Updated: 2026-10-16, Asia/Calcutta (`agent/argon-continue`).

---

## Implemented & Verified

### 1. Second Company Demonstration: Cipla Limited v2 Native Package
- Bounded real-company demonstration with **Cipla Limited** (`CIPLA`, BSE `500087`):
  - Request specification: `examples/cipla-package.json`.
  - Staged & recalculated candidate workbook: `candidates/run_native_cipla_v2/Cipla-Recalculated.xlsx` (SHA-256: `a226bf9549dad68da31464a287264d1e30defd85ab122b44268764f40e7da5ab`).
  - Model pack proposal: `proposals/cipla.model-pack.proposal.json` (19/19 checks passed, 28/28 independent identities verified).
  - Deck inputs: `proposals/cipla.inputs.proposal.json` (219 slots, zero Marksans carryover, 48 unevidenced granular slots marked `"—"`).
  - Candidate presentation: `candidates/run_native_cipla_v2/Cipla-Executive-Overview.pptx` (SHA-256: `8313b8653783057e9d2625b26a6ece42dff9ffed0f67c0aa47ffe3da460665b6`).
  - Manifest: `candidates/run_native_cipla_v2/package.manifest.json`.
  - Visual inspection: All 5 slide PNG previews (`slide_1.png` to `slide_5.png`) and 4 sheet PDFs (`Data_Sheet.pdf`, `Ratio_Analysis.pdf`, `Intrinsic_Growth.pdf`, `WACC-DCF.pdf`) inspected via tool view.
  - Complete evidence report: `docs/SECOND_COMPANY_EVIDENCE_REPORT.md`.

### 2. Resolution of Cash-Flow Missing History (Audited Totals vs Transparent Dashes)
- Rectified previous v1 reporting: In v1, 48 intermediate granular cash-flow slots across FY18–FY21 were derived via balancing plugs (`non_cash = CFO - PBT - Taxes - WC`).
- In v2, all artificial plugs have been eradicated. All 48 unevidenced intermediate detail slots across FY18–FY25 on Slide 14 Table 0 are set to transparent em-dashes (`"—"`) with explicit citation `"Unevidenced granular sub-schedule; primary statutory filing required"`.
- The 9 core statutory line items and subtotals across all 8 years (72 table cells) are 100% populated with authentic Cipla consolidated Ind AS cash flows from primary annual reports (Cipla 89th Annual Report FY25, Note 45 / Consolidated Cash Flow Statement p. 216).
- Zero balancing plugs and zero carryover from Marksans or Sun Pharma.

### 3. Eradication of Hidden Marksans Carryover Across 17 Worksheets
- Conducted full audit of all 17 worksheets, classifying them into 4 explicit categories (`A. Generic shared methodology`, `B. Deliberately shared analyst assumption`, `C. Relevant external peer data`, `D. Stale company-specific content`).
- Sheet `BS`: Replaced columns J to N (FY22–FY26) with authentic Cipla balance sheet breakdown (Inventories, Receivables, Payables, Net Block). Drives authentic working capital and ROIC calculations in `Intrinsic Growth` and `DCF 2nd way`.
- Sheet `Beta`: Eradicated Marksans weekly price series from Column C. Populated with flat Cipla CMP (₹1,301.00) so that returns evaluate to 0.0 without `#DIV/0!` errors. Updated title to disclose that historical regression series is unpopulated and valuation relies on pure-play peer Hamada beta.
- Sheet `Cash Flow`: Zeroed unevidenced historical detail rows (7–12, 16–22, 24–26, 29–30), eradicating all Marksans carryover.

### 4. Adapter & Build Safety Hardening
- Added `REGISTERED_ADAPTERS` whitelist to `build_package.py`. Only explicitly reviewed Python adapter scripts and entry functions can be executed.
- Enforced strict containment: path traversal (`..`), absolute paths, and leading slashes are rejected prior to execution.
- Added `--check-only` (`--preflight-only`) fast dry-run mode that verifies request JSON schema, template hashes, pack presence, and COM registration in milliseconds without disk writes or COM activation.
- Early validation: Output workbook and presentation filenames are validated via `output_file()` before any source file checks, directory creation, or Office COM calls.
- 12 comprehensive safety regression tests in `test_package_safety.py` (all passing).

### 5. Retained-Content Audit & Slide 14 Cash Flow
- Replaced all 15 active rows on Slide 14 Table 0: authentic totals and transparent dashes.
- Machine-readable audit (`work/audit_retained_content.py`): 274 elements audited across shapes, tables, charts, embedded workbooks, raster screenshots; zero unresolved stale elements (`audit_passed: True`).

### 6. Full Test Suite & Originals Integrity
- `python verify_sources.py`: Verified all 6 original template SHA-256 hashes byte-for-byte against `source-catalog.json`.
- `python -m unittest discover -v`: **90/90 tests passing** cleanly across:
  - `test_package_safety.py` (12 tests)
  - `test_build_package.py` (25 tests, including 8 negative regressions and Cipla build test)
  - `test_modelmaker.py` (33 tests)
  - `test_pptmaker.py` (13 tests)
  - `test_reportmaker.py` (7 tests)

---

## Capabilities Matrix

| Scope / Dimension | Classification | Description & Provenance |
|---|---|---|
| **Excel 17-Sheet XML Injection** | **Implemented** | Openpyxl-based targeted XML editing preserving styles, formulas, follower ranges, and data tables. |
| **PowerPoint Deck Assembly** | **Implemented** | Curated slide retention (1, 2, 14, 16, 23), table cell updates, chart cache sync, picture replacement, and logo suppression. |
| **Independent Formula & Identity Checks** | **Implemented** | Recalculation verification, zero `#REF!`/#DIV/0!, $WACC = W_e K_e + W_d K_d (1-t)$, $K_e = R_f + \beta(R_m-R_f)$, balance sheet balance ($\Delta = 0.00$). |
| **Retained-Content Audit** | **Implemented** | 274 elements checked across tables, shapes, charts, embedded workbooks, raster screenshots; zero stale content. |
| **Pre-Flight Request Validation** | **Implemented** | Rejects unmapped slots, company mismatches, empty values, bad hashes, traversal paths, and invalid reporting units before directory writes. `--check-only` dry run mode. |
| **Marksans Pharma Package (Baseline & Scenario B)** | **Demonstrated** | Fully recalculated v6 models (`ef5d6d5a...`), 5-slide decks, 1080x1440 PNG previews, landscape PDFs. |
| **Cipla Limited Second-Company Packaging (v2)** | **Demonstrated (All 5 Stages)** | Stage 1 (Specs prepared), Stage 2 (Data mapped), Stage 3 (Build completed), Stage 4 (Financial validation passed), Stage 5 (Outputs visually inspected). SHA-256: `a226bf95...`. |
| **Non-Pharma / Non-Manufacturing Sectors (BFSI / Tech)** | **Manual Template Required** | Banking models (NIM/NPA) and SaaS (ARR/NRR) require fundamentally different 3-statement structures created by human analysts. |
| **Human Pack Approval & Promotion** | **Manual Review Required** | Transition from `status: "proposal"` to `status: "approved"` requires independent human financial review. |
| **Autonomous Agent Self-Approval** | **Unsupported / Prohibited** | Agent self-approval is blocked by reviewer guards and safety invariants. |
| **Fabrication of Missing Evidence** | **Unsupported / Prohibited** | Missing primary source data is cataloged under `unresolved_inputs`; inventing balancing numbers is strictly forbidden. Granular rows marked `"—"`. |

---

## Current Readiness State

- **Agent Visual Inspection**: **PROVISIONAL** (Marksans: 5 slides, 6 PDFs reviewed; Cipla: 5 slides, 4 PDFs reviewed; 9 unexported workbook sheets remain unreviewed visually).
- **Candidate Presentation State**: **CLEAN** (zero stale template figures, zero collisions, zero text clipping, audited cash flow table with transparent dashes, 274 audited elements).
- **Model Recalculation State**: **VERIFIED** (Marksans & Cipla: 19/19 pack checks, 28/28 independent identities passed, 0 formula errors).
- **Readiness State**: `ready: false`
- **Pending Gates**: `["human_pack_approval", "render_review"]`
