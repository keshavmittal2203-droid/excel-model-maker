# Canonical AI Clean-Checkout & Workflow Guide

A provider-neutral, reproducible guide for any autonomous agent or human engineer starting from a clean git clone of `https://github.com/keshavmittal2203-droid/excel-model-maker`.

---

## 1. Environment Prerequisites & Supported Operating Systems

### Supported Environments
- **Core Engine (Linux, macOS, Windows)**:
  - Python 3.10+ (tested through Python 3.14 on Windows).
  - Openpyxl, Pillow, jsonschema.
  - Generates candidate `.xlsx` and `.pptx` files, validates formulas and shared-formula followers, verifies hashes, and executes unit tests.
- **Native COM Rendering (Windows Only)**:
  - Microsoft Office 16.0+ (Excel & PowerPoint desktop applications).
  - `pywin32` (`pythoncom` / `win32com`) and `winreg`.
  - Recalculates candidate workbooks via `CalculateFullRebuild()` without macros or external link prompts.
  - Exports slide PNG previews (1080x1440) and multi-column landscape schedule PDFs (`FitToPagesWide = 1`).
- **Platform Limitations & `--skip-com`**:
  - `pywin32` is a Windows-only library. On Linux and macOS, install only `openpyxl Pillow jsonschema`.
  - Passing `--skip-com` skips PowerPoint slide PNG exports and Excel sheet PDF exports during package builds.
  - *Important*: `--skip-com` does **not** bypass the Windows Excel COM recalculation requirement in `validate_workbook.py` when recalculating a fresh workbook from scratch. Full recalculation validation requires a Windows environment with Excel installed.

### Setup Commands
```bash
# Clone the repository
git clone https://github.com/keshavmittal2203-droid/excel-model-maker.git
cd excel-model-maker

# Optional: Create and activate virtual environment
python -m venv .venv
# On Windows PowerShell:
.venv\Scripts\Activate.ps1
# On Linux/macOS:
source .venv/bin/activate

# Install dependencies:
# On Windows:
pip install openpyxl Pillow jsonschema pywin32
# On Linux / macOS (pywin32 is omitted):
pip install openpyxl Pillow jsonschema
```

---

## 2. Source Reconstruction & Hash Verification

The repository stores six canonical original financial models and presentation templates in `originals/`. Large presentations are stored as split chunks to adhere to GitHub file size limits and are reconstructed deterministically.

Execute source verification before touching any model or presentation:
```bash
python verify_sources.py
```
This script checks and reconstructs all six templates, verifying their SHA-256 hashes byte-for-byte against `source-catalog.json`:
1. `originals/Marksans Keshav.xlsx` (`4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`)
2. `originals/archive/Marksans Keshav.xlsx` (`5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`)
3. `originals/ERR-4.pptx` (`9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba`)
4. `originals/GMCC 2023 (Bharat, Kanishk).pptx` (`bd6f3920bf32986f6930de011065ebba2c9dd45aba3c8416b6803d9a415f854f`)
5. `originals/GMCC 2023 (Bhavika, Falak).pptx` (`f19c7335e6514ef60a8109d2bf0021ba46f1752ae9be0ad3ad94ab5cd48ae331`)
6. `originals/GMCC 2023 (Pratiksha, Prakhar).pptx` (`c3b8e389ae2b3cfec1731a795efe70df310b962e45487243af2a1499c8537592`)
7. `originals/Healthcare and Pharma Industry Primer Final.pptx` (`0c2c84cea14ef51af7ad15a32dfc8c10b8e0f6acca68ff0ef50bfa6faa343f32`)

---

## 3. Template Selection & Atlas Navigation

To select an uploaded design template, consult `template-workflow/atlas/INDEX.md` or execute:
```bash
python template_atlas.py
```
This inspects the 5 presentation templates (182 total slides) and indexes available slide layouts:
- **ERR-4.pptx**: Institutional equity research format (dark navy palette, financial tables, ratio charts, WACC schedule).
- **GMCC 2023 Decks (Teams 5, 10, 28)**: Valuation case competition pitch decks (market dynamics, strategic catalysts, football field valuation).
- **Healthcare & Pharma Industry Primer**: Comprehensive macroeconomic sector overview deck.

---

## 4. Supported Cells, Shapes & Layout Engine Capabilities

### Excel Engine (`modelmaker.py`)
- **Template-Preserving XML Transformation**: Edits only cells explicitly declared in a versioned model pack. Preserves all other worksheets across the 17-sheet model, custom formatting, styles, grid lines, and formulas.
- **Formula Integrity**: Rejects arbitrary formula tampering. Validates shared-formula ranges, follower cell offsets, and data table attributes (`{=TABLE()}`).
- **Cell Mapping Discovery**: Quote paths containing spaces in PowerShell:
  ```powershell
  python modelmaker.py inspect 'originals/Marksans Keshav.xlsx'
  ```

### PowerPoint Engine (`pptmaker.py`)
- **Slide Retention & Pruning**: Assembles curated decks (`retained_slides`) by re-ordering slide IDs, pruning unreferenced slide parts from the archive, and updating `[Content_Types].xml`.
- **Shape Text & Tokens**: Replaces text while preserving `<a:rPr>` font properties (typeface, size, color, kerning).
- **Tables**: Updates individual table cells (`<a:tc>`) while maintaining borders and shading.
- **Native Charts**: Synchronously updates chart series, categories, `<c:numCache>`, `<c:strCache>`, and embedded Excel worksheets (`Microsoft_Excel_Worksheet.xlsx`).
- **Bounded Image Replacement**: Replaces embedded pictures (`<p:pic>`) with run-local rendered financial schedules.
- **Shape Suppression**: Suppresses shapes (e.g. unverified logos). When given a `replacement_message`, inserts a styled template-faithful card (e.g. unavailable geographic data notice).

---

## 5. Supplying Company Data (Sources, Dates & Units)

Every input supplied to `modelmaker.py` or `pptmaker.py` requires:
1. `value`: Numeric float/int for quantitative model inputs; exact string for text slots.
2. `source`: Traceable provenance locator (e.g., `Audited Annual Report FY25, Note 18, page 94` or `Marksans-Recalculated.xlsx!WACC-DCF!E34`).
3. `unit`: Mandatory unit token (`INR_Cr`, `Percent`, `Ratio`, `Multiple`, `Count`, `Days`, `PerShare`).
4. `period`: Exact fiscal period (`FY18` to `FY25` for historical, `FY26E` to `FY30E` for forecast).

Missing or empty source references will trigger actionable validation errors (`DeckError` / `ModelError`).

---

## 6. Distinguishing Historical Facts from Forecast Assumptions

Never mix historical facts with subjective projections:
- **Historical Financial Facts (`inherited_template_value` or `externally_verified_fact`)**:
  - Financial cutoff date: Audited FY25 ending `2025-03-31`.
  - Historical turnover, EBITDA, debt book value, and balance sheet items are anchored to audited actuals.
- **Market Data Snapshot (`inherited_template_value` / `calculated_estimate`)**:
  - CMP ₹309.10 and Market Cap ₹14,002.74 Cr captured on Screener.in retrieval date `2025-10-05` (Sunday).
  - Friday closing session was `2025-10-03`, but primary exchange trade confirmations are absent from `originals/`; document this limitation.
- **Valuation Assumptions (`analyst_assumption` / `calculated_estimate`)**:
  - Risk-free rate ($R_f=7.30\%$): 10-Year G-Sec benchmark.
  - Marginal tax rate ($30.00\%$): Indian statutory corporate tax assumption.
  - Pre-tax cost of debt ($8.00\%$): Forward marginal replacement borrowing rate.
  - Peer median capital structure ($2.24\%$ debt / $97.76\%$ equity): Institutional Damodaran pure-play modeling assumption.
  - Intrinsic growth / reinvestment rate: Projected from sustainable growth formulas ($g = \text{ROE} \times b$).

---

## 7. Preparing & Validating a Request

A package build request is defined in a JSON file conforming to `package-request.schema.json`. Shipped examples:
- Baseline Marksans: `examples/marksans-package.json`
- Stressed Scenario B: `examples/scenario-b-package.json`

### Baseline Request Format (`examples/marksans-package.json`)
The request format uses nested `workbook`, `presentation`, and `exports` objects:
```json
{
  "$schema": "../package-request.schema.json",
  "schema_version": 1,
  "package_id": "marksans-curated-overview-2025",
  "name": "Marksans Pharma Financial Model & 5-Slide Executive Overview",
  "description": "Repeatable template-driven package generating reconciled Excel financial model and 5-slide PowerPoint deck from originals/ERR-4.pptx.",
  "workbook": {
    "source": "originals/Marksans Keshav.xlsx",
    "source_sha256": "5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a",
    "model_pack": "proposals/marksans.model-pack.proposal.json",
    "output_name": "Marksans-Recalculated.xlsx",
    "expected_recalculated_sha256": "9c1fa6e39d88d3bb597de8e4cf6208170471751cb87580aa7c8b6255dd6054e8"
  },
  "presentation": {
    "template": "originals/ERR-4.pptx",
    "template_sha256": "9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba",
    "deck_pack": "proposals/err4.deck-pack.proposal.json",
    "deck_inputs": "proposals/err4.inputs.proposal.json",
    "output_name": "Marksans-Pharma-Overview.pptx"
  },
  "exports": {
    "export_slide_pngs": true,
    "png_width": 1080,
    "png_height": 1440,
    "export_sheet_pdfs": true,
    "pdf_sheets": [
      "Data Sheet",
      "Ratio Analysis",
      "Intrinsic Growth",
      "WACC-DCF",
      "DCF 1st way",
      "DCF 2nd way"
    ],
    "pdf_orientation": "landscape",
    "pdf_fit_to_page_wide": 1
  }
}
```

### Supported Inputs Workflow & Request Examples
- **Baseline Build (`examples/marksans-package.json`)**: Baseline model input values are declared directly inside `proposals/marksans.model-pack.proposal.json`; presentation text and chart slots are supplied in `proposals/err4.inputs.proposal.json`. No external inputs file is required.
- **Scenario Build (`examples/scenario-b-package.json`)**: Custom or stressed assumptions supply an external inputs file mapped via `workbook.inputs`, e.g., `proposals/scenario-b.inputs.json`.
- **Second Company Demonstration (`examples/cipla-package.json`)**: Full company-request specification demonstrating reuse on a second real company (**Cipla Limited**, BSE `500087` / NSE `CIPLA`), mapped to `proposals/cipla.inputs.proposal.json` and documented in `docs/SECOND_COMPANY_EVIDENCE_REPORT.md`.

### Canonical Company Specification (`package-request.schema.json`)
The request schema enforces formal company metadata and segregated source provenance:
```json
"company": {
  "name": "Cipla Limited",
  "ticker": "CIPLA",
  "bse_code": "500087",
  "exchange": "BSE / NSE",
  "valuation_date": "2025-10-05",
  "historical_cutoff": "FY2025 (2025-03-31)",
  "currency": "INR",
  "reporting_unit": "INR_Cr",
  "fiscal_year_convention": "April-March",
  "source_documents": [
    {
      "title": "Cipla Limited 89th Annual Report FY 2024-25",
      "type": "Annual Report",
      "period": "FY25",
      "basis": "consolidated"
    }
  ],
  "source_data": {
    "user_supplied_facts": { ... },
    "source_extracted_facts": { ... },
    "analyst_assumptions": { ... },
    "calculated_outputs": { ... },
    "unresolved_inputs": { ... }
  }
}
```

### Pre-Flight Request Validation & Non-Fabrication Policy
Before creating directories or activating Office COM, `validate_package_request` runs strict pre-flight checks:
1. **Schema & Integrity**: Verifies `schema_version == 1` and byte-exact SHA-256 hashes of all source models and presentation templates.
2. **Mandatory Company Metadata**: Validates identity, valuation date, cutoff date, currency, reporting unit, and fiscal year convention.
3. **Complete Slot Mapping**: Rejects the build if any deck slot is missing from supplied values without a declared workbook binding. Missing inputs **never silently become zero or retain another company's data**.
4. **Strict Non-Fabrication Policy**: Where primary source evidence is absent (such as older FY18–FY21 cash flow schedule breakdowns in Cipla's FY25 annual report), data points are explicitly documented under `unresolved_inputs` in the evidence report rather than manufacturing synthetic balancing figures.

---

## 8. Building into a Fresh Output Directory

### Strict Safety Invariants
1. **Fresh Directories Required**: Target output directory must NOT exist or must be completely empty. Never overwrite existing output runs.
2. **Path Safety**: Traversal paths (`../`), absolute paths, Windows device names (`CON`, `PRN`), and alternate data streams (`:stream`) are rejected before any write.
3. **Office COM Safety**: Dedicated `DispatchEx` instance is used; macros and link updates are disabled (`AutomationSecurity = 3`, `AskToUpdateLinks = False`). PowerPoint closes only its owned presentation without calling application `Quit()`.

### Execution Commands (PowerShell / Bash)
```powershell
# Build Baseline package (requires fresh non-empty directory)
python build_package.py --request examples/marksans-package.json --output-dir candidates/run_native_baseline_v4 --allow-proposal

# Build Stressed Scenario B package
python build_package.py --request examples/scenario-b-package.json --output-dir candidates/run_native_scenario_b_v4 --allow-proposal

# Build the Four Priority Multi-Sector Target Packages (Workbook + 5-Slide Deck + 6 Sheet PDFs + 5 Slide PNGs)
python build_package.py --request examples/mm-package.json --output-dir candidates/run_verify_mm --allow-proposal
python build_package.py --request examples/ltm-package.json --output-dir candidates/run_verify_ltm --allow-proposal
python build_package.py --request examples/clean-package.json --output-dir candidates/run_verify_clean --allow-proposal
python build_package.py --request examples/edelweiss-package.json --output-dir candidates/run_verify_edelweiss --allow-proposal
```

---

## 9. Recalculating, Inspecting & Reviewing Outputs

### Recalculation & Formula Checks
The orchestrator executes native Excel full recalculation (`CalculateFullRebuild()`) and runs independent formula validation:
- Zero Excel error values (`#REF!`, `#DIV/0!`, `#VALUE!`).
- Mathematical identity check: $WACC = W_e \times K_e + W_d \times K_d \times (1 - t)$.
- Mathematical identity check: $K_e = R_f + \beta \times (R_m - R_f)$.
- Mathematical identity check: Target weights sum to $100\%$.
- Identity check: Balance sheet assets balance liabilities across all historical years.

### Visual Render Review
Inspect the generated previews inside `<output-dir>/previews/`:
- `slide_1.png` to `slide_5.png`: High-resolution slide renders.
- `slide_contact_sheet.png`: 2x3 composite slide contact sheet.
- `Data_Sheet.pdf`, `WACC-DCF.pdf`, etc.: Multi-column landscape schedule PDFs.

Check:
1. Absence of text clipping or overlapping text boxes.
2. Verified suppression of unverified corporate logos.
3. Replacement message box for geographic breakdown (`"Geographic revenue breakdown unavailable — primary filing required"`).
4. Table numbers matching recalculated workbook schedules.

---

## 10. Output Artifact Inventory & File Counts

Each package run generates **18 total files on disk**, categorized as follows:

### 13 Primary Output Artifacts (Cataloged in `package.manifest.json`)
1. `Marksans-Recalculated.xlsx`: Fully recalculated candidate financial model.
2. `Marksans-Pharma-Overview.pptx`: Curated 5-slide presentation deck.
3. `previews/slide_1.png` to `slide_5.png`: 5 presentation slide previews.
4. `previews/*.pdf`: 6 key worksheet landscape schedules (`Data Sheet`, `Ratio Analysis`, `Intrinsic Growth`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`).

### 5 Auxiliary & Metadata Files (On Disk)
1. `package.manifest.json`: Cryptographic manifest recording all input and output SHA-256 hashes, COM capabilities, and financial tie-outs (excluded from self-hash to avoid circular dependencies).
2. `previews/slide_contact_sheet.png`: 2x3 composite contact sheet for visual review.
3. `assets/image22_marksans.png`, `image23_marksans.png`, `image24_marksans.png`: Run-local high-resolution WACC schedule replacement table images embedded into Slide 5.

---

## 11. Reusability Boundaries: Demonstrated Multi-Sector Packages

The repository has demonstrated multi-sector financial modeling across six distinct corporate architectures:

| Company / Target | Sector & Business Model | Modeling Architecture & Valuation | Verification & Output Status |
|---|---|---|---|
| **Marksans Pharma** (`examples/marksans-package.json`) | Pharmaceuticals Formulations | 17-sheet integrated financial model; WACC-DCF, 5-yr FCFF, 10-yr fundamental ROIC DCF, Comps. | Fully recalculated candidate models (Baseline & Scenario B), 5-slide decks, 1080x1440 PNG previews, 6 landscape PDFs. |
| **Cipla Limited** (`examples/cipla-package.json`) | Large-cap Global Pharma | 17-sheet adapted financial model; 10-yr Ind AS statutory statements, pure-play Hamada beta, illustrative DCF models. | Fully recalculated candidate model (`candidates/run_native_cipla_v3/`), 5-slide deck, 5 PNG previews, 4 landscape PDFs. |
| **M&M** (`companies/MM/`, `examples/mm-package.json`) | Automotive & Farm Equipment | 6-sheet model: Segment separation (Auto vs Farm), standalone net debt isolation, SOTP valuation (16.0x Auto + 14.0x Farm + MMFS 52.2% & TechM 25.5% stakes with 20% HoldCo discount = **₹2,929.98/sh**). | Recalculated via Excel COM (`MM-Recalculated.xlsx`), $\Delta = 0.00$ Balance Sheet tie across 12 periods, 6 landscape schedule PDFs, 5-slide workbook-linked deck (`MM-Executive-Overview.pptx`), 5 slide PNGs. |
| **LTM** (`companies/LTM/`, `examples/ltm-package.json`) | IT Services & Digital Solutions | 6-sheet model: Post-merger amalgamation history (LTI & Mindtree), USD revenue ($4,492M), connected 5Y USD CC growth * FX forecast, headcount (84,307), utilization (86.5%), CAPM WACC (11.50%), 5.0% TGR = **₹2,813.51/sh**. | Recalculated via Excel COM (`LTM-Recalculated.xlsx`), $\Delta = 0.00$ Balance Sheet tie across 12 periods, 6 landscape schedule PDFs, 5-slide workbook-linked deck (`LTM-Executive-Overview.pptx`), 5 slide PNGs. |
| **CLEAN** (`companies/CLEAN/`, `examples/clean-package.json`) | Specialty Chemicals Manufacturing | 6-sheet model: Parent Kurkumbh Units 1–3 (29,900 MTPA) + Subsidiary CFCL Unit 4 HALS capacity (10,500 $\to$ 15,000 MTPA) and utilization ramp (55% $\to$ 90%), CAPM WACC (12.00%), 5.0% TGR = **₹415.69/sh**. | Recalculated via Excel COM (`CLEAN-Recalculated.xlsx`), $\Delta = 0.00$ Balance Sheet tie across 12 periods, 6 landscape schedule PDFs, 5-slide workbook-linked deck (`CLEAN-Executive-Overview.pptx`), 5 slide PNGs. |
| **EDELWEISS** (`companies/EDELWEISS/`, `examples/edelweiss-package.json`) | Financial Services Holding Company | 6-sheet model: Financial Services SOTP (EAAA 22.0x PAT, EARCL 1.2x Book, Mutual Fund 4% AAUM, Insurance 1.0x EV, Credit 0.8x Book) less HoldCo net debt = **₹133.54/sh**. **Industrial FCFF DCF strictly withheld and blocked**. | Recalculated via Excel COM (`EDELWEISS-Recalculated.xlsx`), $\Delta = 0.00$ Balance Sheet tie across 12 periods, 6 landscape schedule PDFs, 5-slide workbook-linked deck (`EDELWEISS-Executive-Overview.pptx`), 5 slide PNGs. |

### Inspecting and Testing Target Models
Execute the dedicated regression test suite:
```bash
python -m unittest test_four_targets.py -v
```
All four company workbooks and presentations are validated for zero formula errors, zero missing caches, exact $\Delta = 0.00$ balance sheet ties across all 12 periods, connected 5-year forecasts, zero historical BS plugs, independent 25-cell sensitivity matrices, source registry metadata completeness, 24 landscape PDF previews, 4 workbook-linked `.pptx` presentations, 20 slide PNG previews, Edelweiss FCFF blockage, M&M SOTP double-counting prevention, Clean Science capacity distinctions, and LTM constant-currency/merger handling.


## 12. Financial Modelling Course Reference (Methodology Alignment)

- **Course Playlist**: [Learn Financial Modelling - Step by Step (The Valuation School)](https://www.youtube.com/playlist?list=PL3uUjzLk6PulhRop_ffNeHyK0kprzO4cT)
- **Methodology Documentation**: See `docs/COURSE_REFERENCE.md` for session-by-session mapping.

### Course Evidence Tiers
1. **Reviewed & Harvested Metadata**: All 34 lessons cataloged in `docs/course_lessons.json`.
2. **Inferred Structural Mapping to 17-Sheet Workbook**:
   - Sessions 4, 11 $\to$ 3-Statement financial statements (`Data Sheet`, `BS`, `Cash Flow`, `Main Financials`).
   - Session 5, 18–20 $\to$ Financial ratio analysis & DuPont decomposition (`Ratio Analysis`).
   - Session 6, 12 $\to$ Top-line forecasting & intrinsic growth schedules (`Intrinsic Growth`).
   - Sessions 7, 13, 16 $\to$ Unlevered DCF modelling & sensitivity data tables (`DCF 1st way`, `DCF 2nd way`).
   - Sessions 8–10 $\to$ WACC, CAPM, pure-play beta unlevering/re-levering (`WACC-DCF`, `Beta`, `RM`, `peerdata`).
   - Sessions 14–15 $\to$ Comparable company relative valuation (`Comps`, `compdata`, `peerdata`).
3. **Transcript Status Note**: YouTube watch pages do not supply static transcript text tracks via plain HTTP requests. Granular spoken rules must be verified from local transcripts before being marked as transcript-verified.

---

## 13. Gate Semantics & Approval Governance

The repository strictly enforces a multi-tier governance model separating automated generation from human certification:

### Status & Gate Invariants
- **Unapproved Proposal State**:
  - Model packs in `proposals/` and deck packs in `proposals/` carry `status: "proposal"`.
  - Built package manifests declare `ready: false` with pending gates:
    ```json
    "pending_gates": [
      "human_pack_approval",
      "render_review"
    ]
    ```
- **Prohibition on Autonomous Self-Approval**:
  - Autonomous agents MUST NOT self-approve proposals or remove pending gates.
  - Reviewer guards reject disallowed tokens (`agent`, `ai`, `bot`, `gemini`, `codex`, `claude`, `argon`, `auto`) in approval fields.
  - Candidate builds require the explicit `--allow-proposal` CLI flag to evaluate unapproved proposal packs.
- **Human Approval Transition**:
  - Only human reviewers can approve proposals by updating the pack metadata (`status: "approved"`, `reviewed_by: "<human_reviewer>"`) and graduating files out of `proposals/`.

---

## 14. Paste-Ready Prompt for Another AI Assistant

Copy and paste the prompt below into another coding assistant (such as Claude, Codex, GPT-4, or Gemini) to operate this repository cleanly without breaking safety invariants:

```text
You are pair programming in https://github.com/keshavmittal2203-droid/excel-model-maker on your assigned branch.

Your core operating rules:
1. IMMUTABLE ORIGINALS & TEMPLATES: Never edit, overwrite, or delete files in originals/. Run 'python verify_sources.py' before any work to confirm all 6 SHA-256 hashes match source-catalog.json.
2. DIRECTORY SAFETY: Output directories must be brand new, fresh, and empty. Never overwrite existing output directories or unlink candidate files. Path traversal (../), absolute paths, and Windows device names are strictly prohibited.
3. GOVERNANCE & GATES: Candidate manifests must maintain ready: false and pending_gates: ["human_pack_approval", "render_review"]. Do not self-approve packs or remove pending gates. Use --allow-proposal for testing unapproved candidates.
4. FINANCIAL RIGOR & NON-FABRICATION: Never invent balancing numbers, fictitious historical rows, or placeholder zeros. Sourced data must be segregated into user_supplied_facts, source_extracted_facts, analyst_assumptions, calculated_outputs, and unresolved_inputs. Missing inputs must be cataloged in unresolved_inputs with explicit citations.
5. COM DISPATCH SAFETY: Never call application Quit() on PowerPoint. Dedicated Excel DispatchEx instances must be used with macros and link updates disabled.
6. TESTING: All tests must pass cleanly. Run 'python -m unittest discover' before committing. Do not force-push or merge main.

To build a verified package:
1. Check sources: python verify_sources.py
2. Run tests: python -m unittest discover
3. Build candidate: python build_package.py --request examples/marksans-package.json --output-dir candidates/run_<version> --allow-proposal
4. Audit candidate: python work/audit_retained_content.py --presentation candidates/run_<version>/Marksans-Pharma-Overview.pptx --workbook candidates/run_<version>/Marksans-Recalculated.xlsx --template originals/ERR-4.pptx
5. Inspect previews in candidates/run_<version>/previews/
```
