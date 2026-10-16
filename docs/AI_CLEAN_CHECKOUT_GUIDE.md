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
1. `originals/Marksans Keshav.xlsx` (`5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`)
2. `originals/ERR-4.pptx` (`9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba`)
3. `originals/GMCC 2023 (Bharat, Kanishk).pptx` (`bd6f3920bf32986f6930de011065ebba2c9dd45aba3c8416b6803d9a415f854f`)
4. `originals/GMCC 2023 (Bhavika, Falak).pptx` (`f19c7335e6514ef60a8109d2bf0021ba46f1752ae9be0ad3ad94ab5cd48ae331`)
5. `originals/GMCC 2023 (Pratiksha, Prakhar).pptx` (`c3b8e389ae2b3cfec1731a795efe70df310b962e45487243af2a1499c8537592`)
6. `originals/Healthcare and Pharma Industry Primer Final.pptx` (`0c2c84cea14ef51af7ad15a32dfc8c10b8e0f6acca68ff0ef50bfa6faa343f32`)

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

### Supported Inputs Workflow
- **Baseline Build**: Baseline model input values are declared directly inside `proposals/marksans.model-pack.proposal.json`; presentation text and chart slots are supplied in `proposals/err4.inputs.proposal.json`. No external inputs file is required.
- **Scenario Build**: Custom or stressed assumptions supply an external inputs file mapped via `workbook.inputs`, e.g., `proposals/scenario-b.inputs.json` (as seen in `examples/scenario-b-package.json`).
- *Note*: There is no `proposals/marksans.inputs.proposal.json`; baseline inputs reside directly within the model pack proposal.

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

## 11. Reusability Boundaries: Marksans vs New Company vs New Sector

| Capability Scope | Automation Status | Requirements & Limitations |
|---|---|---|
| **Marksans Re-estimation (Scenario Analysis)** | **Fully Automated** | Supply new inputs (e.g. `examples/scenario-b-package.json`) with modified $R_f$, $K_d$, or growth assumptions. Runs end-to-end. |
| **New Company in Pharma Sector** | **Semi-Automated** | Requires preparing new `company.model-pack.json` matching the 17-sheet pharma model structure and new `company.inputs.json` mapped from audited filings. |
| **New Company in Different Sector (e.g. BFSI / Tech)** | **Manual Template Required** | The 17-sheet Marksans template is designed for manufacturing/pharma. BFSI (banks/NBFCs) require Net Interest Margin / loan book structures, and SaaS requires ARR/NRR schedules. Deriving a new sector model template requires human financial analyst modeling. |
| **Other Uploaded PPT Decks** | **Partially Automated** | Slide inspection and layout mapping supported via `template_atlas.py`. New deck packs must be declared in `proposals/` and reviewed before production use. |

---

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
