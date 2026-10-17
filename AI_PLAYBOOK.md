# Universal AI Playbook: Excel Financial Model & PowerPoint Deck Engine

This playbook gives any AI agent deterministic, step-by-step instructions to take researched company, sector, or case-study data and produce:
1. A recalculated, error-free, formula-verified **Excel financial model (`.xlsx`)** via `modelmaker.py` and `validate_workbook.py`.
2. A template-preserving, provenance-checked **PowerPoint presentation deck (`.pptx`)** via `pptmaker.py`.

---

## 0. Non-Negotiable Safety & Provenance Rules

1. **Never modify `originals/`**: Treat every `.xlsx`, `.pptx`, and `.chunk*` file in `originals/` as immutable read-only reference material. Always write outputs to `candidates/`.
2. **Verify source hashes first**: Run `python verify_sources.py` to reconstruct chunked `.pptx` templates and verify all 6 original SHA-256 digests against `source-catalog.json`.
3. **Explicit units, bounds, and citations**: Every number injected into an Excel model (`inputs.json`) or PowerPoint deck (`err4.inputs.json`) must include a non-empty `source` citation and satisfy the pack's `minimum` / `maximum` bounds and precision limits.
4. **Never declare an `.xlsx` model `validated` without recalculation**: `modelmaker.py build` invalidates cached formula values (`<v>`) and emits `"status": "needs_recalculation"`. Only `modelmaker.py validate --recalculate` (or `validate_workbook.py`) can transition a workbook to `"status": "validated"` after executing actual formula recalculation, scanning all cells for Excel errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`), and verifying 29 independent accounting/valuation identities.

---

## 1. Template Catalog

| Template File | Type | SHA-256 | Structure & Use Case |
| :--- | :--- | :--- | :--- |
| `originals/Marksans Keshav.xlsx` | `.xlsx` | `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a` | 17-sheet institutional Equity Research & Valuation workbook (`Data Sheet`, `BS`, `Cash Flow`, `Main Financials`, `Ratio Analysis`, `Intrinsic Growth`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`, `Comps`, `compdata`, `peerdata`, `sectordata`, `Beta`, `RM`). Reference pack: `marksans.model-pack.json` (93 inputs, 19 checks). |
| `originals/ERR-4.pptx` | `.pptx` | `9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba` | 25-slide Equity Research Report deck (Company Overview, Industry Overview, Porter's 5 Forces, SWOT, Quarterly/P&L/BS/Cash Flow/Ratio Analysis tables, DuPont, Coffee Can, Warren Buffett $1 Test, DCF-WACC Valuation). Reference pack: `err4.deck-pack.json`. |
| `originals/GMCC 2023 (Bharat, Kanishk).pptx` | `.pptx` | `bd6f3920bf32986f6930de011065ebba2c9dd45aba3c8416b6803d9a415f854f` | 37-slide Global Microfinance Case Competition deck (Country/Company Analysis, Proposed Solutions, Key Financials, Implementation & SDG Impact). |
| `originals/GMCC 2023 (Bhavika, Falak).pptx` | `.pptx` | `f19c7335e6514ef60a8109d2bf0021ba46f1752ae9be0ad3ad94ab5cd48ae331` | 37-slide Strategy Case Deck (Socioeconomic & Microfinance Landscape, Short/Long-Term Strategy, Impact & Appendix). |
| `originals/GMCC 2023 (Pratiksha, Prakhar).pptx` | `.pptx` | `c3b8e389ae2b3cfec1731a795efe70df310b962e45487243af2a1499c8537592` | 45-slide Case Competition Deck (Bird's Eye View, Dual Obstacle Perspective, Fintech/Credit Scoring/Timeline/Risk Mitigations/Financials). |
| `originals/Healthcare and Pharma Industry Primer Final.pptx` | `.pptx` | `0c2c84cea14ef51af7ad15a32dfc8c10b8e0f6acca68ff0ef50bfa6faa343f32` | 38-slide Sector Primer Deck (Global & India Healthcare/Pharma Landscape, Value Chain, Regulatory Drivers, Financial Benchmarks). |

---

## 2. Step-by-Step Workflow for Any AI Agent

### Step 1: Verify Originals & Inspect Templates
```bash
python verify_sources.py
python modelmaker.py inspect "originals/Marksans Keshav.xlsx"
python pptmaker.py inspect "originals/ERR-4.pptx"
```

### Step 2: Populate Excel Model Inputs (`marksans.inputs.json` or Custom Company Inputs)
`marksans.model-pack.json` maps **93 editable numeric inputs** across all 10 input sheets of `originals/Marksans Keshav.xlsx`:
- **Company Meta (`Data Sheet!B7:B9`)**: `meta_face_value` (INR/share), `meta_current_price` (INR/share), `meta_market_cap` (INR Cr). Note: `Data Sheet!B6` (`Number of shares`) is a formula (`IF(B9>0, B9/B8, 0)`) and is automatically calculated.
- **5-Year Historical P&L & Balance Sheet (`Data Sheet!G17:K69`)**:
  - `hist_sales_y1` .. `hist_sales_y5` (`G17:K17`, INR Cr)
  - `hist_depreciation_y1` .. `hist_depreciation_y5` (`G26:K26`, INR Cr)
  - `hist_tax_y1` .. `hist_tax_y5` (`G29:K29`, INR Cr)
  - `hist_net_profit_y1` .. `hist_net_profit_y5` (`G30:K30`, INR Cr)
  - `hist_borrowings_y5` (`K59`, INR Cr) and `hist_cash_bank_y5` (`K69`, INR Cr)
- **Detailed Balance Sheet & Working Capital (`BS!M13:N38`)**:
  - `bs_inventories_y4`, `bs_inventories_y5` (`M33:N33`), `bs_receivables_y4`, `bs_receivables_y5` (`M34:N34`), `bs_payables_y4`, `bs_payables_y5` (`M13:N13`), `bs_total_liabilities_y5` (`N16`), `bs_total_assets_y5` (`N38`).
- **Cash Flow & CapEx (`Cash Flow!J15:N15`, `N5`)**:
  - `cf_cfo_y5` (`N5`), `cf_fixed_assets_purchased_y1` .. `cf_fixed_assets_purchased_y5` (`J15:N15`, negative INR Cr outflows).
- **Peer & Comparable Companies (`peerdata`, `compdata`, `sectordata`)**:
  - `peer1_mar_cap` .. `peer5_mar_cap` (`peerdata!E5:E9`) and `peer1_debt` .. `peer5_debt` (`peerdata!G5:G9`).
  - `comp_target_cmp`, `comp_target_shares`, `comp_target_debt`, `comp_target_cash`, `comp_target_sales_ttm`, `comp_target_np_ttm`, `comp_target_ev_ebitda` (`compdata` row 10 & row 25).
  - `sector_pharma_reinvestment_rate` (`sectordata!L29`), `sector_healthcare_reinvestment_rate` (`sectordata!L47`).
- **WACC & DCF Assumptions (`WACC-DCF`, `Intrinsic Growth`, `DCF 1st way`, `DCF 2nd way`)**:
  - Peer tax rates (`WACC-DCF!F11:F15`) and levered betas (`WACC-DCF!I11:I15`).
  - `pre_tax_cost_of_debt` (`WACC-DCF!E22`), `risk_free_rate` (`WACC-DCF!J22`), `wacc_current_total_debt` (`WACC-DCF!C30`), `wacc_current_market_cap` (`WACC-DCF!C31`).
  - `ig_marginal_tax_rate_y1` .. `y5` (`Intrinsic Growth!E47:I47`).
  - `dcf1_tax_rate_base` .. `y5` (`DCF 1st way!F7:K7`), `dcf1_reinvestment_rate_base` .. `y5` (`DCF 1st way!F11:K11`), `dcf1_expected_growth_rate` (`E19`), `dcf1_terminal_growth_rate` (`E20`), `dcf1_reference_market_price` (`E44`).
  - `dcf2_terminal_growth_rate` (`DCF 2nd way!D10`).

Every input in `inputs.json` must follow:
```json
{
  "values": {
    "risk_free_rate": {
      "value": 0.073,
      "source": "RBI 10Y G-Sec Par Yield Benchmark (Oct 2026)"
    }
  }
}
```

### Step 3: Build & Recalculate the Candidate `.xlsx` Workbook
```bash
python modelmaker.py build \
  --template "originals/Marksans Keshav.xlsx" \
  --pack marksans.model-pack.json \
  --inputs marksans.inputs.json \
  --output candidates/Marksans-Candidate.xlsx

python modelmaker.py validate candidates/Marksans-Candidate.xlsx \
  --pack marksans.model-pack.json \
  --recalculate \
  --recalculated-output candidates/Marksans-Recalculated.xlsx \
  --perturb-check \
  --template "originals/Marksans Keshav.xlsx" \
  --inputs marksans.inputs.json \
  --report candidates/Marksans-Validation-Report.json
```
Confirm that the validation output reports:
- `"status": "validated"` and `"ready": true`
- `"formula_error_count": 0` and `"uncalculated_formula_count": 0`
- `"independent_identities_passed": 29` out of `29`
- `"perturbation_and_restoration": { "passed": true, ... }`

### Step 4: Build the Candidate `.pptx` Presentation Bound to the Recalculated `.xlsx`
`pptmaker.py` supports three slot types in `deck-pack.json` (` deck-pack.schema.json`):
1. `"shape_text"`: Replaces the text of a shape (`shape_id`, optional `shape_name`, optional `paragraph_index`) while preserving `<a:pPr>` and `<a:rPr>` run formatting.
2. `"shape_token"`: Replaces an exact `token` substring inside a shape's text runs while preserving run-level formatting.
3. `"table_cell"`: Replaces the text of a specific `(row, col)` cell inside a table (`shape_id`, optional `shape_name`) while preserving `<a:tcPr>`, `<a:pPr>`, and `<a:rPr>`.

Any slot may include a `workbook_binding` block:
```json
"workbook_binding": {
  "sheet": "Data Sheet",
  "cell": "K28",
  "scale": 1.0,
  "decimals": 2,
  "abs_tolerance": 0.01
}
```
When `--workbook candidates/Marksans-Recalculated.xlsx` is supplied, `pptmaker.py` verifies that the bound cell in the recalculated `.xlsx` has no formula error, has a valid cached value, and matches the value injected into the slide within `abs_tolerance`.

Run:
```bash
python pptmaker.py build \
  --template "originals/ERR-4.pptx" \
  --pack err4.deck-pack.json \
  --inputs err4.inputs.json \
  --workbook candidates/Marksans-Recalculated.xlsx \
  --output candidates/ERR-4-Marksans-Candidate.pptx
```

### Step 5: Single-Command Multi-Sector Package Orchestration (`build_package.py`)
In addition to the 17-sheet pharmaceutical template (`originals/Marksans Keshav.xlsx`), the engine supports 6-sheet multi-sector financial models and workbook-linked 5-slide executive decks (`originals/ERR-4.pptx`, retained slides `[1, 2, 14, 16, 23]`) via registered Python workbook adapters (`work/models_<key>.py`) and `package-request.schema.json` (`schema_version: 1`):

```powershell
# Reproducible single-command builds (requires a fresh/non-existent --output-dir)
python build_package.py --request examples/mm-package.json --output-dir candidates/run_verify_mm --allow-proposal
python build_package.py --request examples/ltm-package.json --output-dir candidates/run_verify_ltm --allow-proposal
python build_package.py --request examples/clean-package.json --output-dir candidates/run_verify_clean --allow-proposal
python build_package.py --request examples/edelweiss-package.json --output-dir candidates/run_verify_edelweiss --allow-proposal
```

| Company Key | Request File | Workbook Adapter | Valuation Architecture | Implied Fair Value / Share (vs 15-Oct-2025 Snapshot CMP) |
| :--- | :--- | :--- | :--- | :--- |
| **MM** | `examples/mm-package.json` | `work/models_mm.py` | Conglomerate SOTP (Standalone Auto 16.0x + Farm 14.0x + MMFS/TechM @ 20% HoldCo Disc - Standalone Net Debt) | **₹2,929.98 / sh** (vs CMP ₹2,921.00) |
| **LTM** | `examples/ltm-package.json` | `work/models_ltm.py` | IT Services Connected 5Y FCFF DCF (USD CC Growth * FX, CAPM WACC 11.50%, TGR 5.00%) | **₹2,813.51 / sh** (vs CMP ₹5,568.00) |
| **CLEAN** | `examples/clean-package.json` | `work/models_clean.py` | Specialty Chemicals Manufacturing 5Y FCFF DCF (Parent Units 1–3 + CFCL Unit 4 HALS Ramp, CAPM WACC 12.00%, TGR 5.00%) | **₹415.69 / sh** (vs CMP ₹1,039.00) |
| **EDELWEISS** | `examples/edelweiss-package.json` | `work/models_edelweiss.py` | Financial Services SOTP (EAAA 22.0x PAT + EARCL 1.2x NW + MF 4.0% AAUM + Ins. 1.0x EV + Credit 0.8x NW - HoldCo Net Debt; Industrial FCFF strictly withheld) | **₹133.54 / sh** (vs CMP ₹108.00) |

### Step 6: Run Full Automated Verification Suite
```bash
python verify_sources.py
python -m unittest -v
```
Both commands must exit with status `0`.

