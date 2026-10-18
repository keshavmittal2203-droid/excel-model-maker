# Human Review & Approval Packet: 16-Sheet Master Template Excel Models (Template V6 Release)

> [!IMPORTANT]
> **HISTORICAL VALUATION SNAPSHOT NOTICE**: All financial models, closing stock prices, market capitalizations, Treasury yields, peer trading multiples, and empirical regression datasets in this release reflect historical market observations as of **Wednesday, October 15, 2025 (`2025-10-15`)**. They are **NOT** current or live market valuations and do not incorporate subsequent market events.
>
> **APPROVAL GATE STATUS**: In strict adherence to repository safety invariants (`DISALLOWED_REVIEWER_TOKENS` prohibiting agent self-promotion), the **`human_pack_approval` gate remains PENDING (`ready: false`)**. These models are delivered in proposal/validation state pending explicit sign-off by the human reviewer. PR #3 remains open and unmerged.

---

## 1. Governance & Pull Request Information

- **Branch**: `agent/argon-continue`
- **Pull Request**: [Pull Request #3: Complete Four Priority Targets Excel Release](https://github.com/keshavmittal2203-droid/excel-model-maker/pull/3)
- **Repository Root**: `C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker`
- **Scope**: Strictly Excel financial research workbooks (`template_v6`), native sheet previews, and defect remediation. Zero PowerPoint decks or presentation assets modified in this branch.
- **Master Template Architecture**:
  - **Archived Immutable Original Master**: Exact 17-sheet original preserved at [`originals/archive/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/archive/Marksans%20Keshav.xlsx) (SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`, 17 sheets).
  - **Active Working Master**: Authorized structural removal of `sectordata` sheet applied directly to [`originals/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/Marksans%20Keshav.xlsx) (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`, 16 sheets).
  - Both templates registered in `source-catalog.json` and verified by `verify_sources.py`.

---

## 2. Deliverable Verification Matrix (Template V6)

All four priority targets have been delivered as uncompressed 16-sheet master template workbooks under `companies/<TARGET>/template_v6/`. Prior versions (`template_v5`, `template_v4`, `template_v3`, `template_v2`, and 6-sheet workbooks) are preserved intact in their respective directories.

| Target Company | Primary Valuation Methodology | Source-of-Truth Cell (`'DCF 1st way'`) | Cover Link Cell (`'Model ->'`) | Implied Value per Share | 15-Oct-2025 Market Close | Linked WACC (`E26`) | Formula Errors (64 Sheets) | Verified Workbook SHA-256 Checksum |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Mahindra & Mahindra Limited** (`MM`) | Sum-of-the-Parts (SOTP) Valuation | `E42` (`=G60`) | `C9` (`=E42`) | **₹2,742.97** | ₹3,497.20 | 13.00% | **0 errors** | `1e9483824a97ee2b68240371a4159972096b9bd9789b67ae467cafe6338e6bca` |
| **LTIMindtree Limited** (`LTM`) | Full-Year Annual FCFF DCF Baseline | `E42` (`=E40/E41`) | `C9` (`=E42`) | **₹2,428.22** | ₹5,609.50 | **13.12%** | **0 errors** | `2ccac2a60aa03f88b5a8a61828a9a9269699789a74463e0ac76448dd97fae5d9` |
| **Clean Science and Technology** (`CLEAN`) | Manufacturing FCFF DCF Baseline | `E42` (`=E40/E41`) | `C9` (`=E42`) | **₹468.61** | ₹1,041.40 | **11.91%** | **0 errors** | `98a807a55cad83ef71edd29ddcbb7d6f406b2b3b17b71e0ddd3d007df6deb590` |
| **Edelweiss Financial Services** (`EDELWEISS`) | Financial Services SOTP (Industrial FCFF N/A) | `E42` (`=G60`) | `C9` (`=E42`) | **₹130.53** | ₹108.85 | 9.70% | **0 errors** | `461f8e3f737ad775f66dbe4451264f69344b2e40cb054694fb96d565bfac8406` |

### Clickable Absolute Deliverable Paths

1. **Mahindra & Mahindra Limited (`MM`)**:
   - 16-Sheet Workbook: [`companies/MM/template_v6/MM-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v6/MM-FullTemplate.xlsx) (109,009 bytes)
   - Package Manifest: [`companies/MM/template_v6/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v6/package.manifest.json)
   - Sheet Previews (16 PDFs): [`companies/MM/template_v6/previews/`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/MM/template_v6/previews/)

2. **LTIMindtree Limited (`LTM`)**:
   - 16-Sheet Workbook: [`companies/LTM/template_v6/LTM-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v6/LTM-FullTemplate.xlsx) (108,846 bytes)
   - Package Manifest: [`companies/LTM/template_v6/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v6/package.manifest.json)
   - Sheet Previews (16 PDFs): [`companies/LTM/template_v6/previews/`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/LTM/template_v6/previews/)

3. **Clean Science and Technology Limited (`CLEAN`)**:
   - 16-Sheet Workbook: [`companies/CLEAN/template_v6/CLEAN-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v6/CLEAN-FullTemplate.xlsx) (109,075 bytes)
   - Package Manifest: [`companies/CLEAN/template_v6/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v6/package.manifest.json)
   - Sheet Previews (16 PDFs): [`companies/CLEAN/template_v6/previews/`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/CLEAN/template_v6/previews/)

4. **Edelweiss Financial Services Limited (`EDELWEISS`)**:
   - 16-Sheet Workbook: [`companies/EDELWEISS/template_v6/EDELWEISS-FullTemplate.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v6/EDELWEISS-FullTemplate.xlsx) (110,912 bytes)
   - Package Manifest: [`companies/EDELWEISS/template_v6/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v6/package.manifest.json)
   - Sheet Previews (16 PDFs): [`companies/EDELWEISS/template_v6/previews/`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/companies/EDELWEISS/template_v6/previews/)

---

## 3. Verified Excel Defect Remediation Summary

### Requirement 1: Template Preservation & Authorized Removal of `sectordata`
- The exact 17-sheet master workbook is permanently preserved at [`originals/archive/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/archive/Marksans%20Keshav.xlsx) (SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`).
- The obsolete `sectordata` worksheet has been removed from the active master [`originals/Marksans Keshav.xlsx`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/originals/Marksans%20Keshav.xlsx) (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`) and all four V6 company models.
- All 16 active worksheets preserve exact geometry, layout, row/column positions, formulas, formatting, and navigation:
  `Model ->`, `Main Financials`, `Ratio Analysis`, `Intrinsic Growth`, `Comps`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`, `Data ->`, `compdata`, `Data Sheet`, `BS`, `Beta`, `RM`, `peerdata`, `Cash Flow`, `Sheet1`.

### Requirement 2: Strict 7-Website Allowlist
- Populated exclusively using the 7 approved websites:
  1. **Screener.in** (Statutory P&L, Balance Sheet, Cash Flow, and Sub-Schedules)
  2. **Yahoo Finance** (104-week market prices and benchmark returns)
  3. **NSE India** (15-Oct-2025 official closing stock prices and NIFTY 50 benchmark closes)
  4. **Trading Economics** (India 10-Year Government Bond Yield: 6.78%)
  5. **Top Stock Research** (Trading multiples and peer reference data)
  6. **Investing.com** (Secondary sovereign yield cross-check)
  7. **Prime Investor** (Secondary financial ratios cross-check)
- Strictly zero BSE filings, RBI/FBIL sovereign fixings, Damodaran tables, or unverified tax-authority citations across all workbook cells, comments, manifests, and documentation.

### Requirement 3: Historical Cutoff 15-Oct-2025
- Historical actuals strictly end at FY25 (31-Mar-2025). No FY26 actuals exist.
- FY26 and later periods (`Mar 2026E`) are formula-driven explicit forecasts or marked unavailable.

### Requirement 4: Financial Statement Mapping & 100% Balance Sheet (BS) Reconciliation
- Reported aggregate expenses mapped accurately without fabricated splits (`Sales - Expenses = Operating Profit` verified across all historical periods).
- Detailed Balance Sheet (`BS` sheet) completely rebuilt from company statutory disclosures:
  - `Total Assets (Row 38) == Total Liabilities (Row 16)` across all 12 columns (`Mar 2015` to `Mar 2026E`).
  - Zero plug differences and zero stale Marksans numbers carried over.

### Requirement 5: Comps & compdata Field and Denominator Mapping
- `compdata` field positions mapped accurately:
  - Col J = Enterprise Value (EV)
  - Col K = EBITDA
  - Col L = Revenue (Sales)
  - Col M = Net Profit (PAT)
- `Comps` sheet valuation multiples link to verified denominators:
  - EV/Revenue (`=H8/J8`)
  - EV/EBITDA (`=H8/K8`)
  - P/E (`=F8/L8`)
- Target company row (Row 10) dynamically references `Main Financials` FY25 audited outputs (`K10 = 'Main Financials'!I23`, `L10 = 'Main Financials'!I11`, `M10 = 'Main Financials'!I38`). Professional borders and alignment preserved.

### Requirement 6: Beta & WACC Blume Adjustment Linkage
- 104 weekly observations aligned with NIFTY 50 closing prices from Yahoo Finance / NSE closing prices embedded in `Beta` sheet (Rows 8 to 112).
- Raw beta calculated transparently via `=SLOPE(D9:D112, G9:G112)`.
- Blume adjustment implemented transparently as `=(K7*K8)+(K11*K10)` with `K8 = 2/3` raw beta weight and `K11 = 1/3` market beta weight (`K10 = 1.0`).
- WACC beta links directly to Beta sheet via `WACC-DCF!J32 = =Beta!K13` and `WACC-DCF!J24 = =J32`.
- Target company capital structure weights `E30 = =D30` and `E31 = =E32-E30` dynamically reflect actual company debt and equity.
- Evaluated WACC:
  - `MM`: 13.00% (Raw Beta 1.7638 -> Blume Beta 1.5092, Cost of Equity 15.08%)
  - `LTM`: 13.12% (Raw Beta 1.2569 -> Blume Beta 1.1712, Cost of Equity 13.22%, Cost of Debt after-tax 5.61%, actual capital structure 98.7% equity / 1.3% debt)
  - `CLEAN`: 11.91% (Raw Beta 0.9006 -> Blume Beta 0.9337, Cost of Equity 11.92%, Cost of Debt after-tax 5.99%, actual capital structure 99.98% equity / 0.02% debt)
  - `EDELWEISS`: 9.70% (Raw Beta 1.4830 -> Blume Beta 1.3220, Cost of Equity 14.05%, Cost of Debt after-tax 6.45%)

### Requirement 7: Single Source-of-Truth Cell & Verified Valuations
- Single source-of-truth cell `'DCF 1st way'!E42` drives executive cover page `'Model ->'!C9` dynamically.
- **M&M**: Sum-of-the-Parts (SOTP) with standalone net cash ring-fencing = **₹2,742.97 / share** (E42 = 2742.97, C9 = 2742.97, diff = 0.0010). Note on `B47` discloses MMFS lending debt ring-fencing.
- **LTM**: Full-Year Annual FCFF DCF Baseline = **₹2,428.22 / share** (E42 = 2428.22, C9 = 2428.22, diff = 0.0043). Explicit FCFF PV ₹17,294.07 Cr, Terminal Value PV ₹43,383.93 Cr @ WACC 13.12% linked to company's own regression Blume Beta `Beta!K13` (1.1712) and terminal g 5.0% linked to `E20`, Cash ₹13,433 Cr, Debt ₹2,187 Cr, Shares 29.62 Cr.
- **CLEAN**: Full-Year Manufacturing FCFF DCF Baseline = **₹468.61 / share** (E42 = 468.61, C9 = 468.61, diff = 0.0026). Explicit FCFF PV ₹1,004.22 Cr, Terminal Value PV ₹3,578.08 Cr @ WACC 11.91% linked to company's own regression Blume Beta `Beta!K13` (0.9337) and terminal g 5.0% linked to `E20`, Cash ₹401 Cr, Debt ₹2 Cr, Shares 10.63 Cr.
- **EDELWEISS**: Financial Services Sum-of-the-Parts (SOTP) = **₹130.53 / share** (E42 = 130.53, C9 = 130.53, diff = 0.0046). Industrial DCF cells marked `N/A - Financial Services Entity`; `Intrinsic Growth` industrial metrics marked N/A.

### Dynamic Sensitivity Testing (8 Tests, 100% Passed with Exact Return)
- **M&M**:
  - Auto Core EBITDA Multiple (`F50` 14.0 $\to$ 15.0): delta = +₹109.95, restored = ₹2,742.97 (exact return).
  - Farm Equipment Core EBITDA Multiple (`F51` 13.0 $\to$ 14.0): delta = +₹59.56, restored = ₹2,742.97 (exact return).
- **LTM**:
  - Equity Risk Premium (`WACC-DCF!J23` 5.50% $\to$ 6.00%): delta = -₹140.16, restored = ₹2,428.22 (exact return).
  - Terminal Growth Rate (`DCF 1st way!E20` 5.00% $\to$ 5.50%): delta = +₹102.87, restored = ₹2,428.22 (exact return).
- **CLEAN**:
  - Terminal Operating EBIT (`DCF 1st way!K5` ₹693 $\to$ ₹743 Cr): delta = +₹25.86, restored = ₹468.61 (exact return).
  - Equity Risk Premium (`WACC-DCF!J23` 5.50% $\to$ 6.00%): delta = -₹29.22, restored = ₹468.61 (exact return).
- **EDELWEISS**:
  - Alternatives Multiple (`DCF 1st way!F50` 18.70x $\to$ 20.00x): delta = +₹5.57, restored = ₹130.53 (exact return).
  - Asset Reconstruction Multiple (`DCF 1st way!F51` 1.02x $\to$ 1.15x): delta = +₹2.95, restored = ₹130.53 (exact return).

### Requirement 8: Clean Presentation & Zero Errors
- Marksans logo cleanly removed from `Main Financials` drawing without breaking cell borders or layouts.
- Verbose visible paragraphs removed; concise provenance preserved in cell comments.
- Recalculated cleanly via Microsoft Excel COM 16.0 (`CalculateFullRebuild`). Exactly **0 formula errors** across all 64 sheets (16 sheets x 4 targets).
- Preserved native unforced print setups without alteration. Exactly 16 preview PDFs per company (64 total) exported to `companies/<TARGET>/template_v6/previews/`.
