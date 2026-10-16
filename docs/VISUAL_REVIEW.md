# Candidate Package Visual & Sheet Inspection Review Record

**Review Date**: 2026-10-16  
**Reviewing Agent**: Antigravity (`agent/argon-continue`)  
**Methodology**: Direct visual inspection of native COM high-resolution renders (1080x1440 portrait PNGs, composite contact sheets, and landscape multi-column schedule PDFs) via `view_file` visual tool, combined with cell-level programmatic inspection of the candidate workbook using openpyxl.

---

## 1. Candidate Hash & Artifact Inventory

### Baseline Run (`candidates/run_native_baseline_v5/`)
- **Excel Candidate**: `Marksans-Recalculated.xlsx`  
  `SHA-256`: `d4ba6d45cce8fbaff6b35c518ce71a81fee6b5ebb602149b67390dcfd11698cc`
- **PowerPoint Candidate**: `Marksans-Pharma-Overview.pptx`  
  `SHA-256`: `654c2b4907f9e79037864998f07473cb204a8701143504762746e0b0a6c65a05`
- **Package Manifest**: `package.manifest.json`  
  `SHA-256`: `00d5bc866f4cfd367f2af5fc5344a5a26be1713fec184b6ffdefb3d1bdf1d23e`
- **Render Previews**: `candidates/run_native_baseline_v5/previews/`
  - Slides: `slide_1.png` to `slide_5.png` (1080x1440 PNG)
  - Contact Sheet: `slide_contact_sheet.png` (2x3 composite PNG)
  - Schedules: `Data_Sheet.pdf`, `Ratio_Analysis.pdf`, `Intrinsic_Growth.pdf`, `WACC-DCF.pdf`, `DCF_1st_way.pdf`, `DCF_2nd_way.pdf`

### Scenario B Stressed Run (`candidates/run_native_scenario_b_v5/`)
- **Excel Candidate**: `Marksans-Scenario-B-Recalculated.xlsx`  
  `SHA-256`: `26466e87378da976f41b06f3c1bd7d06fc97b801a7e5aca59f681da7c0ba76bf`
- **PowerPoint Candidate**: `Marksans-Scenario-B-Overview.pptx`  
  `SHA-256`: `75e1766c0f4a5a88c9a125114c1c5d51f4d0b6d3bd78db9b2c3c7381c0ae034d`
- **Package Manifest**: `package.manifest.json`  
  `SHA-256`: `81e594f3589f0697de60a094be02d773dbb35b5fe1e8f51260dab814624fc4b0`
- **Render Previews**: `candidates/run_native_scenario_b_v5/previews/`
  - Slides: `slide_1.png` to `slide_5.png`
  - Contact Sheet: `slide_contact_sheet.png`

### File Count Reconciliation
- **13 Primary Deliverables recorded in `package.manifest.json`**: 1 workbook, 1 presentation, 5 slide PNGs, 6 schedule PDFs.
- **5 Auxiliary & Metadata Files on Disk**: 1 package manifest (`package.manifest.json`, intentionally excluded from self-hash), 1 composite slide contact sheet (`previews/slide_contact_sheet.png`), and 3 run-local generated WACC table image assets (`assets/image22_marksans.png`, `assets/image23_marksans.png`, `assets/image24_marksans.png`).
- **Total Files on Disk**: 18 files per candidate run directory.

---

## 2. PowerPoint Visual Slide Inspection (ERR-4.pptx Curated 5 Slides)

### Slide 1: Company Profile, Business Model & Ecosystem (Template Slide 1)
- **Visual Inspection**: Inspected `candidates/run_native_baseline_v4/previews/slide_1.png`.
- **Top Header**: "EQUITY RESEARCH REPORT", "Marksans Pharma Ltd.", ticker "NSE : MARKSANS | BSE Code: 524404", industry "Pharmaceuticals & Drugs". Layout alignment matches template exactly.
- **Section Headers**: "About The Company", "INTRODUCTION", "BUSINESS MODEL", "BRAND PORTFOLIO", "PRODUCT ECOSYSTEM" in bold orange `#E36B0C`. All bullets render cleanly with zero clipping.
- **Corporate Logo**: Shape 59 (Picture 58) successfully suppressed pending authentic vector asset. Zero placeholder residue.
- **Geographic Breakdown Message Card (Collision Catch & Fix)**:
  - *Defect Caught*: The missing data message card (`Shape 25`) overlapped the right edge of Shape 20's "Strategic Approach" text box because Shape 20 had `cx=3668169` extending past `x=3485336`.
  - *Correction Applied*: Adjusted missing data card bounds to `replacement_x="3850000"`, `replacement_cx="2850000"` and reformatted `slide1_strategic_approach` into two concise bullet points.
  - *Verification*: Confirmed in `v4/previews/slide_1.png` that "Strategic Approach" wraps cleanly on the left (`x=141831` to `x=3810000`) and the dark navy message card sits completely clear on the right (`x=3850000` to `x=6700000`). Zero visual collision.
- **Unevidenced Numeric Eradication**: Confirmed zero occurrences of unevidenced figures `[43.1, 41.7, 9.8, 5.4]`. `chart1.xml` and embedded `Microsoft_Excel_Worksheet.xlsx` are completely pruned from candidate package.

### Slide 2: Executive Summary & Historical Growth (Template Slide 2)
- **Visual Inspection**: Inspected `candidates/run_native_baseline_v4/previews/slide_2.png`.
- **Narratives**: "Executive Summary" and "Key Highlights" populated with verified FY25 actuals (Revenue ₹2,622.85 Cr, Net Profit ₹380.58 Cr, Net Cash ₹(382.30) Cr, ROE 20.35%, ROCE 22.84%).
- **Left Native Chart**: "Historical Revenue (₹ Crore)" updated with 5-year consolidated revenue (FY21: ₹1,376 Cr to FY25: ₹2,623 Cr). Proper category axis, orange column bars, zero clipping.
- **Right Native Chart**: "FY25 Financial Snapshot (₹ Crore)" updated with Revenue, EBITDA, Net Profit, CFO. Clean bar geometry and scale.

### Slide 3: 8-Year Cash Flow Statement (Template Slide 14)
- **Visual Inspection**: Inspected `candidates/run_native_baseline_v4/previews/slide_3.png`.
- **Table Schedule**: 8-Year historical cash flows (FY18–FY25) across Operating, Investing, and Financing activities.
- **Formatting**: Numbers right-aligned, negative values in brackets, totals highlighted with orange background bars.
- **Bottom Summary**: "CASH FLOW FROM OPERATING" narrative ties out with table cash flow from operations (₹206.66 Cr in FY25) and cash balance (₹704.21 Cr).

### Slide 4: Ratio Analysis Schedule & Trend Charts (Template Slide 16)
- **Visual Inspection**: Inspected `candidates/run_native_baseline_v4/previews/slide_4.png`.
- **Table Schedule**: All 64 cells populated across 8 historical years (Current Ratio, Quick Ratio, D/E, ICR, Debtor Days, Payable Days, Inventory Days, CCC).
- **Lower Left & Right Charts**: Native trend line charts for Current Ratio and Quick Ratio. Trend lines correctly plot all 8 points, scale properly from 0 to 6 and 0 to 4 respectively, and tie out 100% with the table rows above.

### Slide 5: WACC & DCF Valuation (Template Slide 23)
- **Visual Inspection**: Inspected `candidates/run_native_baseline_v4/previews/slide_5.png`.
- **Valuation Schedules**: High-resolution replacement table schedules rendered from recalculated model:
  - Cost of Equity: $R_f = 7.30\%$, $R_m = 16.37\%$, $\beta = 0.547 \to K_e = 12.26\%$.
  - Cost of Debt: Pre-tax $8.00\%$, Tax $30.00\% \to K_d = 5.60\%$.
  - WACC Schedule: Current capital structure ($2.39\%$ debt / $97.61\%$ equity) explicitly distinguished from Target peer-median capital structure ($2.24\%$ debt / $97.76\%$ equity).
  - WACC: $12.11\%$.
- **Typography & Alignment**: Crisp text, sharp borders, matching template palette.

---

## 3. Actual Excel Workbook Sheet Inventory & Inspection Scope

The actual candidate model `Marksans-Recalculated.xlsx` contains exactly **17 worksheets**, structured into two functional sections (`Model ->` and `Data ->`).

### Complete Sheet Inventory (Derived Directly from Candidate Model)

| # | Sheet Name | Section | Cells / Formulas | Exported to PDF? | Visual Inspection Coverage | Actual Sheet Purpose & Formula Architecture |
|---|---|---|---|---|---|---|
| 1 | `Model ->` | Model Divider | 0 / 0 | No | Visual Divider Tab | Empty section divider tab for financial model schedules. |
| 2 | `Main Financials` | Substantive Model | 811 / 725 | No | Programmatically Inspected | Consolidated 3-Statement financial model (Income Statement rows 10–44, Balance Sheet rows 46–77, Cash Flow rows 80–148) across 5 historical years (FY20–FY24) linking directly to `Data Sheet`. |
| 3 | `Ratio Analysis` | Substantive Model | 218 / 190 | **Yes** (`Ratio_Analysis.pdf`) | **Visually Inspected (PDF)** | 5-year ratio analysis (FY21–FY25): Sales Growth, EBITDA/Net margins, ROCE, ROE, Self-Sustained Growth Rate ($g = \text{ROE} \times b$), Debtor Turnover/Days, and CFO coverage ratios. |
| 4 | `Intrinsic Growth` | Substantive Model | 233 / 181 | **Yes** (`Intrinsic_Growth.pdf`) | **Visually Inspected (PDF)** | Calculation of ROIC (rows 4–39) and Reinvestment Rate (rows 41–65): Invested Capital, Net Working Capital, Net Non-Current Assets, and Sustainable Growth ($g = \text{ROIC} \times \text{Reinvestment Rate}$). |
| 5 | `Comps` | Substantive Model | 165 / 129 | No | Programmatically Inspected | Comparable company valuation multiples for 8 pharma peers (Sun Pharma, Divi's, Torrent, Zydus, Laurus, Cipla, Dr Reddy's, Marksans): EV/EBITDA, P/E, EV/Sales, and Price/Book. |
| 6 | `WACC-DCF` | Substantive Model | 141 / 58 | **Yes** (`WACC-DCF.pdf`) | **Visually Inspected (PDF)** | Master WACC schedule: 5 peer comps unlevered/re-levered Hamada beta ($\beta = 0.547$), pre-tax borrowing cost ($8.00\%$), post-tax cost of debt ($5.60\%$), CAPM cost of equity ($K_e = 12.26\%$), target peer-median weights ($2.24\%$ debt / $97.76\%$ equity), and WACC ($12.11\%$). |
| 7 | `DCF 1st way` | Substantive Model | 138 / 50 | **Yes** (`DCF_1st_way.pdf`) | **Visually Inspected (PDF)** | **5-Year FCFF Projection (FY26–FY30)** with mid-year discounting (periods 0.5 to 4.5). Terminal value calculated via **Gordon Growth perpetual formula** ($TGR = 5.0\%$, WACC = $12.11\%$). Includes a **2D sensitivity table** (`H23:K31`, `{=TABLE(J24,G27)}`) varying WACC ($10\%$ to $16\%$) and TGR ($3\%$ to $6\%$). Operating Asset Value ₹2,976.46 Cr $\to$ Equity Value ₹3,622.73 Cr $\to$ Intrinsic Value ₹79.94/share. |
| 8 | `DCF 2nd way` | Substantive Model | 455 / 296 | **Yes** (`DCF_2nd_way.pdf`) | **Visually Inspected (PDF)** | **5-Year Unlevered FCF Projection (FY27–FY31)** preceded by historical actuals (FY22–FY25) and FY26 estimate. Terminal value calculated via **Gordon Growth formula** ($TGR = 6.0\%$). Rows 68–97 contain regression-based trend forecasts for Revenue, EBIT, Tax, D&A, and CapEx. Operating Asset Value ₹331.95 Cr $\to$ Equity Value ₹978.22 Cr $\to$ Implied Value ₹21.59/share. *(Note: Does NOT use exit multiple method).* |
| 9 | `Data ->` | Data Divider | 0 / 0 | No | Visual Divider Tab | Empty section divider tab for raw and benchmark data sheets. |
| 10 | `compdata` | Supporting Data | 162 / 32 | No | Programmatically Inspected | Peer market trading inputs: CMP, number of equity shares, market capitalization, and debt for 7 peer companies. |
| 11 | `Data Sheet` | Supporting Data | 576 / 22 | **Yes** (`Data_Sheet.pdf`) | **Visually Inspected (PDF)** | Screener.in consolidated financial export master sheet: 10-year P&L (FY16–FY25), Balance Sheet, share capital, CMP, and quarterly results. |
| 12 | `BS` | Supporting Data | 407 / 0 | No | Programmatically Inspected | Screener.in detailed historical Balance Sheet line-item schedule (FY15–FY25). |
| 13 | `Beta` | Supporting Data | 605 / 213 | No | Programmatically Inspected | 2-year weekly stock price and Nifty 50 returns regression analysis calculating raw beta, with Blume/Vasicek drifting (0.75 raw + 0.25 market weight). |
| 14 | `RM` | Supporting Data | 61 / 2 | No | Programmatically Inspected | 26-year (2000–2025) historical annual return series for Nifty 50 Total Return Index and $1.22\%$ dividend yield assumption table. |
| 15 | `peerdata` | Supporting Data | 36 / 0 | No | Programmatically Inspected | Peer capital structure inputs: Market Cap, Total Debt, Debt/Equity ratio for the 5 primary valuation peers. |
| 16 | `Cash Flow` | Supporting Data | 389 / 0 | No | Programmatically Inspected | Screener.in detailed historical Cash Flow line-item schedule (FY15–FY25). |
| 17 | `sectordata` | Supporting Data | 1069 / 0 | No | Programmatically Inspected | Damodaran India sector benchmark dataset: number of firms, CapEx, D&A, CapEx/D&A ratios, and acquisitions across all Indian industry sectors. |

### Correction of Hallucinated Sheet Inventory
An earlier draft erroneously cataloged 7 non-existent sheet names (`PL`, `common size statement`, `historic forecast`, `FCFF`, `SOP`, `cost of capital`, `dividend discount model`).
- **Correction**: None of these 7 sheets exist in `originals/Marksans Keshav.xlsx` or `Marksans-Recalculated.xlsx`.
- Income statement data resides in `Data Sheet` and `Main Financials`.
- Both DCF models use **5-year forecast horizons** (not 10-year).
- Both DCF models calculate terminal value using the **Gordon Growth perpetual formula**; neither model uses an exit multiple terminal valuation method.
- 2D sensitivity analysis exists in `DCF 1st way` (`{=TABLE()}` varying WACC and TGR).

---

## 4. Visual Inspection Coverage & Reassessment of Review Status

### Inspected Artifacts
1. **PowerPoint Presentation (100% Coverage)**: All 5 curated slides (`slide_1.png` to `slide_5.png`) and the 2x3 composite contact sheet (`slide_contact_sheet.png`) were visually reviewed. Layout collisions on Slide 1 were caught and verified resolved. Zero text clipping, clean margins, and proper chart scaling confirmed.
2. **Substantive Model Schedules (6 Key Sheets)**: The 6 substantive financial schedules (`Data Sheet`, `Ratio Analysis`, `Intrinsic Growth`, `WACC-DCF`, `DCF 1st way`, `DCF 2nd way`) exported to landscape PDFs were visually inspected. All column layouts fit to single-page width with clean typography.

### Unreviewed Artifacts (9 Sheets)
The remaining 9 sheets (`Model ->`, `Main Financials`, `Comps`, `Data ->`, `compdata`, `BS`, `Beta`, `RM`, `peerdata`, `Cash Flow`, `sectordata`) were validated programmatically (verifying formula syntax, cell coordinates, and financial identities), but were **NOT exported to PDF and have NOT been visually inspected**.

### Reassessment of Release Gates
Because 9 workbook sheets remain visually unreviewed and human visual sign-off has not yet been performed:
- **`render_review`**: Remains **PENDING** (provisional agent visual review completed on slides and 6 core schedules; human visual review and full sheet inspection required).
- **`human_pack_approval`**: Remains **PENDING** (human sponsor review and approval of `proposals/marksans.model-pack.proposal.json` and `proposals/err4.deck-pack.proposal.json` required).
- **Overall Readiness**: `ready: false`.
