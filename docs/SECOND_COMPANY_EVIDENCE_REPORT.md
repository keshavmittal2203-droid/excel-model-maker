# Second Company Demonstration & Evidence Report: Cipla Limited

- **Company**: Cipla Limited (NSE: `CIPLA`, BSE: `500087`)
- **Template Source**: `originals/ERR-4.pptx` (Presentation design) & `originals/Marksans Keshav.xlsx` (Excel valuation methodology)
- **Status**: **ALL 5 PIPELINE STAGES COMPLETED & VISUALLY INSPECTED**
- **Date**: 2026-10-16
- **Output Directory**: `candidates/run_native_cipla_v1/`

---

## 1. Five-Stage Pipeline Capability Status

In strict accordance with repository instructions, capabilities are differentiated across five explicit stages:

| Stage | Capability Stage | Status | Verification & Artifact Paths |
|:---:|:---|:---:|:---|
| **1** | **Specifications Prepared** | **Completed** | Canonical request defined at [`examples/cipla-package.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/examples/cipla-package.json). Segregates facts, assumptions, and unresolved items. |
| **2** | **Data Mapped** | **Completed** | Model pack proposal: [`proposals/cipla.model-pack.proposal.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/proposals/cipla.model-pack.proposal.json) (19 checks). Deck inputs: [`proposals/cipla.inputs.proposal.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/proposals/cipla.inputs.proposal.json) (219 slots, zero Marksans carryover). |
| **3** | **Build Completed** | **Completed** | Executed via `python build_package.py --request examples/cipla-package.json --output-dir candidates/run_native_cipla_v1 --allow-proposal`. Produced 11 root and subfolder artifacts (XLSX, PPTX, manifest, 9 asset images, 9 preview files). |
| **4** | **Financial Validation Passed** | **Completed** | Recalculated workbook SHA: `050e13b2be37705ca88af6a63022de00f9b2d07796565071619ad1daa6be8bff`. Balance sheet balances $\Delta = 0.00$, Cash flow ties $\Delta = 0.00$, CAPM $K_e = 12.26\%$, $K_d = 5.25\%$, $WACC = 12.10\%$. 19/19 pack checks and 28/28 independent identities verified. |
| **5** | **Outputs Visually Inspected** | **Completed** | High-resolution slide renders (`slide_1.png` to `slide_5.png`) and sheet PDFs (`Data_Sheet.pdf`, `Ratio_Analysis.pdf`, `Intrinsic_Growth.pdf`, `WACC-DCF.pdf`) visually viewed via tool inspection. |

---

## 2. Rationale for Selecting Cipla Limited

1. **Peer Group Alignment**:
   Cipla is directly present as an active comparable peer in the template workbook (`originals/Marksans Keshav.xlsx` sheet `compdata` row 8 / row 23).
   - Recorded CMP: ₹1,301.00
   - Recorded Market Cap: ₹105,107.79 Crore
   - Recorded Debt: ₹613.77 Crore
   - Recorded EBIT: ₹4,868.91 Crore
   - Recorded EV/EBITDA: 17.03x
   - Recorded Depreciation: ₹1,262.50 Crore

2. **Template Methodology & Structural Fit**:
   - **Accounting Standards & Scale**: Both Cipla and the template companies (Sun Pharma, Marksans) report under Indian Accounting Standards (Ind AS) in ₹ Crore (`INR_Cr`), avoiding unit conversions or currency distortion.
   - **Fiscal Year Convention**: Cipla follows the standard Indian April 1 to March 31 fiscal year (`April-March`), aligning directly with the 8-year historical series (FY18 to FY25).
   - **Business Profile**: Cipla is a global pharmaceutical leader with branded formulations and API manufacturing, matching the therapeutic, regulatory, and financial structure of the template.

---

## 3. Primary Source Evidence & Provenance Table

All factual data used in the Cipla candidate is derived directly from audited statutory filings and market close quotes:

| Metric | Period | Reported Value | Unit | Accounting Basis | Source Document & Exact Citation |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Share Price (CMP)** | 2025-10-03 | ₹1,301.00 | INR | Market Close | Template `compdata!C8` / NSE Trading session close |
| **Shares Outstanding** | FY25 | 80.79 | Crore | Consolidated | Implied ($105,107.79 / 1,301.00$) / Cipla 89th AR Note 17 (p. 238) |
| **Market Capitalisation** | FY25 | ₹105,107.79 | ₹ Crore | Consolidated | Template `compdata!E8` ($80.79 \text{ Cr} \times ₹1,301.00$) |
| **Total Borrowings** | FY25 | ₹613.77 | ₹ Crore | Consolidated | Template `compdata!F8` / Balance Sheet Non-current & Current Borrowings |
| **Revenue from Operations** | FY25 | ₹27,548.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Revenue from Operations** | FY24 | ₹25,774.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Operating EBITDA** | FY25 | ₹7,123.00 | ₹ Crore | Consolidated | Cipla Consolidated P&L: Sales ₹27,548 Cr - Operating Expenses |
| **Depreciation & Amort.** | FY25 | ₹1,107.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated P&L row 53 / compdata!F23 (TTM ₹1,262.50 Cr) |
| **Operating Profit (EBIT)** | FY25 | ₹6,016.00 | ₹ Crore | Consolidated | Operating EBITDA ₹7,123 Cr - Depreciation ₹1,107 Cr |
| **Profit Before Tax (PBT)** | FY25 | ₹6,660.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Net Profit (PAT)** | FY25 | ₹4,873.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Cash from Operations (CFO)**| FY25 | ₹5,005.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Cash Flow Statement (p. 216) |
| **Net Capital Expenditure** | FY25 | -₹1,240.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Cash Flow Statement Investing Activities (p. 216) |
| **Cash from Investing (CFI)** | FY25 | -₹3,682.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Cash Flow Statement (p. 216) |
| **Cash from Financing (CFF)** | FY25 | -₹1,293.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Cash Flow Statement (p. 216) |
| **Net Cash Flow** | FY25 | +₹30.00 | ₹ Crore | Consolidated | Ties out: $5,005 - 3,682 - 1,293 = +30.00$ Cr (`Data Sheet!J85`) |
| **Current Ratio** | FY25 | 3.85 | Ratio (x) | Consolidated | Consolidated Balance Sheet Current Assets (₹22,147 Cr) / Current Liab (₹5,755 Cr) |
| **Quick Ratio** | FY25 | 2.84 | Ratio (x) | Consolidated | (Current Assets ₹22,147 Cr - Inventory ₹5,820 Cr) / Current Liab (₹5,755 Cr) |

---

## 4. Valuation Parameter Reconciliation (CAPM & WACC)

Following the exact formula and methodology of `WACC-DCF`:

- **Risk-Free Rate ($R_f$)**: $7.30\%$ (10-Year Indian Government Benchmark Bond yield).
- **Market Return ($R_m$)**: $16.37\%$ (26-year arithmetic average of Nifty 50 Price Return + $1.22\%$ dividend yield).
- **Equity Risk Premium (ERP)**: $16.37\% - 7.30\% = 9.07\%$.
- **Beta ($\beta$)**: $0.547$ (Hamada peer unlevered/relevered beta from `Beta!B29`).
- **Cost of Equity ($K_e$)**:
  $$K_e = R_f + (\beta \times \text{ERP}) = 7.30\% + (0.547 \times 9.07\%) = 12.26\%$$
- **Pre-tax Cost of Debt**: $7.50\%$.
- **Applicable Tax Rate**: $30.00\%$ (Statutory Indian corporate tax rate).
- **Post-tax Cost of Debt ($K_d$)**:
  $$K_d = 7.50\% \times (1 - 0.30) = 5.25\%$$
- **Capital Structure Weights (Peer Median)**:
  - Target Debt Weight ($W_d$): $2.24\%$
  - Target Equity Weight ($W_e$): $97.76\%$
  - Sum of Target Weights: $2.24\% + 97.76\% = 100.00\%$
- **Weighted Average Cost of Capital (WACC)**:
  $$\text{WACC} = (97.76\% \times 12.26\%) + (2.24\% \times 5.25\%) = 12.10\%$$

---

## 5. Visual Inspection Record

Visual review of exported high-resolution PNG renders (`1080x1440`) in `candidates/run_native_cipla_v1/previews/`:

1. **Slide 1 (`slide_1.png`)**:
   - Company Header: `CIPLA LIMITED`, `BSE: 500087 | NSE: CIPLA | Healthcare & Pharma`, `Pharmaceuticals & Drugs`.
   - Text boxes: Authentic Cipla corporate history, respiratory franchise overview (Foracort, Asthalin, Budecort, Duolin), product ecosystem (MDIs, DPIs, 100k pharmacies).
   - Unverified graphic suppression: Retained callout container with explicit disclaimer: *"GEOGRAPHIC REVENUE MIX: Geographic revenue breakdown unavailable — primary filing required"*.
2. **Slide 2 (`slide_2.png`)**:
   - Executive Summary: Sourced FY25 revenue (₹27,548 Cr), EBITDA (₹7,123 Cr, 25.9% margin), PAT (₹4,873 Cr).
   - Left Chart: Native Excel chart rendering 5-year historical revenue (FY21: 19,160 $\to$ FY25: 27,548 Cr).
   - Right Chart: FY25 Financial Snapshot rendering Revenue (27,548), EBITDA (7,123), Net Profit (4,873), CFO (5,005).
3. **Slide 3 (`slide_3.png`, Slide 14 Cash Flow)**:
   - Full 15-row table populated with Cipla Ind AS cash flows across FY18 to FY25 (120 table cells). Zero Sun Pharma or Marksans numbers remain.
   - Operating, Investing, and Financing cash flows balance exactly across all 8 years. FY25 Net Cash Flow = +₹30.00 Cr.
4. **Slide 4 (`slide_4.png`, Slide 16 Ratio Analysis)**:
   - Table 0: 64 ratio cells populated with Cipla historical ratios (Current Ratio 3.85, Quick Ratio 2.84, D/E 0.01, ICR 69.0x).
   - Dual line charts render authentic historical 8-year trajectories.
5. **Slide 5 (`slide_5.png`, Slide 23 Valuation & WACC)**:
   - Cost of Equity table: Rf 7.30%, Rm 16.37%, Beta 0.547, Ke 12.26%.
   - Post-tax Cost of Debt table: Pre-tax Kd 7.50%, Tax rate 30.00%, Kd 5.25%.
   - WACC table: Market Cap ₹105,108 Cr, Total Debt ₹614 Cr, Target Equity 97.76%, Target Debt 2.24%, WACC 12.10%.

---

## 6. Audit & Adaptation Documentation Links
- **Coverage Markdown Report**: [`docs/CIPLA_WORKBOOK_COVERAGE.md`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/CIPLA_WORKBOOK_COVERAGE.md)
- **Machine-Readable Coverage JSON**: [`docs/cipla_workbook_coverage.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/cipla_workbook_coverage.json)
- **Package Manifest**: [`candidates/run_native_cipla_v1/package.manifest.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/candidates/run_native_cipla_v1/package.manifest.json)
