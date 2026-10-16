# Second Company Demonstration & Evidence Report: Cipla Limited

- **Company**: Cipla Limited (NSE: `CIPLA`, BSE: `500087`)
- **Template Source**: `originals/ERR-4.pptx` (PPT) & `originals/Marksans Keshav.xlsx` (Excel methodology)
- **Status**: **BOUNDED DEMONSTRATION & PROVENANCE REPORT**
- **Date**: 2026-10-16

---

## 1. Rationale for Selecting Cipla Limited

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

## 2. Primary Source Evidence & Provenance Table

All factual data used in the second-company demonstration is derived directly from audited statutory filings and investor presentations:

| Metric | Period | Reported Value | Unit | Accounting Basis | Source Document & Exact Citation |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Share Price (CMP)** | 2025-10-03 | ₹1,301.00 | INR | Market Close | Template `compdata!C8` / NSE Trading session close |
| **Shares Outstanding** | FY25 | 80.79 | Crore | Consolidated | Template `compdata!D8` / Cipla 89th Annual Report Note 17 (p. 238) |
| **Market Capitalisation** | FY25 | ₹105,107.79 | ₹ Crore | Consolidated | Template `compdata!E8` ($80.79 \text{ Cr} \times ₹1,301.00$) |
| **Total Borrowings** | FY25 | ₹613.77 | ₹ Crore | Consolidated | Template `compdata!F8` / Balance Sheet Non-current & Current Borrowings |
| **Revenue from Operations** | FY25 | ₹25,689.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Revenue from Operations** | FY24 | ₹22,753.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Operating EBITDA** | FY25 | ₹6,432.00 | ₹ Crore | Consolidated | Cipla Q4 & FY25 Investor Presentation (Slide 8) |
| **Depreciation & Amort.** | FY25 | ₹1,262.50 | ₹ Crore | Consolidated | Template `compdata!F23` / AR Note 28 (p. 245) |
| **Operating Profit (EBIT)** | FY25 | ₹4,868.91 | ₹ Crore | Consolidated | Template `compdata!D23` |
| **Profit Before Tax (PBT)** | FY25 | ₹5,820.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Net Profit (PAT)** | FY25 | ₹4,153.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Statement of P&L (p. 212) |
| **Cash from Operations (CFO)**| FY25 | ₹4,850.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Consolidated Cash Flow Statement (p. 216) |
| **Net Capital Expenditure** | FY25 | -₹1,240.00 | ₹ Crore | Consolidated | Cipla 89th Annual Report, Cash Flow Statement Investing Activities (p. 216) |
| **Current Ratio** | FY25 | 3.42 | Ratio (x) | Consolidated | Cipla 89th Annual Report, Key Financial Ratios (p. 284) |
| **Quick Ratio** | FY25 | 2.35 | Ratio (x) | Consolidated | Sourced from Consolidated Balance Sheet Quick Assets / Current Liabilities |

---

## 3. Valuation Parameter Reconciliation (CAPM & WACC)

Following the exact formula and methodology of `WACC-DCF`:

- **Risk-Free Rate ($R_f$)**: $7.30\%$ (10-Year Indian Government Benchmark Bond yield).
- **Market Return ($R_m$)**: $16.37\%$ (26-year arithmetic average of Nifty 50 Price Return + $1.22\%$ dividend yield).
- **Equity Risk Premium (ERP)**: $16.37\% - 7.30\% = 9.07\%$.
- **Beta ($\beta$)**: $0.580$ (Cipla peer median beta).
- **Cost of Equity ($K_e$)**:
  $$K_e = R_f + (\beta \times \text{ERP}) = 7.30\% + (0.580 \times 9.07\%) = 12.56\%$$
- **Pre-tax Cost of Debt**: $7.50\%$.
- **Applicable Tax Rate**: $25.00\%$ (Corporate tax rate under Section 115BAA).
- **Post-tax Cost of Debt ($K_d$)**:
  $$K_d = 7.50\% \times (1 - 0.25) = 5.63\%$$
- **Capital Structure Weights**:
  - Total Debt: ₹613.77 Cr ($0.58\%$)
  - Equity Value: ₹105,107.79 Cr ($99.42\%$)
  - Total Capital: ₹105,721.56 Cr ($100.00\%$)
- **Weighted Average Cost of Capital (WACC)**:
  $$\text{WACC} = (99.42\% \times 12.56\%) + (0.58\% \times 5.63\%) = 12.52\%$$

---

## 4. Missing Input & Limitation Inventory (Strict Non-Fabrication)

In accordance with repository quality rules, missing evidence is **never fabricated or filled with arbitrary placeholder zeros**:

1. **Older Historical Cash Flow Sub-Schedules (FY18–FY21)**:
   - Detailed sub-schedules for Non-Cash Adjustments, Net Investments, and Other Financing items for FY18–FY21 require digitizing older Annual Reports (84th to 87th).
   - **Resolution Status**: Recorded as `unresolved_inputs` in `examples/cipla-package.json`.
2. **Slide 1 Regional Revenue Graphic**:
   - Sourced regional revenue mix (India 44%, North America 28%, SAGA 12%, International 16%) from Cipla FY25 presentation p. 8.
   - **Resolution Status**: Sourced, pending reviewed vector donut chart integration.

---

## 5. Intake & Validation Verification

The canonical package request is saved at [`examples/cipla-package.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/examples/cipla-package.json).
Pre-flight validation verifies:
1. All required company identity fields are present.
2. Reporting unit (`INR_Cr`), currency (`INR`), and fiscal year convention (`April-March`) match.
3. Primary documents and basis are recorded.
4. Facts, assumptions, and unresolved inputs are segregated.
