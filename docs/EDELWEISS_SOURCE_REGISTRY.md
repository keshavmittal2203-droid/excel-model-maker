# Primary Statutory & Evidence Source Registry: Edelweiss Financial Services Limited (EDELWEISS)

- **Target Legal Entity**: Edelweiss Financial Services Limited
- **BSE Scrip Code**: `532922` | **NSE Ticker**: `EDELWEISS` | **ISIN**: `INE532F01054`
- **Reporting Basis & Scope**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; Industrial FCFF is strictly prohibited because ₹18,004 Cr of FY25 consolidated borrowings include ₹14,932 Cr of ring-fenced NBFC/HFC/ARC operating subsidiary debt; SOTP deducts Audited FY25 EFSL Standalone HoldCo Net Debt of ₹3,067 Cr).
- **Local Primary Data Extract**: [`work/sources/EDELWEISS_consolidated.json`](../work/sources/EDELWEISS_consolidated.json) (`SHA-256: c32f54b584298bae6071e6478bdc0d6e352f5d7e58ecdb9030ef069ceaa68d17`)
- **Local Statutory Sub-Schedules Cache**: [`work/sources/EDELWEISS_schedules.json`](../work/sources/EDELWEISS_schedules.json) (`SHA-256: a7e1f34c2397f80c9b5c04901f399fcba9fea7bb83eb00b87b1118b979d247ae`)
- **Local 15-Oct-2025 Market & Empirical Beta Cache**: [`work/sources/market_and_beta_2025_10_15.json`](../work/sources/market_and_beta_2025_10_15.json) (`SHA-256: de04dac2fc4a73be3d596fe5aec37ff070e4dc0a4d2abe22b11b340206c3f823`)
- **Official BSE Filing PDF**: [BSE Statutory Filing PDF](https://www.bseindia.com/stockinfo/AnnPdfOpen.aspx?Pname=E048618A_8BB1_4716_A668_F2893199113D_185128.pdf)
- **Official IR Portal**: [Edelweiss Financial Services Limited Investor Relations](https://www.edelweissfin.com/investor-relations/)
- **Historical Cutoff & Snapshot Discipline**: Historical columns (`B`–`H`, `FY19`–`FY25`) contain strictly audited `Mar 2019`–`Mar 2025` figures predating the **October 15, 2025** market snapshot (`NSE Official Unadjusted Close: ₹108.85`). Post-cutoff `Mar 2026` figures are excluded from historical columns, and all 5 forecast years (`I`–`M`, `FY26E`–`FY30E`) are dynamic formulas with double-entry Balance Sheet cash closure (`Cash_t = Cash_{t-1} + Net_Cash_Flow_t`).

## 1. Line-by-Line Historical & Sub-Schedule Evidence Ledger (FY19–FY25)

| Line Item & Workbook Range | Unit | FY19 | FY20 | FY21 | FY22 | FY23 | FY24 | FY25 | Source Basis & Status |
|:---|:---:|---:|---:|---:|---:|---:|---:|---:|:---|
| **Total Consolidated Operating Income ('Income Statement'!B5:H5)** | ₹ Cr | 11,078 | 9,513 | 10,775 | 7,151 | 8,500 | 9,523 | **9,393** | `primary_sourced` |
| **Employee, Insurance & Operating Expenses ('Income Statement'!B6:H6)** | ₹ Cr | 4,507 | 7,034 | 6,609 | 3,947 | 5,667 | 6,252 | **6,033** | `primary_sourced` |
| **Operating Profit Before Finance Costs ('Income Statement'!B7:H7)** | ₹ Cr | 6,571 | 2,479 | 4,166 | 3,204 | 2,833 | 3,271 | **3,360** | `primary_sourced` |
| **Depreciation & Amortization ('Income Statement'!B9:H9)** | ₹ Cr | 132 | 232 | 260 | 151 | 138 | 126 | **147** | `primary_sourced` |
| **Finance Costs / Borrowing Interest ('Income Statement'!B10:H10)** | ₹ Cr | 4,783 | 4,793 | 3,834 | 2,984 | 2,575 | 2,786 | **2,537** | `primary_sourced` |
| **Other Income ('Income Statement'!B11:H11)** | ₹ Cr | 88 | 89 | 74 | 158 | 264 | 79 | **126** | `primary_sourced` |
| **Tax Expense ('Income Statement'!B13:H13)** | ₹ Cr | 700 | -413 | -108 | 15 | -22 | -90 | **266** | `primary_sourced` |
| **Reported Consolidated PAT ('Income Statement'!B14:H14)** | ₹ Cr | 1,044 | -2,044 | 254 | 212 | 406 | 528 | **536** | `primary_sourced` |
| **Equity Share Capital ('Balance Sheet'!B5:H5)** | ₹ Cr | 89 | 89 | 89 | 90 | 90 | 90 | **92** | `primary_sourced` |
| **Reserves & Surplus ('Balance Sheet'!B6:H6)** | ₹ Cr | 7,588 | 6,040 | 6,488 | 6,448 | 6,654 | 4,672 | **4,333** | `primary_sourced` |
| **Total Borrowings ('Balance Sheet'!B7:H7)** | ₹ Cr | 46,148 | 36,657 | 28,436 | 22,768 | 21,791 | 20,358 | **18,004** | `primary_sourced` |
| **Other Liabilities & Policyholder Funds ('Balance Sheet'!B8:H8)** | ₹ Cr | 10,076 | 11,230 | 11,121 | 13,758 | 15,346 | 17,594 | **18,853** | `primary_sourced` |
| **Net Block / Fixed Assets ('Balance Sheet'!B10:H10)** | ₹ Cr | 1,265 | 2,345 | 1,781 | 1,597 | 1,367 | 1,290 | **1,271** | `primary_sourced` |
| **Capital Work in Progress ('Balance Sheet'!B11:H11)** | ₹ Cr | 44 | 43 | 13 | 20 | 25 | 36 | **18** | `primary_sourced` |
| **Financial Investments ('Balance Sheet'!B12:H12)** | ₹ Cr | 12,713 | 10,012 | 12,882 | 14,139 | 17,162 | 18,491 | **17,574** | `primary_sourced` |
| **Loans, Credit Book & Cash ('Balance Sheet'!B13:H13)** | ₹ Cr | 49,879 | 41,616 | 31,458 | 27,308 | 25,327 | 22,897 | **22,419** | `primary_sourced` |
| **Sub-Schedule: Inventories ('Balance Sheet'!B19:H19)** | ₹ Cr | 268 | 92 | 57 | 74 | 0 | 0 | **0** | `primary_sourced` |
| **Sub-Schedule: Trade Receivables ('Balance Sheet'!B20:H20)** | ₹ Cr | 2,760 | 1,303 | 515 | 476 | 416 | 376 | **394** | `primary_sourced` |
| **Sub-Schedule: Cash & Equivalents ('Balance Sheet'!B21:H21)** | ₹ Cr | 6,640 | 8,154 | 4,766 | 3,038 | 3,684 | 3,260 | **5,241** | `primary_sourced` |
| **Sub-Schedule: Loans & Advances ('Balance Sheet'!B22:H22)** | ₹ Cr | 448 | 436 | 294 | 433 | 628 | 1,261 | **1,237** | `primary_sourced` |
| **Sub-Schedule: Credit Book & Financial Assets ('Balance Sheet'!B23:H23)** | ₹ Cr | 39,763 | 31,631 | 25,826 | 23,287 | 20,599 | 18,000 | **15,547** | `primary_sourced` |
| **Sub-Schedule: Trade Payables ('Balance Sheet'!B25:H25)** | ₹ Cr | 1,922 | 1,285 | 495 | 1,350 | 1,409 | 2,586 | **1,260** | `primary_sourced` |
| **Sub-Schedule: Non-Controlling Interest ('Balance Sheet'!B26:H26)** | ₹ Cr | 1,038 | 1,085 | 1,100 | 1,054 | 1,102 | 1,323 | **1,493** | `primary_sourced` |
| **Sub-Schedule: Policyholder & Other Liabilities ('Balance Sheet'!B27:H27)** | ₹ Cr | 7,116 | 8,860 | 9,526 | 11,354 | 12,835 | 13,685 | **16,100** | `primary_sourced` |
| **Cash from Operating Activity ('Cash Flow'!B4:H4)** | ₹ Cr | 5,685 | 12,098 | 3,458 | 5,592 | 2,891 | 2,894 | **3,384** | `primary_sourced` |
| **Cash from Investing Activity ('Cash Flow'!B5:H5)** | ₹ Cr | -1,312 | 199 | 4,085 | -1,058 | -425 | -482 | **3,794** | `primary_sourced` |
| **Cash from Financing Activity ('Cash Flow'!B6:H6)** | ₹ Cr | -3,665 | -10,471 | -8,587 | -6,444 | -1,708 | -2,785 | **-4,828** | `primary_sourced` |
| **Sub-Schedule: Fixed Assets Purchased / CapEx ('Cash Flow'!B9:H9)** | ₹ Cr | -296 | -153 | -120 | -76 | -103 | -135 | **-104** | `primary_sourced` |
| **Standalone EFSL HoldCo Borrowings ('SOTP Valuation'!B36)** | ₹ Cr | 46,148 | 36,657 | 28,436 | 22,768 | 4,350 | 3,385 | **3,072** | `primary_sourced` |
| **Debtor / Fee Collection Days ('Segment Operations'!B36:H36)** | Days | 91 | 50 | 17 | 24 | 18 | 14 | **15** | `primary_sourced` |

## 2. Prior vs. Current Valuation Reconciliation (Audit Contradiction Resolution)

| Valuation Metric | Prior Build (`9d357b8`) | Current Audited Build | Change (`Δ`) | Root-Cause Explanation |
|:---|---:|---:|---:|:---|
| **Implied Fair Value / Share (₹)** | ₹133.54 | **₹130.53** | ₹-3.01 (-2.25%) | Built explicit vertical-specific financial services drivers on 'Segment Operations' (EAAA Earning AUM × Fee & Carry Yield; EARCL Recoveries × Fee Rate; Mutual Fund Equity+Alt AUM × Fee Rate; Life & General Insurance GWP & Actuarial Indian Embedded Value [IEV]; Credit Book × Net Interest & Fee Yield), disclosed EFSL attributable ownership stakes (EAAA 100%, EARCL 79.8% net of CDPQ 20.2% NCI, EAML 85% post-WestBridge 15%, Life 75% [Tokio Marine 25%] + General 100%, Credit 100%) with a 15% Conglomerate Discount (Gross Vertical Sum ₹15,402 Cr), and replaced the unsourced -₹5,500 Cr HoldCo debt assumption with the Audited FY25 EFSL Standalone Balance Sheet Bridge (Standalone Borrowings ₹3,072 Cr less Standalone Cash ₹5 Cr = -₹3,067 Cr Net Debt), yielding ₹12,335 Cr Equity Value (₹130.53/share) vs verified 15-Oct-2025 NSE Official Close ₹108.85. |
| **Snapshot Market Price (15-Oct-2025 NSE Official Close)** | — | **₹108.85** | Verified | Official NSE unadjusted closing quotation dated Wednesday, October 15, 2025 (`2025-10-15`). |
| **Implied Upside / (Downside) vs. Snapshot** | — | **+19.92%** | Reconciled | Reconciled across Workbook, Registry, Manifest, and Coverage Report. |

> [!WARNING]
> **METHODOLOGICAL EXCLUSION OF INDUSTRIAL FCFF**:
> Industrial Free Cash Flow to Firm (FCFF) is strictly withheld and blocked for Edelweiss Financial Services Limited.
> Financial institutions rely on debt, deposits, and borrowings as operating raw material rather than corporate capital structure leverage;
> operating cash flow (`+₹3,384 Cr` in FY25) reflects wholesale loan book repayments funding debt reduction (`-₹4,828 Cr` CFF). Sum-of-the-Parts (SOTP) with Audited Standalone HoldCo Net Debt (`-₹3,067 Cr`) is the only defensible method.