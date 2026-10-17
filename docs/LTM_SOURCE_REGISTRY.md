# Primary Statutory & Evidence Source Registry: LTIMindtree Limited (LTM)

- **Target Legal Entity**: LTIMindtree Limited
- **BSE Scrip Code**: `540005` | **NSE Ticker**: `LTIM` | **ISIN**: `INE214T01019`
- **Reporting Basis & Scope**: FY19–FY21 reflect Pre-Merger Standalone/Consolidated LTI (Equity Cap ₹17–18 Cr); FY22 reflects Restated Combined Amalgamation Comparative (Equity Cap ₹30 Cr, +111.1% merger step); FY23–FY25 reflect Audited Merged LTIMindtree (effective Nov 14, 2022).
- **Local Primary Data Extract**: [`work/sources/LTM_consolidated.json`](../work/sources/LTM_consolidated.json) (`SHA-256: 7710eb4f648d664655b60b377e453983e0e2317aeb1783f72b488f8555cbde26`)
- **Local Statutory Sub-Schedules Cache**: [`work/sources/LTM_schedules.json`](../work/sources/LTM_schedules.json) (`SHA-256: fd04a7681cfa508f4eddf380a263840477a0d53d5bc2b1c9cd8d4b180ae98b41`)
- **Local 15-Oct-2025 Market & Empirical Beta Cache**: [`work/sources/market_and_beta_2025_10_15.json`](../work/sources/market_and_beta_2025_10_15.json) (`SHA-256: de04dac2fc4a73be3d596fe5aec37ff070e4dc0a4d2abe22b11b340206c3f823`)
- **Official BSE Filing PDF**: [BSE Statutory Filing PDF](https://www.bseindia.com/stockinfo/AnnPdfOpen.aspx?Pname=F6781B53_1516_46EF_B458_7E8518B12A74_175251.pdf)
- **Official IR Portal**: [LTIMindtree Limited Investor Relations](https://www.ltimindtree.com/investors/)
- **Historical Cutoff & Snapshot Discipline**: Historical columns (`B`–`H`, `FY19`–`FY25`) contain strictly audited `Mar 2019`–`Mar 2025` figures predating the **October 15, 2025** market snapshot (`NSE Official Unadjusted Close: ₹5,609.50`). Post-cutoff `Mar 2026` figures are excluded from historical columns, and all 5 forecast years (`I`–`M`, `FY26E`–`FY30E`) are dynamic formulas with double-entry Balance Sheet cash closure (`Cash_t = Cash_{t-1} + Net_Cash_Flow_t`).

## 1. Line-by-Line Historical & Sub-Schedule Evidence Ledger (FY19–FY25)

| Line Item & Workbook Range | Unit | FY19 | FY20 | FY21 | FY22 | FY23 | FY24 | FY25 | Source Basis & Status |
|:---|:---:|---:|---:|---:|---:|---:|---:|---:|:---|
| **Revenue from Operations ('Income Statement'!B5:H5)** | ₹ Cr | 9,446 | 10,879 | 12,370 | 26,109 | 33,183 | 35,517 | **38,008** | `restated_amalgamation` |
| **Employee & Operating Expenses ('Income Statement'!B6:H6)** | ₹ Cr | 7,562 | 8,849 | 9,645 | 20,860 | 27,075 | 29,130 | **31,513** | `restated_amalgamation` |
| **Operating Profit / EBITDA ('Income Statement'!B7:H7)** | ₹ Cr | 1,884 | 2,030 | 2,725 | 5,249 | 6,108 | 6,387 | **6,495** | `restated_amalgamation` |
| **Depreciation & Amortization ('Income Statement'!B9:H9)** | ₹ Cr | 147 | 273 | 332 | 597 | 723 | 819 | **992** | `restated_amalgamation` |
| **Finance Costs ('Income Statement'!B10:H10)** | ₹ Cr | 11 | 83 | 79 | 123 | 150 | 222 | **279** | `restated_amalgamation` |
| **Other Income ('Income Statement'!B11:H11)** | ₹ Cr | 302 | 329 | 274 | 766 | 557 | 702 | **990** | `restated_amalgamation` |
| **Tax Expense ('Income Statement'!B13:H13)** | ₹ Cr | 512 | 483 | 650 | 1,345 | 1,382 | 1,463 | **1,612** | `restated_amalgamation` |
| **Reported Consolidated PAT ('Income Statement'!B14:H14)** | ₹ Cr | 1,516 | 1,520 | 1,938 | 3,950 | 4,410 | 4,585 | **4,602** | `restated_amalgamation` |
| **Equity Share Capital ('Balance Sheet'!B5:H5)** | ₹ Cr | 17 | 17 | 18 | 30 | 30 | 30 | **30** | `primary_sourced` |
| **Reserves & Surplus ('Balance Sheet'!B6:H6)** | ₹ Cr | 4,876 | 5,387 | 7,286 | 14,258 | 16,562 | 19,988 | **22,669** | `primary_sourced` |
| **Borrowings / Lease Liabilities ('Balance Sheet'!B7:H7)** | ₹ Cr | 0 | 912 | 798 | 258 | 354 | 2,071 | **2,187** | `primary_sourced` |
| **Other Liabilities ('Balance Sheet'!B8:H8)** | ₹ Cr | 1,770 | 2,499 | 2,604 | 6,027 | 6,535 | 5,456 | **5,712** | `primary_sourced` |
| **Net Block / Fixed Assets ('Balance Sheet'!B10:H10)** | ₹ Cr | 930 | 1,920 | 1,906 | 3,473 | 3,679 | 4,981 | **5,285** | `primary_sourced` |
| **Capital Work in Progress ('Balance Sheet'!B11:H11)** | ₹ Cr | 12 | 59 | 66 | 503 | 902 | 551 | **681** | `primary_sourced` |
| **Treasury Investments ('Balance Sheet'!B12:H12)** | ₹ Cr | 1,740 | 2,219 | 3,730 | 6,048 | 5,458 | 8,744 | **9,845** | `primary_sourced` |
| **Other Assets ('Balance Sheet'!B13:H13)** | ₹ Cr | 3,981 | 4,617 | 5,004 | 10,549 | 13,442 | 13,269 | **14,787** | `primary_sourced` |
| **Sub-Schedule: Inventories ('Balance Sheet'!B19:H19)** | ₹ Cr | 0 | 0 | 4 | 4 | 3 | 3 | **3** | `primary_sourced` |
| **Sub-Schedule: Trade Receivables ('Balance Sheet'!B20:H20)** | ₹ Cr | 1,827 | 2,313 | 2,084 | 4,557 | 5,627 | 5,706 | **5,868** | `primary_sourced` |
| **Sub-Schedule: Cash & Equivalents ('Balance Sheet'!B21:H21)** | ₹ Cr | 415 | 525 | 759 | 2,914 | 2,939 | 2,816 | **3,588** | `primary_sourced` |
| **Sub-Schedule: Loans & Advances ('Balance Sheet'!B22:H22)** | ₹ Cr | 191 | 200 | 179 | 424 | 336 | 1,030 | **771** | `primary_sourced` |
| **Sub-Schedule: Trade Payables ('Balance Sheet'!B25:H25)** | ₹ Cr | 467 | 695 | 828 | 1,325 | 1,296 | 1,492 | **1,550** | `primary_sourced` |
| **Sub-Schedule: Non-Controlling Interest ('Balance Sheet'!B26:H26)** | ₹ Cr | 1 | 1 | 4 | 6 | 7 | 9 | **13** | `primary_sourced` |
| **Cash from Operating Activity ('Cash Flow'!B4:H4)** | ₹ Cr | 1,395 | 1,644 | 2,400 | 3,251 | 3,095 | 5,670 | **4,546** | `primary_sourced` |
| **Cash from Investing Activity ('Cash Flow'!B5:H5)** | ₹ Cr | -749 | -643 | -1,657 | -1,643 | -271 | -3,918 | **-1,729** | `primary_sourced` |
| **Cash from Financing Activity ('Cash Flow'!B6:H6)** | ₹ Cr | -594 | -890 | -509 | -1,680 | -1,932 | -2,269 | **-2,574** | `primary_sourced` |
| **Sub-Schedule: Fixed Assets Purchased / CapEx ('Cash Flow'!B9:H9)** | ₹ Cr | -159 | -242 | -257 | -1,052 | -935 | -839 | **-1,042** | `primary_sourced` |
| **Debtor Days ('IT Services Operations'!B19:H19)** | Days | 71 | 78 | 61 | 64 | 62 | 59 | **56** | `primary_sourced` |
| **Working Capital Days ('IT Services Operations'!B20:H20)** | Days | 53 | 42 | 33 | 37 | 44 | 36 | **39** | `primary_sourced` |

## 2. Prior vs. Current Valuation Reconciliation (Audit Contradiction Resolution)

| Valuation Metric | Prior Build (`9d357b8`) | Current Audited Build | Change (`Δ`) | Root-Cause Explanation |
|:---|---:|---:|---:|:---|
| **Implied Fair Value / Share (₹)** | ₹2,813.51 | **₹2,811.09** | ₹-2.42 (-0.09%) | Eliminated double-counted operational growth (where +0*$I$5 ignored headcount/utilization/billing) so USD Revenue ($M) = Headcount * Utilization * Realized Billing per Billed FTE / 1,000,000 and INR Revenue = USD Revenue * FX / 10. Replaced reverse-solved 0.950 beta (11.50% round WACC) with 5-peer Tier-1 IT Hamada Relevered Beta (0.9405 -> Ke = 11.95%, WACC = 11.87%, alongside Direct 2Y Weekly OLS Beta 1.2569) and removed the +₹2,500 Cr hardcoded cash bridge offset (Cash & Treasury = ₹3,588 Cr Cash + ₹9,845 Cr Investments = ₹13,433 Cr). Implied DCF value is ₹2,811.09/share vs verified 15-Oct-2025 NSE Official Close ₹5,609.50. |
| **Snapshot Market Price (15-Oct-2025 NSE Official Close)** | — | **₹5,609.50** | Verified | Official NSE unadjusted closing quotation dated Wednesday, October 15, 2025 (`2025-10-15`). |
| **Implied Upside / (Downside) vs. Snapshot** | — | **-49.89%** | Reconciled | Reconciled across Workbook, Registry, Manifest, and Coverage Report. |