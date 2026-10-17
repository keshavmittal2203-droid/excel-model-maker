# Primary Statutory & Evidence Source Registry: Mahindra & Mahindra Limited (MM)

- **Target Legal Entity**: Mahindra & Mahindra Limited
- **BSE Scrip Code**: `500520` | **NSE Ticker**: `M&M` | **ISIN**: `INE146A01017`
- **Reporting Basis & Scope**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; SOTP uses Audited FY25 M&M Standalone Core Operating Profit ₹18,416 Cr [Auto ₹11,700 Cr + Farm ₹6,716 Cr] and Audited FY25 Standalone Net Cash +₹9,109 Cr to prevent MMFS NBFC debt double-counting).
- **Local Primary Data Extract**: [`work/sources/MM_consolidated.json`](../work/sources/MM_consolidated.json) (`SHA-256: d4a77cd285356e31bd99ca8530ca43aef943a77da2b86a5330b05d1418956c68`)
- **Local Statutory Sub-Schedules Cache**: [`work/sources/MM_schedules.json`](../work/sources/MM_schedules.json) (`SHA-256: 844d2059b1477e16cb808cea2dd20bd7dc4b279b94a176d23fa491d7527c263d`)
- **Local 15-Oct-2025 Market & Empirical Beta Cache**: [`work/sources/market_and_beta_2025_10_15.json`](../work/sources/market_and_beta_2025_10_15.json) (`SHA-256: de04dac2fc4a73be3d596fe5aec37ff070e4dc0a4d2abe22b11b340206c3f823`)
- **Local 104-Week Empirical Weekly Market Returns Cache**: [`work/sources/weekly_market_returns_104w.json`](../work/sources/weekly_market_returns_104w.json) (`SHA-256: 109e1f9726130ffc6fc9121e941b6bedb66a90b3a8d249d8beca40bc0bd50eef`)
- **Official BSE Filing PDF**: [BSE Statutory Filing PDF](https://www.bseindia.com/stockinfo/AnnPdfOpen.aspx?Pname=6F653FFC_204D_4551_A741_4F4EE3C56EBF_182346.pdf)
- **Official IR Portal**: [Mahindra & Mahindra Limited Investor Relations](https://www.mahindra.com/resources/investor-reports)
- **Historical Cutoff & Snapshot Discipline**: Historical columns (`B`–`H`, `FY19`–`FY25`) contain strictly audited `Mar 2019`–`Mar 2025` figures predating the **October 15, 2025** market snapshot (`NSE Official Unadjusted Close: ₹3,497.20`). Post-cutoff `Mar 2026` figures are excluded from historical columns, and all 5 forecast years (`I`–`M`, `FY26E`–`FY30E`) are dynamic formulas with double-entry Balance Sheet cash closure (`Cash_t = Cash_{t-1} + Net_Cash_Flow_t`).

## 1. Line-by-Line Historical & Sub-Schedule Evidence Ledger (FY19–FY25)

| Line Item & Workbook Range | Unit | FY19 | FY20 | FY21 | FY22 | FY23 | FY24 | FY25 | Source Basis & Status |
|:---|:---:|---:|---:|---:|---:|---:|---:|---:|:---|
| **Revenue from Operations ('Income Statement'!B5:H5)** | ₹ Cr | 104,721 | 75,382 | 74,278 | 90,171 | 121,362 | 139,078 | **159,211** | `primary_sourced` |
| **Total Operating Expenditure ('Income Statement'!B6:H6)** | ₹ Cr | 89,514 | 65,224 | 62,791 | 75,488 | 101,037 | 114,186 | **128,693** | `primary_sourced` |
| **Operating Profit / Consolidated EBITDA ('Income Statement'!B7:H7)** | ₹ Cr | 15,207 | 10,158 | 11,487 | 14,683 | 20,325 | 24,892 | **30,518** | `primary_sourced` |
| **Depreciation & Amortization ('Income Statement'!B9:H9)** | ₹ Cr | 3,991 | 3,367 | 3,378 | 3,508 | 4,357 | 4,724 | **6,074** | `primary_sourced` |
| **Finance Costs / Interest ('Income Statement'!B10:H10)** | ₹ Cr | 5,021 | 6,021 | 6,102 | 5,018 | 5,830 | 7,488 | **9,083** | `primary_sourced` |
| **Other Income ('Income Statement'!B11:H11)** | ₹ Cr | 2,676 | 885 | 1,152 | 3,204 | 3,922 | 3,298 | **3,718** | `primary_sourced` |
| **Tax Expense ('Income Statement'!B13:H13)** | ₹ Cr | 2,854 | 1,976 | 1,647 | 2,108 | 2,686 | 3,708 | **5,006** | `primary_sourced` |
| **Reported Consolidated PAT ('Income Statement'!B14:H14)** | ₹ Cr | 6,017 | -321 | 1,512 | 7,253 | 11,374 | 12,270 | **14,073** | `primary_sourced` |
| **Equity Share Capital ('Balance Sheet'!B4:H4)** | ₹ Cr | 544 | 554 | 555 | 556 | 557 | 557 | **558** | `primary_sourced` |
| **Reserves & Surplus ('Balance Sheet'!B5:H5)** | ₹ Cr | 39,439 | 39,415 | 41,027 | 46,567 | 55,809 | 65,633 | **76,481** | `primary_sourced` |
| **Total Borrowings ('Balance Sheet'!B6:H6)** | ₹ Cr | 70,848 | 82,092 | 80,625 | 77,605 | 92,647 | 108,647 | **129,025** | `primary_sourced` |
| **Other Liabilities ('Balance Sheet'!B7:H7)** | ₹ Cr | 51,057 | 43,655 | 42,762 | 47,661 | 55,270 | 59,884 | **69,950** | `primary_sourced` |
| **Net Block / Fixed Assets ('Balance Sheet'!B10:H10)** | ₹ Cr | 28,983 | 29,689 | 21,380 | 26,018 | 27,140 | 28,129 | **34,636** | `primary_sourced` |
| **Capital Work in Progress ('Balance Sheet'!B11:H11)** | ₹ Cr | 4,760 | 6,856 | 7,873 | 6,703 | 3,969 | 8,039 | **8,298** | `primary_sourced` |
| **Investments ('Balance Sheet'!B12:H12)** | ₹ Cr | 18,268 | 19,210 | 28,778 | 30,060 | 35,272 | 35,208 | **41,309** | `primary_sourced` |
| **Other Assets ('Balance Sheet'!B13:H13)** | ₹ Cr | 109,877 | 109,961 | 106,938 | 109,608 | 137,902 | 163,345 | **191,771** | `primary_sourced` |
| **Sub-Schedule: Inventories ('Balance Sheet'!B19:H19)** | ₹ Cr | 12,200 | 11,112 | 9,615 | 11,596 | 16,855 | 18,590 | **20,306** | `primary_sourced` |
| **Sub-Schedule: Trade Receivables ('Balance Sheet'!B20:H20)** | ₹ Cr | 8,678 | 6,928 | 6,008 | 6,374 | 7,028 | 7,459 | **8,456** | `primary_sourced` |
| **Sub-Schedule: Cash & Equivalents ('Balance Sheet'!B21:H21)** | ₹ Cr | 9,329 | 7,911 | 12,948 | 11,116 | 11,273 | 12,028 | **13,606** | `primary_sourced` |
| **Sub-Schedule: Loans & Advances / MMFS Book ('Balance Sheet'!B22:H22)** | ₹ Cr | 70,924 | 73,862 | 66,700 | 67,890 | 89,882 | 111,695 | **132,387** | `primary_sourced` |
| **Sub-Schedule: Trade Payables ('Balance Sheet'!B25:H25)** | ₹ Cr | 20,885 | 14,981 | 15,506 | 19,037 | 23,836 | 25,605 | **29,972** | `primary_sourced` |
| **Sub-Schedule: Non-Controlling Interest ('Balance Sheet'!B26:H26)** | ₹ Cr | 8,078 | 7,420 | 9,095 | 9,568 | 10,796 | 11,596 | **13,103** | `primary_sourced` |
| **Cash from Operating Activity ('Cash Flow'!B4:H4)** | ₹ Cr | -4,347 | -1,457 | 17,909 | 9,248 | -7,074 | -5,630 | **3,176** | `primary_sourced` |
| **Cash from Investing Activity ('Cash Flow'!B5:H5)** | ₹ Cr | -7,174 | -6,870 | -19,686 | -3,252 | -8,866 | -5,615 | **-18,616** | `primary_sourced` |
| **Cash from Financing Activity ('Cash Flow'!B6:H6)** | ₹ Cr | 13,194 | 6,933 | 406 | -5,883 | 15,946 | 12,281 | **15,834** | `primary_sourced` |
| **Sub-Schedule: Fixed Assets Purchased / CapEx ('Cash Flow'!B10:H10)** | ₹ Cr | -7,339 | -6,896 | -5,959 | -6,040 | -6,305 | -9,679 | **-11,889** | `primary_sourced` |
| **Standalone Revenue from Operations ('Segment Operations'!B19:H19)** | ₹ Cr | 52,848 | 44,866 | 44,630 | 57,787 | 84,960 | 98,763 | **116,484** | `primary_sourced` |
| **Standalone Core Operating Profit / EBITDA ('Segment Operations'!B20:H20)** | ₹ Cr | 7,530 | 6,351 | 6,957 | 7,027 | 10,442 | 13,135 | **18,416** | `primary_sourced` |
| **Debtor Days ('Segment Operations'!B29:H29)** | Days | 30 | 34 | 30 | 26 | 21 | 20 | **19** | `primary_sourced` |
| **Inventory Days ('Segment Operations'!B30:H30)** | Days | 74 | 106 | 91 | 82 | 83 | 81 | **78** | `primary_sourced` |
| **Days Payable ('Segment Operations'!B31:H31)** | Days | 127 | 142 | 147 | 134 | 118 | 112 | **118** | `primary_sourced` |

## 2. Prior vs. Current Valuation Reconciliation (Audit Contradiction Resolution)

| Valuation Metric | Prior Build (`9d357b8`) | Current Audited Build | Change (`Δ`) | Root-Cause Explanation |
|:---|---:|---:|---:|:---|
| **Implied Fair Value / Share (₹)** | ₹2,929.98 | **₹2,742.97** | ₹-187.01 (-6.38%) | Eliminated the backwards-solved ₹242,400 Cr target preservation (15,150 × 16x vs 20,200 × 12x) and replaced the unsourced -₹12,000 Cr standalone net debt plug with the Audited FY25 M&M Standalone Balance Sheet Bridge (Standalone Borrowings ₹1,682 Cr less Standalone Cash & Treasury ₹10,791 Cr = -₹9,109 Cr Net Debt / +₹9,109 Cr Standalone Net Cash). Connected FY26E Standalone Auto Core EBITDA (₹13,667 Cr @ 14.0x = ₹1,91,338 Cr) and Farm Core EBITDA (₹7,403 Cr @ 13.0x = ₹96,239 Cr) to Volume × ASP drivers, updated MMFS (52.16% of ₹36,305 Cr @ ₹293.85 15-Oct-2025 close × 80% = ₹15,149.35 Cr) and TechM (25.48% of ₹1,42,836 Cr @ ₹1,459.00 15-Oct-2025 close × 80% = ₹29,115.69 Cr), yielding ₹3,40,951.04 Cr Equity Value (₹2,742.97/share). |
| **Snapshot Market Price (15-Oct-2025 NSE Official Close)** | — | **₹3,497.20** | Verified | Official NSE unadjusted closing quotation dated Wednesday, October 15, 2025 (`2025-10-15`). |
| **Implied Upside / (Downside) vs. Snapshot** | — | **-21.57%** | Reconciled | Reconciled across Workbook, Registry, Manifest, and Coverage Report. |

## 3. Empirical 104-Week OLS Regression Diagnostics (vs. Nifty 50, N=104 Weeks)

| Statistical Metric / Regression Parameter | Value | Unit / Format | Benchmark / Regression Methodology | Statistical Interpretation |
|:---|---:|:---:|:---|:---|
| **Raw OLS Beta Slope** | **1.7638** | Ratio | Covariance(R_i, R_m) / Variance(R_m) | Direct empirical slope over 104 paired weekly returns |
| **Blume Adjusted Beta** | **1.5092** | Ratio | 2/3 Raw Beta + 1/3 * 1.0000 | Institutional mean-reverting adjusted beta |
| **Weekly Covariance with Nifty 50** | 0.00078712 | Covariance | Sample covariance over 104 weekly return bars | Measure of systematic co-movement with broad index |
| **Nifty 50 Weekly Return Variance** | 0.00044626 | Variance | Sample variance of Nifty 50 weekly returns | Annualized benchmark volatility ~15.2% |
| **Pearson Correlation (R)** | 0.6388 | Correlation | Cov(R_i, R_m) / [Std(R_i) * Std(R_m)] | Systematic directional co-movement |
| **Coefficient of Determination (R²)** | 0.4081 | R-Squared | Explained return variance / Total variance | Proportion of systematic variance explained |
| **Standard Error of Regression (SE)** | 0.0267 | Std Error | Residual standard error of regression line | Idiosyncratic return dispersion |
| **Regression t-Statistic** | 8.38 | t-Stat | Beta / SE_Beta (p < 0.0001) | Statistically significant at >99.9% confidence |
| **Peer Cohort Reference** | — | Peer Median | Auto & Tractor Peer Median: 1.1244 (Maruti 0.7740, Bajaj 1.1643, Eicher 1.0845, Escorts 1.4160) | Contextual industry benchmark |