# Source Registry & Provenance Ledger: CLEAN (Template V8)

## 1. Executive Summary
- **Entity**: `CLEAN`
- **Active Deliverable**: `companies/CLEAN/template_v8/CLEAN-FullTemplate.xlsx`
- **Workbook SHA-256**: `b868ea5e0b1350219ebd28a1e276d135898abf4d5e87e430ea5c05c5118e1796`
- **Master Template**: `originals/Marksans Keshav.xlsx` (SHA-256 `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`)
- **Valuation Status**: `SUPPORTED`
- **Fair Value per Share**: `408.55` (linked from `'DCF 1st way'!E42` to `'Model ->'!C9`)
- **Evaluated Cost of Capital**: WACC = `13.19%` (Raw Beta = `1.3212`, Blume-Adjusted Beta = `1.2141`)
- **Beta Source**: Yahoo Finance 104-week weekly price regression ending 15-Oct-2025 (`work/sources/weekly_market_returns_104w_v2.json`, SHA-256 `19bd51f0ce6bcadfd4a7bfa7e0432532251bbcda6812486411300d6df9f16fb0`)
- **Risk-Free Rate**: 6.51% (Trading Economics dated prior-session proxy as of 14-Oct-2025 preceding 15-Oct-2025 cutoff)
- **Balance Sheet Integrity**: Historical FY15–FY25 and forecast FY26E reconcile to exact `0.00 Cr` (Row 40 = Row 38 - Row 16 = 0.00 Cr) with zero plugs.

## 2. Valuation Provenance & Acceptance Audit
Fully supported from allowlisted Screener.in statements, Yahoo Finance 104-week regression beta (1.2141), and Trading Economics 6.51% Rf proxy. Unverified 10,500 MTPA capacity figure is removed from forecast calculation and labels; forecast is driven strictly by explicit analyst revenue growth and reinvestment assumptions applied to Screener historical financials.

## 3. Audited Source Entries
| Index | Metric | Stated / Evaluated Value | Status | Sourced Cell | Official Allowlisted URL |
|:-----:|:-------|:------------------------:|:------:|:------------:|:-------------------------|
| 1 | Revenue from Operations (FY19–FY25 Historical Series) | 967 | `primary_sourced` | `'Income Statement'!B5:H5` | https://www.screener.in/company/CLEAN/consolidated/ |
| 2 | Raw Material, Power & Operating Expenses (FY19–FY25 Historical Series) | 579 | `primary_sourced` | `'Income Statement'!B6:H6` | https://www.screener.in/company/CLEAN/consolidated/ |
| 3 | Operating Profit / EBITDA (FY19–FY25 Historical Series) | 388 | `primary_sourced` | `'Income Statement'!B7:H7` | https://www.screener.in/company/CLEAN/consolidated/ |
| 4 | Depreciation & Amortization (FY19–FY25 Historical Series) | 69 | `primary_sourced` | `'Income Statement'!B9:H9` | https://www.screener.in/company/CLEAN/consolidated/ |
| 5 | Finance Costs (FY19–FY25 Historical Series) | 1 | `primary_sourced` | `'Income Statement'!B10:H10` | https://www.screener.in/company/CLEAN/consolidated/ |
| 6 | Other Income (FY19–FY25 Historical Series) | 39 | `primary_sourced` | `'Income Statement'!B11:H11` | https://www.screener.in/company/CLEAN/consolidated/ |
| 7 | Tax Expense (FY19–FY25 Historical Series) | 93 | `primary_sourced` | `'Income Statement'!B13:H13` | https://www.screener.in/company/CLEAN/consolidated/ |
| 8 | Reported Consolidated PAT (FY19–FY25 Historical Series) | 264 | `primary_sourced` | `'Income Statement'!B14:H14` | https://www.screener.in/company/CLEAN/consolidated/ |
| 9 | Equity Share Capital (FY19–FY25 Historical Series) | 11 | `primary_sourced` | `'Balance Sheet'!B5:H5` | https://www.screener.in/company/CLEAN/consolidated/ |
| 10 | Reserves & Surplus (FY19–FY25 Historical Series) | 1406 | `primary_sourced` | `'Balance Sheet'!B6:H6` | https://www.screener.in/company/CLEAN/consolidated/ |
| 11 | Borrowings (FY19–FY25 Historical Series) | 2 | `primary_sourced` | `'Balance Sheet'!B7:H7` | https://www.screener.in/company/CLEAN/consolidated/ |
| 12 | Other Liabilities (FY19–FY25 Historical Series) | 181 | `primary_sourced` | `'Balance Sheet'!B8:H8` | https://www.screener.in/company/CLEAN/consolidated/ |
| 13 | Net Block / Fixed Assets (FY19–FY25 Historical Series) | 715 | `primary_sourced` | `'Balance Sheet'!B10:H10` | https://www.screener.in/company/CLEAN/consolidated/ |
| 14 | Capital Work in Progress (FY19–FY25 Historical Series) | 30 | `primary_sourced` | `'Balance Sheet'!B11:H11` | https://www.screener.in/company/CLEAN/consolidated/ |
| 15 | Treasury Investments (FY19–FY25 Historical Series) | 381 | `primary_sourced` | `'Balance Sheet'!B12:H12` | https://www.screener.in/company/CLEAN/consolidated/ |
| 16 | Other Assets (FY19–FY25 Historical Series) | 474 | `primary_sourced` | `'Balance Sheet'!B13:H13` | https://www.screener.in/company/CLEAN/consolidated/ |
| 17 | Sub-Schedule: Inventories (FY19–FY25 Historical Series) | 149 | `primary_sourced` | `'Balance Sheet'!B19:H19` | https://www.screener.in/company/CLEAN/consolidated/ |
| 18 | Sub-Schedule: Trade Receivables (FY19–FY25 Historical Series) | 203 | `primary_sourced` | `'Balance Sheet'!B20:H20` | https://www.screener.in/company/CLEAN/consolidated/ |
| 19 | Sub-Schedule: Cash & Equivalents (FY19–FY25 Historical Series) | 20 | `primary_sourced` | `'Balance Sheet'!B21:H21` | https://www.screener.in/company/CLEAN/consolidated/ |
| 20 | Sub-Schedule: Loans & Advances (FY19–FY25 Historical Series) | 4 | `primary_sourced` | `'Balance Sheet'!B22:H22` | https://www.screener.in/company/CLEAN/consolidated/ |
| 21 | Sub-Schedule: Trade Payables (FY19–FY25 Historical Series) | 99 | `primary_sourced` | `'Balance Sheet'!B25:H25` | https://www.screener.in/company/CLEAN/consolidated/ |
| 22 | Sub-Schedule: Non-Controlling Interest (FY19–FY25 Historical Series) | 0 | `primary_sourced` | `'Balance Sheet'!B26:H26` | https://www.screener.in/company/CLEAN/consolidated/ |
| 23 | Cash from Operating Activity (FY19–FY25 Historical Series) | 214 | `primary_sourced` | `'Cash Flow'!B4:H4` | https://www.screener.in/company/CLEAN/consolidated/ |
| 24 | Cash from Investing Activity (FY19–FY25 Historical Series) | -152 | `primary_sourced` | `'Cash Flow'!B5:H5` | https://www.screener.in/company/CLEAN/consolidated/ |
| 25 | Cash from Financing Activity (FY19–FY25 Historical Series) | -53 | `primary_sourced` | `'Cash Flow'!B6:H6` | https://www.screener.in/company/CLEAN/consolidated/ |
| 26 | Sub-Schedule: Fixed Assets Purchased / CapEx (FY19–FY25 Historical Series) | -181 | `primary_sourced` | `'Cash Flow'!B9:H9` | https://www.screener.in/company/CLEAN/consolidated/ |
| 27 | Debtor Days (FY19–FY25 Historical Series) | 77 | `primary_sourced` | `'Manufacturing & Capacity'!B25:H25` | https://www.screener.in/company/CLEAN/consolidated/ |
| 28 | Inventory Days (FY19–FY25 Historical Series) | 156 | `primary_sourced` | `'Manufacturing & Capacity'!B26:H26` | https://www.screener.in/company/CLEAN/consolidated/ |
| 29 | Days Payable (FY19–FY25 Historical Series) | 103 | `primary_sourced` | `'Manufacturing & Capacity'!B27:H27` | https://www.screener.in/company/CLEAN/consolidated/ |
| 30 | Diluted Equity Share Count (FY25 Close) | 10.63 | `primary_sourced` | `'Cover & Notes'!B12` | https://www.screener.in/company/CLEAN/consolidated/ |
| 31 | Historical Market Snapshot Price (15-Oct-2025 NSE Official Unadjusted Close) | 1041.4 | `historical_market_snapshot` | `'Cover & Notes'!B10` | https://www.nseindia.com |
| 32 | 2-Year Weekly OLS Regression Beta vs Nifty 50 (^NSEI) (Raw & Blume-Adjusted) | 1.2141 | `empirical_regression_derived` | `'Beta'!K13 (Adjusted) / 'Beta'!K7 (Raw)` | https://finance.yahoo.com |
| 33 | 104-Week Empirical Weekly Market Returns & OLS Regression Dataset | 104 | `empirical_market_dataset` | `'Beta'!B8:G112` | https://finance.yahoo.com |
| 34 | Primary Valuation Output (Manufacturing FCFF DCF & CAPM WACC) | 408.55 | `supported` | `'DCF 1st way'!E42` | None (Unavailable on allowed sources) |
