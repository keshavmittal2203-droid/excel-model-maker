# Source Registry & Provenance Ledger: LTM (Template V8)

## 1. Executive Summary
- **Entity**: `LTM`
- **Active Deliverable**: `companies/LTM/template_v8/LTM-FullTemplate.xlsx`
- **Workbook SHA-256**: `5ec03127d0535f7b7e1d05219a96a4cb873b08647960a68273166e1b8ea41e7d`
- **Master Template**: `originals/Marksans Keshav.xlsx` (SHA-256 `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`)
- **Valuation Status**: `SUPPORTED`
- **Fair Value per Share**: `2,828.60` (linked from `'DCF 1st way'!E42` to `'Model ->'!C9`)
- **Evaluated Cost of Capital**: WACC = `12.48%` (Raw Beta = `1.1526`, Blume-Adjusted Beta = `1.1018`)
- **Beta Source**: Yahoo Finance 104-week weekly price regression ending 15-Oct-2025 (`work/sources/weekly_market_returns_104w_v2.json`, SHA-256 `19bd51f0ce6bcadfd4a7bfa7e0432532251bbcda6812486411300d6df9f16fb0`)
- **Risk-Free Rate**: 6.51% (Trading Economics dated prior-session proxy as of 14-Oct-2025 preceding 15-Oct-2025 cutoff)
- **Balance Sheet Integrity**: Historical FY15–FY25 and forecast FY26E reconcile to exact `0.00 Cr` (Row 40 = Row 38 - Row 16 = 0.00 Cr) with zero plugs.

## 2. Valuation Provenance & Acceptance Audit
Fully supported from allowlisted Screener.in statements, Yahoo Finance 104-week regression beta (1.1018), and Trading Economics 6.51% Rf proxy. Headcount and utilization metrics are quarterly presentation disclosures not on Screener.in and are excluded from model calculations.

## 3. Audited Source Entries
| Index | Metric | Stated / Evaluated Value | Status | Sourced Cell | Official Allowlisted URL |
|:-----:|:-------|:------------------------:|:------:|:------------:|:-------------------------|
| 1 | Revenue from Operations (FY19–FY25 Historical Series) | 38008 | `restated_amalgamation_comparative` | `'Income Statement'!B5:H5` | https://www.screener.in/company/LTM/consolidated/ |
| 2 | Employee & Operating Expenses (FY19–FY25 Historical Series) | 31513 | `restated_amalgamation_comparative` | `'Income Statement'!B6:H6` | https://www.screener.in/company/LTM/consolidated/ |
| 3 | Operating Profit / EBITDA (FY19–FY25 Historical Series) | 6495 | `restated_amalgamation_comparative` | `'Income Statement'!B7:H7` | https://www.screener.in/company/LTM/consolidated/ |
| 4 | Depreciation & Amortization (FY19–FY25 Historical Series) | 992 | `restated_amalgamation_comparative` | `'Income Statement'!B9:H9` | https://www.screener.in/company/LTM/consolidated/ |
| 5 | Finance Costs (FY19–FY25 Historical Series) | 279 | `restated_amalgamation_comparative` | `'Income Statement'!B10:H10` | https://www.screener.in/company/LTM/consolidated/ |
| 6 | Other Income (FY19–FY25 Historical Series) | 990 | `restated_amalgamation_comparative` | `'Income Statement'!B11:H11` | https://www.screener.in/company/LTM/consolidated/ |
| 7 | Tax Expense (FY19–FY25 Historical Series) | 1612 | `restated_amalgamation_comparative` | `'Income Statement'!B13:H13` | https://www.screener.in/company/LTM/consolidated/ |
| 8 | Reported Consolidated PAT (FY19–FY25 Historical Series) | 4602 | `restated_amalgamation_comparative` | `'Income Statement'!B14:H14` | https://www.screener.in/company/LTM/consolidated/ |
| 9 | Equity Share Capital (FY19–FY25 Historical Series) | 30 | `primary_sourced` | `'Balance Sheet'!B5:H5` | https://www.screener.in/company/LTM/consolidated/ |
| 10 | Reserves & Surplus (FY19–FY25 Historical Series) | 22669 | `primary_sourced` | `'Balance Sheet'!B6:H6` | https://www.screener.in/company/LTM/consolidated/ |
| 11 | Borrowings / Lease Liabilities (FY19–FY25 Historical Series) | 2187 | `primary_sourced` | `'Balance Sheet'!B7:H7` | https://www.screener.in/company/LTM/consolidated/ |
| 12 | Other Liabilities (FY19–FY25 Historical Series) | 5712 | `primary_sourced` | `'Balance Sheet'!B8:H8` | https://www.screener.in/company/LTM/consolidated/ |
| 13 | Net Block / Fixed Assets (FY19–FY25 Historical Series) | 5285 | `primary_sourced` | `'Balance Sheet'!B10:H10` | https://www.screener.in/company/LTM/consolidated/ |
| 14 | Capital Work in Progress (FY19–FY25 Historical Series) | 681 | `primary_sourced` | `'Balance Sheet'!B11:H11` | https://www.screener.in/company/LTM/consolidated/ |
| 15 | Treasury Investments (FY19–FY25 Historical Series) | 9845 | `primary_sourced` | `'Balance Sheet'!B12:H12` | https://www.screener.in/company/LTM/consolidated/ |
| 16 | Other Assets (FY19–FY25 Historical Series) | 14787 | `primary_sourced` | `'Balance Sheet'!B13:H13` | https://www.screener.in/company/LTM/consolidated/ |
| 17 | Sub-Schedule: Inventories (FY19–FY25 Historical Series) | 3 | `primary_sourced` | `'Balance Sheet'!B19:H19` | https://www.screener.in/company/LTM/consolidated/ |
| 18 | Sub-Schedule: Trade Receivables (FY19–FY25 Historical Series) | 5868 | `primary_sourced` | `'Balance Sheet'!B20:H20` | https://www.screener.in/company/LTM/consolidated/ |
| 19 | Sub-Schedule: Cash & Equivalents (FY19–FY25 Historical Series) | 3588 | `primary_sourced` | `'Balance Sheet'!B21:H21` | https://www.screener.in/company/LTM/consolidated/ |
| 20 | Sub-Schedule: Loans & Advances (FY19–FY25 Historical Series) | 771 | `primary_sourced` | `'Balance Sheet'!B22:H22` | https://www.screener.in/company/LTM/consolidated/ |
| 21 | Sub-Schedule: Trade Payables (FY19–FY25 Historical Series) | 1550 | `primary_sourced` | `'Balance Sheet'!B25:H25` | https://www.screener.in/company/LTM/consolidated/ |
| 22 | Sub-Schedule: Non-Controlling Interest (FY19–FY25 Historical Series) | 13 | `primary_sourced` | `'Balance Sheet'!B26:H26` | https://www.screener.in/company/LTM/consolidated/ |
| 23 | Cash from Operating Activity (FY19–FY25 Historical Series) | 4546 | `primary_sourced` | `'Cash Flow'!B4:H4` | https://www.screener.in/company/LTM/consolidated/ |
| 24 | Cash from Investing Activity (FY19–FY25 Historical Series) | -1729 | `primary_sourced` | `'Cash Flow'!B5:H5` | https://www.screener.in/company/LTM/consolidated/ |
| 25 | Cash from Financing Activity (FY19–FY25 Historical Series) | -2574 | `primary_sourced` | `'Cash Flow'!B6:H6` | https://www.screener.in/company/LTM/consolidated/ |
| 26 | Sub-Schedule: Fixed Assets Purchased / CapEx (FY19–FY25 Historical Series) | -1042 | `primary_sourced` | `'Cash Flow'!B9:H9` | https://www.screener.in/company/LTM/consolidated/ |
| 27 | Debtor Days (FY19–FY25 Historical Series) | 56 | `primary_sourced` | `'IT Services Operations'!B19:H19` | https://www.screener.in/company/LTM/consolidated/ |
| 28 | Working Capital Days (FY19–FY25 Historical Series) | 39 | `primary_sourced` | `'IT Services Operations'!B20:H20` | https://www.screener.in/company/LTM/consolidated/ |
| 29 | Diluted Equity Share Count (FY25 Close) | 29.62 | `primary_sourced` | `'Cover & Notes'!B12` | https://www.screener.in/company/LTM/consolidated/ |
| 30 | Historical Market Snapshot Price (15-Oct-2025 NSE Official Unadjusted Close) | 5609.5 | `historical_market_snapshot` | `'Cover & Notes'!B10` | https://www.nseindia.com |
| 31 | 2-Year Weekly OLS Regression Beta vs Nifty 50 (^NSEI) (Raw & Blume-Adjusted) | 1.1018 | `empirical_regression_derived` | `'Beta'!K13 (Adjusted) / 'Beta'!K7 (Raw)` | https://finance.yahoo.com |
| 32 | 104-Week Empirical Weekly Market Returns & OLS Regression Dataset | 104 | `empirical_market_dataset` | `'Beta'!B8:G112` | https://finance.yahoo.com |
| 33 | Primary Valuation Output (Unlevered Free Cash Flow (FCFF) DCF & CAPM WACC) | 2828.6 | `supported` | `'DCF 1st way'!E42` | None (Unavailable on allowed sources) |
