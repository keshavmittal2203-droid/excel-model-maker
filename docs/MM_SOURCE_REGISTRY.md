# Source Registry & Provenance Ledger: MM (Template V8)

## 1. Executive Summary
- **Entity**: `MM`
- **Active Deliverable**: `companies/MM/template_v8/MM-FullTemplate.xlsx`
- **Workbook SHA-256**: `3705342cc188f489628ad8d3b0345c932cc0ec7fecaec2f0458d08b766f9ea97`
- **Master Template**: `originals/Marksans Keshav.xlsx` (SHA-256 `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`)
- **Valuation Status**: `UNAVAILABLE`
- **Fair Value per Share**: `Unavailable` (linked from `'DCF 1st way'!E42` to `'Model ->'!C9`)
- **Evaluated Cost of Capital**: WACC = `11.97%` (Raw Beta = `1.4746`, Blume-Adjusted Beta = `1.3164`)
- **Beta Source**: Yahoo Finance 104-week weekly price regression ending 15-Oct-2025 (`work/sources/weekly_market_returns_104w_v2.json`, SHA-256 `19bd51f0ce6bcadfd4a7bfa7e0432532251bbcda6812486411300d6df9f16fb0`)
- **Risk-Free Rate**: 6.51% (Trading Economics dated prior-session proxy as of 14-Oct-2025 preceding 15-Oct-2025 cutoff)
- **Balance Sheet Integrity**: Historical FY15–FY25 and forecast FY26E reconcile to exact `0.00 Cr` (Row 40 = Row 38 - Row 16 = 0.00 Cr) with zero plugs.

## 2. Valuation Provenance & Acceptance Audit
Valuation is marked Unavailable: Factual segment EBITDA inputs (Auto, Farm) and Standalone Net Cash are statutory filing items not published on allowlisted sources (Screener.in). Under the locked allowlist rule, they cannot be used in calculation or relabeled as analyst assumptions.

## 3. Audited Source Entries
| Index | Metric | Stated / Evaluated Value | Status | Sourced Cell | Official Allowlisted URL |
|:-----:|:-------|:------------------------:|:------:|:------------:|:-------------------------|
| 1 | Revenue from Operations (FY19–FY25 Historical Series) | 159211 | `primary_sourced` | `'Income Statement'!B5:H5` | https://www.screener.in/company/M&M/consolidated/ |
| 2 | Total Operating Expenditure (FY19–FY25 Historical Series) | 128693 | `primary_sourced` | `'Income Statement'!B6:H6` | https://www.screener.in/company/M&M/consolidated/ |
| 3 | Operating Profit / Consolidated EBITDA (FY19–FY25 Historical Series) | 30518 | `primary_sourced` | `'Income Statement'!B7:H7` | https://www.screener.in/company/M&M/consolidated/ |
| 4 | Depreciation & Amortization (FY19–FY25 Historical Series) | 6074 | `primary_sourced` | `'Income Statement'!B9:H9` | https://www.screener.in/company/M&M/consolidated/ |
| 5 | Finance Costs / Interest (FY19–FY25 Historical Series) | 9083 | `primary_sourced` | `'Income Statement'!B10:H10` | https://www.screener.in/company/M&M/consolidated/ |
| 6 | Other Income (FY19–FY25 Historical Series) | 3718 | `primary_sourced` | `'Income Statement'!B11:H11` | https://www.screener.in/company/M&M/consolidated/ |
| 7 | Tax Expense (FY19–FY25 Historical Series) | 5006 | `primary_sourced` | `'Income Statement'!B13:H13` | https://www.screener.in/company/M&M/consolidated/ |
| 8 | Reported Consolidated PAT (FY19–FY25 Historical Series) | 14073 | `primary_sourced` | `'Income Statement'!B14:H14` | https://www.screener.in/company/M&M/consolidated/ |
| 9 | Equity Share Capital (FY19–FY25 Historical Series) | 558 | `primary_sourced` | `'Balance Sheet'!B4:H4` | https://www.screener.in/company/M&M/consolidated/ |
| 10 | Reserves & Surplus (FY19–FY25 Historical Series) | 76481 | `primary_sourced` | `'Balance Sheet'!B5:H5` | https://www.screener.in/company/M&M/consolidated/ |
| 11 | Total Borrowings (FY19–FY25 Historical Series) | 129025 | `primary_sourced` | `'Balance Sheet'!B6:H6` | https://www.screener.in/company/M&M/consolidated/ |
| 12 | Other Liabilities (FY19–FY25 Historical Series) | 69950 | `primary_sourced` | `'Balance Sheet'!B7:H7` | https://www.screener.in/company/M&M/consolidated/ |
| 13 | Net Block / Fixed Assets (FY19–FY25 Historical Series) | 34636 | `primary_sourced` | `'Balance Sheet'!B10:H10` | https://www.screener.in/company/M&M/consolidated/ |
| 14 | Capital Work in Progress (FY19–FY25 Historical Series) | 8298 | `primary_sourced` | `'Balance Sheet'!B11:H11` | https://www.screener.in/company/M&M/consolidated/ |
| 15 | Investments (FY19–FY25 Historical Series) | 41309 | `primary_sourced` | `'Balance Sheet'!B12:H12` | https://www.screener.in/company/M&M/consolidated/ |
| 16 | Other Assets (FY19–FY25 Historical Series) | 191771 | `primary_sourced` | `'Balance Sheet'!B13:H13` | https://www.screener.in/company/M&M/consolidated/ |
| 17 | Sub-Schedule: Inventories (FY19–FY25 Historical Series) | 20306 | `unavailable_on_allowed_sources` | `'Balance Sheet'!B19:H19` | None (Unavailable on allowed sources) |
| 18 | Sub-Schedule: Trade Receivables (FY19–FY25 Historical Series) | 8456 | `unavailable_on_allowed_sources` | `'Balance Sheet'!B20:H20` | None (Unavailable on allowed sources) |
| 19 | Sub-Schedule: Cash & Equivalents (FY19–FY25 Historical Series) | 13606 | `unavailable_on_allowed_sources` | `'Balance Sheet'!B21:H21` | None (Unavailable on allowed sources) |
| 20 | Sub-Schedule: Loans & Advances / MMFS Book (FY19–FY25 Historical Series) | 132387 | `unavailable_on_allowed_sources` | `'Balance Sheet'!B22:H22` | None (Unavailable on allowed sources) |
| 21 | Sub-Schedule: Trade Payables (FY19–FY25 Historical Series) | 29972 | `unavailable_on_allowed_sources` | `'Balance Sheet'!B25:H25` | None (Unavailable on allowed sources) |
| 22 | Sub-Schedule: Non-Controlling Interest (FY19–FY25 Historical Series) | 13103 | `unavailable_on_allowed_sources` | `'Balance Sheet'!B26:H26` | None (Unavailable on allowed sources) |
| 23 | Cash from Operating Activity (FY19–FY25 Historical Series) | 3176 | `primary_sourced` | `'Cash Flow'!B4:H4` | https://www.screener.in/company/M&M/consolidated/ |
| 24 | Cash from Investing Activity (FY19–FY25 Historical Series) | -18616 | `primary_sourced` | `'Cash Flow'!B5:H5` | https://www.screener.in/company/M&M/consolidated/ |
| 25 | Cash from Financing Activity (FY19–FY25 Historical Series) | 15834 | `primary_sourced` | `'Cash Flow'!B6:H6` | https://www.screener.in/company/M&M/consolidated/ |
| 26 | Sub-Schedule: Fixed Assets Purchased / CapEx (FY19–FY25 Historical Series) | -11889 | `primary_sourced` | `'Cash Flow'!B10:H10` | https://www.screener.in/company/M&M/consolidated/ |
| 27 | Standalone Revenue from Operations (FY19–FY25 Historical Series) | 116484 | `unavailable_on_allowed_sources` | `'Segment Operations'!B19:H19` | None (Unavailable on allowed sources) |
| 28 | Standalone Core Operating Profit / EBITDA (FY19–FY25 Historical Series) | 18416 | `unavailable_on_allowed_sources` | `'Segment Operations'!B20:H20` | None (Unavailable on allowed sources) |
| 29 | Debtor Days (FY19–FY25 Historical Series) | 19 | `primary_sourced` | `'Segment Operations'!B29:H29` | https://www.screener.in/company/M&M/consolidated/ |
| 30 | Inventory Days (FY19–FY25 Historical Series) | 78 | `primary_sourced` | `'Segment Operations'!B30:H30` | https://www.screener.in/company/M&M/consolidated/ |
| 31 | Days Payable (FY19–FY25 Historical Series) | 118 | `primary_sourced` | `'Segment Operations'!B31:H31` | https://www.screener.in/company/M&M/consolidated/ |
| 32 | Diluted Equity Share Count (FY25 Close) | 124.3 | `primary_sourced` | `'Cover & Notes'!B12` | https://www.screener.in/company/M&M/consolidated/ |
| 33 | Historical Market Snapshot Price (15-Oct-2025 NSE Official Unadjusted Close) | 3497.2 | `historical_market_snapshot` | `'Cover & Notes'!B10` | https://www.nseindia.com |
| 34 | 2-Year Weekly OLS Regression Beta vs Nifty 50 (^NSEI) (Raw & Blume-Adjusted) | 1.3164 | `empirical_regression_derived` | `'Beta'!K13 (Adjusted) / 'Beta'!K7 (Raw)` | https://finance.yahoo.com |
| 35 | 104-Week Empirical Weekly Market Returns & OLS Regression Dataset | 104 | `empirical_market_dataset` | `'Beta'!B8:G112` | https://finance.yahoo.com |
| 36 | Primary Valuation Output (Sum-of-the-Parts (SOTP) Valuation) | None | `unavailable` | `'DCF 1st way'!E42` | None (Unavailable on allowed sources) |
