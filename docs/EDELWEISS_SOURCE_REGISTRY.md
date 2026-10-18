# Source Registry & Provenance Ledger: EDELWEISS (Template V8)

## 1. Executive Summary
- **Entity**: `EDELWEISS`
- **Active Deliverable**: `companies/EDELWEISS/template_v8/EDELWEISS-FullTemplate.xlsx`
- **Workbook SHA-256**: `196d46b5e3d0283b0cba0f4f787cc9516d43b7093da3b16cef6c2b930a49c6e5`
- **Master Template**: `originals/Marksans Keshav.xlsx` (SHA-256 `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`)
- **Valuation Status**: `UNAVAILABLE`
- **Fair Value per Share**: `Unavailable` (linked from `'DCF 1st way'!E42` to `'Model ->'!C9`)
- **Evaluated Cost of Capital**: WACC = `10.47%` (Raw Beta = `2.1231`, Blume-Adjusted Beta = `1.7487`)
- **Beta Source**: Yahoo Finance 104-week weekly price regression ending 15-Oct-2025 (`work/sources/weekly_market_returns_104w_v2.json`, SHA-256 `19bd51f0ce6bcadfd4a7bfa7e0432532251bbcda6812486411300d6df9f16fb0`)
- **Risk-Free Rate**: 6.51% (Trading Economics dated prior-session proxy as of 14-Oct-2025 preceding 15-Oct-2025 cutoff)
- **Balance Sheet Integrity**: Historical FY15–FY25 and forecast FY26E reconcile to exact `0.00 Cr` (Row 40 = Row 38 - Row 16 = 0.00 Cr) with zero plugs.

## 2. Valuation Provenance & Acceptance Audit
Valuation is marked Unavailable: Industrial FCFF is methodologically prohibited for financial institutions. Vertical financial metrics (Alternatives PAT, EARCL BV, Mutual Fund AUM, Insurance IEV) and Standalone HoldCo Net Debt are statutory disclosures not published on Screener.in. Under the locked allowlist rule, they cannot be used in calculation or relabeled as analyst assumptions.

> [!WARNING]
> **METHODOLOGICAL EXCLUSION OF INDUSTRIAL FCFF**:
> Industrial Free Cash Flow to Firm (FCFF) is strictly withheld and blocked for Edelweiss Financial Services Limited.
> Financial institutions rely on debt, deposits, and borrowings as operating raw material rather than corporate capital structure leverage;
> operating cash flow (`+₹3,384 Cr` in FY25) reflects wholesale loan book repayments funding debt reduction (`-₹4,828 Cr` CFF). Sum-of-the-Parts (SOTP) is the only defensible method.

## 3. Audited Source Entries
| Index | Metric | Stated / Evaluated Value | Status | Sourced Cell | Official Allowlisted URL |
|:-----:|:-------|:------------------------:|:------:|:------------:|:-------------------------|
| 1 | Total Consolidated Operating Income (FY19–FY25 Historical Series) | 9393 | `primary_sourced` | `'Income Statement'!B5:H5` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 2 | Employee, Insurance & Operating Expenses (FY19–FY25 Historical Series) | 6033 | `primary_sourced` | `'Income Statement'!B6:H6` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 3 | Operating Profit Before Finance Costs (FY19–FY25 Historical Series) | 3360 | `primary_sourced` | `'Income Statement'!B7:H7` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 4 | Depreciation & Amortization (FY19–FY25 Historical Series) | 147 | `primary_sourced` | `'Income Statement'!B9:H9` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 5 | Finance Costs / Borrowing Interest (FY19–FY25 Historical Series) | 2537 | `primary_sourced` | `'Income Statement'!B10:H10` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 6 | Other Income (FY19–FY25 Historical Series) | 126 | `primary_sourced` | `'Income Statement'!B11:H11` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 7 | Tax Expense (FY19–FY25 Historical Series) | 266 | `primary_sourced` | `'Income Statement'!B13:H13` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 8 | Reported Consolidated PAT (FY19–FY25 Historical Series) | 536 | `primary_sourced` | `'Income Statement'!B14:H14` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 9 | Equity Share Capital (FY19–FY25 Historical Series) | 92 | `primary_sourced` | `'Balance Sheet'!B5:H5` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 10 | Reserves & Surplus (FY19–FY25 Historical Series) | 4333 | `primary_sourced` | `'Balance Sheet'!B6:H6` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 11 | Total Borrowings (FY19–FY25 Historical Series) | 18004 | `primary_sourced` | `'Balance Sheet'!B7:H7` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 12 | Other Liabilities & Policyholder Funds (FY19–FY25 Historical Series) | 18853 | `primary_sourced` | `'Balance Sheet'!B8:H8` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 13 | Net Block / Fixed Assets (FY19–FY25 Historical Series) | 1271 | `primary_sourced` | `'Balance Sheet'!B10:H10` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 14 | Capital Work in Progress (FY19–FY25 Historical Series) | 18 | `primary_sourced` | `'Balance Sheet'!B11:H11` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 15 | Financial Investments (FY19–FY25 Historical Series) | 17574 | `primary_sourced` | `'Balance Sheet'!B12:H12` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 16 | Loans, Credit Book & Cash (FY19–FY25 Historical Series) | 22419 | `primary_sourced` | `'Balance Sheet'!B13:H13` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 17 | Sub-Schedule: Inventories (FY19–FY25 Historical Series) | 0 | `primary_sourced` | `'Balance Sheet'!B19:H19` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 18 | Sub-Schedule: Trade Receivables (FY19–FY25 Historical Series) | 394 | `primary_sourced` | `'Balance Sheet'!B20:H20` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 19 | Sub-Schedule: Cash & Equivalents (FY19–FY25 Historical Series) | 5241 | `primary_sourced` | `'Balance Sheet'!B21:H21` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 20 | Sub-Schedule: Loans & Advances (FY19–FY25 Historical Series) | 1237 | `primary_sourced` | `'Balance Sheet'!B22:H22` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 21 | Sub-Schedule: Credit Book & Financial Assets (FY19–FY25 Historical Series) | 15547 | `primary_sourced` | `'Balance Sheet'!B23:H23` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 22 | Sub-Schedule: Trade Payables (FY19–FY25 Historical Series) | 1260 | `primary_sourced` | `'Balance Sheet'!B25:H25` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 23 | Sub-Schedule: Non-Controlling Interest (FY19–FY25 Historical Series) | 1493 | `primary_sourced` | `'Balance Sheet'!B26:H26` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 24 | Sub-Schedule: Policyholder & Other Liabilities (FY19–FY25 Historical Series) | 16100 | `primary_sourced` | `'Balance Sheet'!B27:H27` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 25 | Cash from Operating Activity (FY19–FY25 Historical Series) | 3384 | `primary_sourced` | `'Cash Flow'!B4:H4` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 26 | Cash from Investing Activity (FY19–FY25 Historical Series) | 3794 | `primary_sourced` | `'Cash Flow'!B5:H5` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 27 | Cash from Financing Activity (FY19–FY25 Historical Series) | -4828 | `primary_sourced` | `'Cash Flow'!B6:H6` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 28 | Sub-Schedule: Fixed Assets Purchased / CapEx (FY19–FY25 Historical Series) | -104 | `primary_sourced` | `'Cash Flow'!B9:H9` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 29 | Standalone EFSL HoldCo Borrowings (FY19–FY25 Historical Series) | 3072 | `unavailable_on_allowed_sources` | `'SOTP Valuation'!B36` | None (Unavailable on allowed sources) |
| 30 | Debtor / Fee Collection Days (FY19–FY25 Historical Series) | 15 | `primary_sourced` | `'Segment Operations'!B36:H36` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 31 | Diluted Equity Share Count (FY25 Close) | 94.5 | `primary_sourced` | `'Cover & Notes'!B12` | https://www.screener.in/company/EDELWEISS/consolidated/ |
| 32 | Historical Market Snapshot Price (15-Oct-2025 NSE Official Unadjusted Close) | 108.85 | `historical_market_snapshot` | `'Cover & Notes'!B10` | https://www.nseindia.com |
| 33 | 2-Year Weekly OLS Regression Beta vs Nifty 50 (^NSEI) (Raw & Blume-Adjusted) | 1.7487 | `empirical_regression_derived` | `'Beta'!K13 (Adjusted) / 'Beta'!K7 (Raw)` | https://finance.yahoo.com |
| 34 | 104-Week Empirical Weekly Market Returns & OLS Regression Dataset | 104 | `empirical_market_dataset` | `'Beta'!B8:G112` | https://finance.yahoo.com |
| 35 | Primary Valuation Output (Financial Services Sum-of-the-Parts (SOTP) Valuation) | None | `unavailable` | `'DCF 1st way'!E42` | None (Unavailable on allowed sources) |
