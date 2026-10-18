# Source Registry & Provenance Ledger: Mahindra & Mahindra (Template V17, Issue #8)

## 1. Executive Summary
- **Entity**: `MM` (Mahindra & Mahindra Limited)
- **Active Deliverable**: `companies/MM/template_v17/MM-FullTemplate.xlsx` (also packaged at `deliverables/V16_user_handoff/workbooks/MM-FullTemplate.xlsx`)
- **Workbook SHA-256**: `bdfd6a9903f1c6be34d13641dbdb7ecf0a04c83e04653996239d4d720704c54d` (92,153 bytes)
- **Master Template**: `originals/Marksans Keshav.xlsx` (SHA-256 `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`, 16 sheets)
- **Valuation Status**: `UNAVAILABLE`
- **Fair Value per Share**: `Unavailable` (linked dynamically from `'DCF 1st way'!E42` to `'Model ->'!C9`)
- **Evaluated Cost of Capital**: WACC = `Unavailable` (`'WACC-DCF'!J39`); Raw OLS Beta = `1.4746` (`'Beta'!K7`), Blume-Adjusted Beta = `1.3164` (`'Beta'!K13`)
- **Beta Source**: Yahoo Finance 104-week weekly price regression ending 15-Oct-2025 (`work/sources/weekly_market_returns_104w_v2.json`, SHA-256 `19bd51f0ce6bcadfd4a7bfa7e0432532251bbcda6812486411300d6df9f16fb0`)
- **Risk-Free Rate**: 6.51% (Trading Economics dated prior-session proxy as of 14-Oct-2025 preceding 15-Oct-2025 cutoff)
- **Balance Sheet Integrity**: Historical FY15–FY25 ties to exact `0.00 Cr` (`BS!C40:M40 = 0.00 Cr`). Forecast FY26E (`BS!N40`) is marked `BLOCKED` due to unevidenced consolidated NBFC funding bridge.

## 2. Valuation Provenance & Acceptance Audit
Valuation is marked `Unavailable`: Factual segment EBITDA inputs (Automotive, Farm Equipment) and Standalone Net Cash are statutory filing items not published on allowlisted sources (Screener.in). Under the locked allowlist rule, they cannot be used in calculation or relabeled as analyst assumptions. Industrial FCFF DCF is methodologically disallowed for a conglomerate where >50% of assets belong to a systemic NBFC (MMFS).

## 3. Audited Source Entries
| Index | Metric | Stated / Evaluated Value | Status | Sourced Cell | Official Allowlisted URL |
|:-----:|:-------|:------------------------:|:------:|:------------:|:-------------------------|
| 1 | Revenue from Operations (FY15–FY25 Historical Series) | 1,59,211 (FY25) | `primary_sourced` | `'Data Sheet'!B6:L6` | https://www.screener.in/company/M&M/consolidated/ |
| 2 | Total Operating Expenditure (FY15–FY25 Historical Series) | 1,28,693 (FY25) | `primary_sourced` | `'Data Sheet'!B7:L7` | https://www.screener.in/company/M&M/consolidated/ |
| 3 | Operating Profit / Consolidated EBITDA (FY15–FY25 Historical Series) | 30,518 (FY25) | `primary_sourced` | `'Data Sheet'!B8:L8` | https://www.screener.in/company/M&M/consolidated/ |
| 4 | Depreciation & Amortization (FY15–FY25 Historical Series) | 6,074 (FY25) | `primary_sourced` | `'Data Sheet'!B10:L10` | https://www.screener.in/company/M&M/consolidated/ |
| 5 | Finance Costs / Interest (FY15–FY25 Historical Series) | 9,083 (FY25) | `primary_sourced` | `'Data Sheet'!B11:L11` | https://www.screener.in/company/M&M/consolidated/ |
| 6 | Other Income (FY15–FY25 Historical Series) | 3,718 (FY25) | `primary_sourced` | `'Data Sheet'!B12:L12` | https://www.screener.in/company/M&M/consolidated/ |
| 7 | Tax Expense (FY15–FY25 Historical Series) | 5,006 (FY25) | `primary_sourced` | `'Data Sheet'!B14:L14` | https://www.screener.in/company/M&M/consolidated/ |
| 8 | Reported Consolidated PAT (FY15–FY25 Historical Series) | 14,073 (FY25) | `primary_sourced` | `'Data Sheet'!B15:L15` | https://www.screener.in/company/M&M/consolidated/ |
| 9 | Equity Share Capital (FY15–FY25 Historical Series) | 597 (FY25) | `primary_sourced` | `'BS'!C4:M4` | https://www.screener.in/company/M&M/consolidated/ |
| 10 | Reserves & Surplus (FY15–FY25 Historical Series) | 76,481 (FY25) | `primary_sourced` | `'BS'!C5:M5` | https://www.screener.in/company/M&M/consolidated/ |
| 11 | Total Borrowings (FY15–FY25 Historical Series) | 1,29,025 (FY25) | `primary_sourced` | `'BS'!C7:M7` | https://www.screener.in/company/M&M/consolidated/ |
| 12 | Other Liabilities (FY15–FY25 Historical Series) | 69,950 (FY25) | `primary_sourced` | `'BS'!C8:M8` | https://www.screener.in/company/M&M/consolidated/ |
| 13 | Net Block / Fixed Assets (FY15–FY25 Historical Series) | 34,636 (FY25) | `primary_sourced` | `'BS'!C20:M20` | https://www.screener.in/company/M&M/consolidated/ |
| 14 | Capital Work in Progress (FY15–FY25 Historical Series) | 8,298 (FY25) | `primary_sourced` | `'BS'!C21:M21` | https://www.screener.in/company/M&M/consolidated/ |
| 15 | Investments (FY15–FY25 Historical Series) | 41,309 (FY25) | `primary_sourced` | `'BS'!C22:M22` | https://www.screener.in/company/M&M/consolidated/ |
| 16 | Other Assets (FY15–FY25 Historical Series) | 1,91,771 (FY25) | `primary_sourced` | `'BS'!C23:M23` | https://www.screener.in/company/M&M/consolidated/ |
| 17 | Sub-Schedule: Inventories | Unavailable | `unavailable_on_allowed_sources` | Left blank | None (Unavailable on allowed sources) |
| 18 | Sub-Schedule: Trade Receivables | Unavailable | `unavailable_on_allowed_sources` | Left blank | None (Unavailable on allowed sources) |
| 19 | Sub-Schedule: Loans & Advances / MMFS Loan Book | Unavailable | `unavailable_on_allowed_sources` | Left blank | None (Unavailable on allowed sources) |
| 20 | Cash from Operating Activity (FY15–FY25 Historical Series) | 3,176 (FY25) | `primary_sourced` | `'Cash Flow'!C5:M5` | https://www.screener.in/company/M&M/consolidated/ |
| 21 | Cash from Investing Activity (FY15–FY25 Historical Series) | -18,616 (FY25) | `primary_sourced` | `'Cash Flow'!C14:M14` | https://www.screener.in/company/M&M/consolidated/ |
| 22 | Cash from Financing Activity (FY15–FY25 Historical Series) | 15,834 (FY25) | `primary_sourced` | `'Cash Flow'!C23:M23` | https://www.screener.in/company/M&M/consolidated/ |
| 23 | CapEx / Fixed Assets Purchased (FY15–FY25 Historical Series) | -11,889 (FY25) | `primary_sourced` | `'Cash Flow'!C15:M15` | https://www.screener.in/company/M&M/consolidated/ |
| 24 | Historical Market Snapshot Price (15-Oct-2025 Official NSE Unadjusted Close) | ₹3,497.20 | `historical_market_snapshot` | `'Comps'!C8` | https://www.nseindia.com |
| 25 | 2-Year Weekly OLS Regression Beta vs Nifty 50 (^NSEI) (Raw) | 1.4746 | `empirical_regression_derived` | `'Beta'!K7` | https://finance.yahoo.com |
| 26 | 2-Year Weekly OLS Regression Beta vs Nifty 50 (^NSEI) (Blume Adjusted) | 1.3164 | `empirical_regression_derived` | `'Beta'!K13` | https://finance.yahoo.com |
| 27 | 104-Week Empirical Weekly Market Returns Dataset (105 Aligned Prices) | 104 pairs | `empirical_market_dataset` | `'Beta'!A1:D105` | https://finance.yahoo.com |
| 28 | Benchmark 10-Year Sovereign Yield Proxy | 6.51% | `primary_sourced` | `'WACC-DCF'!J22` | https://tradingeconomics.com/india/government-bond-yield |
| 29 | Primary Valuation Output (SOTP Valuation) | Unavailable | `unavailable` | `'DCF 1st way'!E42` | None (SOTP segment facts unavailable on approved sources) |
| 30 | Cover Link Fair Value Output | Unavailable | `dynamic_formula` | `'Model ->'!C9` | `='DCF 1st way'!E42` |
