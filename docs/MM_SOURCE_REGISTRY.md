# Active Source Registry — Mahindra & Mahindra Template V17

Scope: active financial inputs use Screener and Yahoo Finance from the original workbook allowlist. The cover business-profile label uses Screener only and is descriptive, not a financial input. The valuation/data cutoff is 2025-10-15; P&L actuals are FY17-FY25, balance sheet and cash flow actuals are FY15-FY25.

| Metric / output | Period | FY25 / evaluated value | Workbook location | Status / source |
|---|---|---:|---|---|
| Business profile summary | Retrieved 2026-10-19; descriptive only | Automotive, farm equipment and financial-services presence | 'Model ->'!C5 | profile_sourced; Screener (About / Business Positioning) |
| Revenue from Operations (FY17-FY25) | FY17-FY25 | 159211 | 'Data Sheet'!B17:J17 | primary_sourced; Screener |
| Operating expenses (FY17-FY25) | FY17-FY25 | 128693 | 'Data Sheet'!B18:J18 | primary_sourced; Screener |
| Other income (FY17-FY25) | FY17-FY25 | 3718 | 'Data Sheet'!B25:J25 | primary_sourced; Screener |
| Depreciation and amortization (FY17-FY25) | FY17-FY25 | 6074 | 'Data Sheet'!B26:J26 | primary_sourced; Screener |
| Finance costs / interest (FY17-FY25) | FY17-FY25 | 9083 | 'Data Sheet'!B27:J27 | primary_sourced; Screener |
| Profit before tax (FY17-FY25) | FY17-FY25 | 19080 | 'Data Sheet'!B28:J28 | primary_sourced; Screener |
| Net profit (FY17-FY25) | FY17-FY25 | 14073 | 'Data Sheet'!B30:J30 | primary_sourced; Screener |
| Dividend payout ratio (FY17-FY25) | FY17-FY25 | 0.22 | 'Data Sheet'!B31:J31 | primary_sourced; Screener |
| Earnings per share (FY17-FY25) | FY17-FY25 | 103.97 | 'Data Sheet'!B32:J32 | primary_sourced; Screener |
| Operating profit / EBITDA proxy (sales less expenses) | FY17-FY25 | 30518 | 'Data Sheet'!B33:J33 | formula_derived; Screener inputs |
| Tax expense (PBT less net profit) | FY17-FY25 | 5007 | 'Data Sheet'!B29:J29 | formula_derived; Screener inputs |
| Equity share capital (FY15-FY25) | FY15-FY25 | 558 | 'BS'!C4:M4 | primary_sourced; Screener |
| Reserves (FY15-FY25) | FY15-FY25 | 76481 | 'BS'!C5:M5 | primary_sourced; Screener |
| Borrowings (FY15-FY25) | FY15-FY25 | 129025 | 'BS'!C10:M10 | primary_sourced; Screener |
| Other liabilities (FY15-FY25) | FY15-FY25 | 69950 | 'BS'!C15:M15 | primary_sourced; Screener |
| Total liabilities (FY15-FY25) | FY15-FY25 | 276013 | 'BS'!C16:M16 | primary_sourced; Screener |
| Net block / fixed assets (FY15-FY25) | FY15-FY25 | 34636 | 'BS'!C28:M28 | primary_sourced; Screener |
| Capital work in progress (FY15-FY25) | FY15-FY25 | 8298 | 'BS'!C30:M30 | primary_sourced; Screener |
| Investments (FY15-FY25) | FY15-FY25 | 41309 | 'BS'!C31:M31 | primary_sourced; Screener |
| Other assets aggregate (FY15-FY25) | FY15-FY25 | 191771 | 'BS'!C37:M37 | primary_sourced; Screener |
| Total assets (FY15-FY25) | FY15-FY25 | 276013 | 'BS'!C38:M38 | primary_sourced; Screener |
| Inventories detail | FY15-FY25 | Unavailable | 'BS'!C33:M33 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Trade receivables detail | FY15-FY25 | Unavailable | 'BS'!C34:M34 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Cash equivalents detail | FY15-FY25 | Unavailable | 'BS'!C35:M35 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Loans and advances / MMFS loan detail detail | FY15-FY25 | Unavailable | 'BS'!C36:M36 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Trade payables detail | FY15-FY25 | Unavailable | 'BS'!C13:M13 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Cash from operating activities (FY15-FY25) | FY15-FY25 | 3176 | 'Cash Flow'!C5:M5 | primary_sourced; Screener |
| Cash from investing activities (FY15-FY25) | FY15-FY25 | -18616 | 'Cash Flow'!C14:M14 | primary_sourced; Screener |
| Cash from financing activities (FY15-FY25) | FY15-FY25 | 15834 | 'Cash Flow'!C23:M23 | primary_sourced; Screener |
| Free cash flow (FY15-FY25) | FY15-FY25 | -6633 | 'Cash Flow'!C32:M32 | primary_sourced; Screener |
| CFO / operating profit (FY15-FY25) | FY15-FY25 | 0.26 | 'Cash Flow'!C33:M33 | primary_sourced; Screener |
| Net cash flow (CFO + CFI + CFF) | FY15-FY25 | 394 | 'Cash Flow'!C31:M31 | formula_derived; Screener inputs |
| CapEx / fixed assets purchased | FY15-FY25 | Unavailable | No input cell | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Unadjusted close (2025-10-15) | 2025-10-15 | 3497.2 | 'Data Sheet'!B8; 'Comps'!D8 | historical_market_snapshot; Yahoo Finance |
| Shares outstanding at 2025-10-15 cutoff | 2025-10-15 | Unavailable | 'Data Sheet'!B6; 'Comps'!E8 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Market capitalization / equity value | 2025-10-15 | Unavailable | 'Data Sheet'!B9; 'Comps'!F8 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Face value | Screener source snapshot (timestamp not recorded) | 5 | 'Data Sheet'!B7 | primary_sourced; Screener |
| Historical P/E (close / FY25 EPS) | 2025-10-15 price / FY25 reported EPS | 33.6366 | 'Comps'!P8 | formula_derived; Yahoo Finance + Screener |
| Raw OLS beta vs NIFTY 50 | 104 weekly returns ending 2025-10-12 | 1.4746 | 'Beta'!K7 | formula_derived; Yahoo Finance |
| Blume adjusted beta | 2025-10-15 historical cutoff | 1.3164 | 'Beta'!K13 | formula_derived; Yahoo Finance |
| Aligned weekly price history | 2023-10-15 to 2025-10-12 | 104 | 'Beta'!B8:G112 | empirical_market_dataset; Yahoo Finance |
| Historical balance checks | FY15-FY25 | 0 | 'BS'!C40:M40 | formula_derived; Screener inputs |
| FY26E forecast balance | FY26E | BLOCKED | 'BS'!N40 | blocked; Screener |
| WACC | 2025-10-15 cutoff | Unavailable | 'WACC-DCF'!J39 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |
| Primary fair value / SOTP | 2025-10-15 cutoff | Unavailable | 'DCF 1st way'!E42; 'Model ->'!C9 | unavailable_on_allowed_sources; Screener — unavailable in approved snapshot |

## Important corrections and limitations

- CapEx detail is unavailable: the Screener cash-flow snapshot reports CFO, CFI, CFF, net cash flow, free cash flow, and CFO/operating-profit ratio, but no consolidated fixed-assets-purchased sub-schedule. The former -11,889 Cr entry was not a CapEx fact and is withdrawn.
- A cutoff-dated diluted share count could not be reproduced from the approved-source snapshots. The old 124.3 crore hardcode is removed; market capitalization/equity value are unavailable. The Comps sheet P/E is a formula using the Yahoo Finance 15-Oct-2025 close and Screener FY25 EPS, not an intrinsic valuation or peer multiple.
- Screener reports FY25 Other Assets as an aggregate; it does not classify the included MMFS loan book into an operating current/non-current split. ROIC and intrinsic growth are therefore unavailable; detail cells remain blank rather than zero.
- Historical asset-component totals can differ from Screener reported total assets by INR 1 crore because components are rounded. The difference is shown in Main Financials; reported total assets and liabilities remain unchanged and no plug is applied.
- No BSE, company investor-relations files, annual-report PDFs, RBI, Damodaran, or other provider was used in the active V17 workbook. Legacy V16 source-coverage files are superseded and must not be used as V17 evidence.
- WACC and fair value remain unavailable; FY26E is blocked; no numeric fair value is approved. See docs/LOCKED_SOURCE_ALLOWLIST.md and docs/USER_APPROVAL_PACKET.md for release conditions.
