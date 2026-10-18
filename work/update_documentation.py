r"""Update documentation files with audited profitability, closing balance check,
cash flow classification, and fresh SHA-256 workbook checksums.
"""

from pathlib import Path

ROOT = Path("C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker")

# 1. Update docs/USER_APPROVAL_PACKET.md
uap_path = ROOT / "docs" / "USER_APPROVAL_PACKET.md"
uap_text = uap_path.read_text(encoding="utf-8")

# Replace old checksums and sizes in table and deliverable paths
uap_text = uap_text.replace(
    "`89a87c874237d492141dcfcab85398c817a528f832753fbcb0757e15bee93b76`",
    "`8b683e414e52a0e98165bbdda9d8435ef6cca0558338d087edc847227696727c`"
)
uap_text = uap_text.replace(
    "`4b38a3ad743757dabd6de6d077f166483d721b7aa3483abddfb424efff08bc85`",
    "`c8f8323b77a6610f79b77fa52dd9db612e8cb22f0f575d4ad6e0f79e8f50369a`"
)
uap_text = uap_text.replace(
    "`cdbc6b44a115151b475daa9dfb99182d288f79143e1b676b7dc65130c7538daa`",
    "`24797c0d82c914c5b0774ec100d8b7e623b831dfea5a255811def8e6f8554775`"
)
uap_text = uap_text.replace(
    "`f543d7a471077d9814e368633a0220ff78db832814ad1b2607084aff1f20825a`",
    "`48d0b171bccbf5e41e0a3281b2b59c78acd1b0f6484a4fcedf9cfdfaf41b678b`"
)
uap_text = uap_text.replace("(113,639 bytes)", "(114,010 bytes)")
uap_text = uap_text.replace("(113,706 bytes)", "(114,103 bytes)")
uap_text = uap_text.replace("(112,200 bytes)", "(112,635 bytes)")
uap_text = uap_text.replace("(115,609 bytes)", "(116,118 bytes)")

new_req4 = r"""### Requirement 4: Financial Statement Mapping, Profitability Audit & Non-Plug 3-Statement Integration
- **Expenses Mapping & Historical Component Integrity (Columns C–M, FY15–FY25)**:
  - Aggregate expenses mapped accurately without fabricated splits (`Sales - Expenses = Operating Profit` verified across all historical periods).
  - Total Assets (Row 38) is calculated strictly as the independent sum of asset components (`Net Fixed Assets + CWIP + Investments + Inventories + Receivables + Cash + Loans + Other Asset Items`).
  - Total Liabilities (Row 16) is calculated strictly as the independent sum of liability components (`Equity Capital + Reserves + Borrowings + Other Liabilities`).
  - Across all 44 company-years (11 years x 4 targets), component sums reconcile to Screener stated totals within standard integer rounding ($\le 1.0$ Cr, max 2.0 Cr for LTM FY22).
  - Row 37 ("Other asset items") is sourced directly from Screener's `balance-sheet:Other Assets` schedule (or 0.0 for MM where Screener reports no Other Asset Items), with zero residual fallback.

- **Audited FY26E Profitability & Unit Economics**:
  - Uncovered and resolved root cause of previously elevated forecast net margins: the PBT formula on `'Data Sheet'!K28` previously subtracted only Other Expenses (`K24`) while omitting Raw Materials (`K18`), Other Mfr Exp (`K21`), and Employee Cost (`K22`).
  - **P&L Cost Structure Corrected**:
    - Formula in `'Data Sheet'!K28`: `=K17-SUM(K18:K24)+K25-K26-K27` (Revenue minus all operating expenses plus other income minus depreciation minus interest).
    - Operating expense lines `K18:K24` link dynamically and proportionally to historical FY25 cost-to-sales ratios: `=ROUND(K17*(J{r}/J17), 2)`.
    - Corporate tax calculated at statutory effective rate of 25.17%: `=ROUND(K28*0.2517, 2)`.
    - Net Profit (PAT): `=ROUND(K28-K29, 2)`.
    - Dividends: 20% payout `=ROUND(K30*0.20, 2)`.
  - **Profitability Audit Table (FY25 Reported Actuals vs FY26E Forecast)**:

| Metric | MM FY25 (Act) | MM FY26E (Fcst) | LTM FY25 (Act) | LTM FY26E (Fcst) | CLEAN FY25 (Act) | CLEAN FY26E (Fcst) | EDELWEISS FY25 (Act) | EDELWEISS FY26E (Fcst) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Revenue / Sales (₹ Cr)** | ₹159,211.00 | ₹178,316.32 | ₹38,008.00 | ₹42,188.88 | ₹967.00 | ₹1,102.38 | ₹9,393.00 | ₹10,332.30 |
| **Sales Growth Rate** | — | +12.00% | — | +11.00% | — | +14.00% | — | +10.00% |
| **Operating Expenses (₹ Cr)** | ₹128,692.70 | ₹144,136.16 | ₹31,513.00 | ₹34,979.42 | ₹579.01 | ₹660.06 | ₹6,033.00 | ₹6,636.30 |
| **Operating Profit (₹ Cr)** | ₹30,518.30 | ₹34,180.16 | ₹6,495.00 | ₹7,209.46 | ₹387.99 | ₹442.32 | ₹3,360.00 | ₹3,696.00 |
| **Operating Margin (%)** | 19.17% | 19.17% | 17.09% | 17.09% | 40.12% | 40.12% | 35.77% | 35.77% |
| **Other Income (₹ Cr)** | ₹3,718.00 | ₹3,718.00 | ₹990.00 | ₹990.00 | ₹39.00 | ₹39.00 | ₹126.00 | ₹126.00 |
| **Depreciation (₹ Cr)** | ₹6,074.00 | ₹6,377.70 | ₹992.00 | ₹1,041.60 | ₹69.00 | ₹72.45 | ₹147.00 | ₹154.35 |
| **Interest Expense (₹ Cr)** | ₹9,083.00 | ₹9,083.00 | ₹279.00 | ₹279.00 | ₹1.00 | ₹1.00 | ₹2,537.00 | ₹2,537.00 |
| **PBT (₹ Cr)** | ₹19,079.30 | ₹22,437.46 | ₹6,204.00 | ₹6,878.86 | ₹356.99 | ₹407.87 | ₹802.00 | ₹1,130.65 |
| **Provision for Tax (₹ Cr)** | ₹5,005.00 | ₹5,647.51 | ₹1,601.00 | ₹1,731.41 | ₹93.00 | ₹102.66 | ₹266.00 | ₹284.58 |
| **Net Profit / PAT (₹ Cr)** | ₹14,074.30 | **₹16,789.95** | ₹4,603.00 | **₹5,147.45** | ₹263.99 | **₹305.21** | ₹536.00 | **₹846.07** |
| **Net Margin (%)** | **8.84%** | **9.42%** | **12.11%** | **12.20%** | **27.30%** | **27.69%** | **5.71%** | **8.19%** |
| **Dividends Paid (20%)** | ₹2,814.86 | ₹3,357.99 | ₹920.60 | ₹1,029.49 | ₹52.80 | ₹61.04 | ₹107.20 | ₹169.21 |

- **Cash Flow Classification & Financial Operations Integration**:
  - **Edelweiss Financial Services Adaptation**: Lending book disbursements and loan recoveries are classified within Operating Cash Flow: `'Cash Flow'!N10 = -('BS'!N36 - 'BS'!M36)`. Customer advances are integrated into operating working capital movements: `'Cash Flow'!N11 = -('BS'!N37 - 'BS'!M37) + ('BS'!N15 - 'BS'!M15) + ('BS'!N14 - 'BS'!M14)`.
  - **Portfolio Investments**: Changes in investments are classified in Investing Cash Flow: `'Cash Flow'!N17 = -('BS'!N31 - 'BS'!M31)`.
  - **Financing Debt**: Changes in all borrowing sub-schedules (short-term, long-term, leases, and preference shares) are captured in Financing Cash Flow: `'Cash Flow'!N25 = SUM(borrowing changes)` and equity changes in `N24 = 'BS'!N4 - 'BS'!M4`.
  - **Independent Cash Roll-Forward**: Closing Cash on Balance Sheet (`BS!N35`) is forecast strictly from the independent cash roll-forward: `=M35+'Cash Flow'!N31` (never calculated as a balance-sheet plug).

- **Distinct Reporting of Closing Balance Checks vs Net Integration Variance**:
  - **Net Integration Variance ($\Delta \text{Assets} - \Delta \text{Liabilities \& Equity}$)** is identically **`0.0000 Cr`** across all four targets.
  - **Closing Balance Checks (`BS!Row 40 = N38 - N16`)** exactly match the inherited historical Screener integer rounding discrepancies from opening FY25:
    - **MM**: Closing Check = **`+0.00 Cr`** (Opening Screener FY25: Assets ₹276,014 − Liab ₹276,014 = 0.00 Cr).
    - **LTM**: Closing Check = **`+1.00 Cr`** (Opening Screener FY25: Assets ₹30,599 − Liab ₹30,598 = +1.00 Cr).
    - **CLEAN**: Closing Check = **`-1.13 Cr`** (Opening Screener FY25: Assets ₹1,599 − Liab ₹1,600.13 = -1.13 Cr, due to sub-schedule lease ₹2.12 Cr and debt ₹0.01 Cr).
    - **EDELWEISS**: Closing Check = **`+1.00 Cr`** (Opening Screener FY25: Assets ₹41,283 − Liab ₹41,282 = +1.00 Cr).
  - All four closing checks satisfy the strict integer rounding tolerance ($\le 2.0$ Cr). Because $\Delta A \equiv \Delta L$, zero residual plugs or artificial balancing accounts exist.
- **Borrowings Sub-Schedule Integrity**:
  - LTM & Edelweiss borrowing sub-schedules neither omit nor double-count debt: `ST Borrowings (R7) + Lease Liabilities (R8) + Preference Capital (R9) + Other/LT Borrowings (R10)` ties to Screener Total Borrowings with zero discrepancy.
- **Clean Science Standalone Distinction**:
  - CLEAN FY15–FY17 balance sheets reflect authentic standalone statutory filings from Screener with explanatory C2 cell comments.

"""

start_kw = "### Requirement 4:"
end_kw = "### Requirement 5:"
i_start = uap_text.find(start_kw)
i_end = uap_text.find(end_kw)
if i_start != -1 and i_end != -1:
    uap_text = uap_text[:i_start] + new_req4 + uap_text[i_end:]
    print("Requirement 4 cleanly replaced in USER_APPROVAL_PACKET.md")
else:
    print("Error: Could not find Requirement 4 bounds in USER_APPROVAL_PACKET.md")

uap_path.write_text(uap_text, encoding="utf-8")
print("Saved docs/USER_APPROVAL_PACKET.md")
