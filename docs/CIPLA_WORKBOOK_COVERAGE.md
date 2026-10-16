# Cipla Workbook Adaptation Coverage Report

## 1. Executive Summary
This report documents the cell-by-cell adaptation of the template financial model (`originals/Marksans Keshav.xlsx`, SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`) to **Cipla Limited** (`CIPLA`, BSE `500087`).

- **Target Candidate Workbook**: `candidates/Cipla-Recalculated.xlsx`
- **SHA-256**: `050e13b2be37705ca88af6a63022de00f9b2d07796565071619ad1daa6be8bff`
- **Total Cell Entries Adapted**: 359
- **Original Sheets Preserved**: 17 of 17 (100%)
- **Formula Invariance**: Zero underlying Excel formulas deleted or corrupted; all downstream recalculation flows via native Excel calculation engine.

---

## 2. Coverage by Classification

| Classification | Cell Count | Percentage | Description |
|:---|:---:|:---:|:---|
| `replaced_with_cipla_data` | 351 | 97.8% | Sourced financial facts from Cipla FY18–FY25 annual reports, BSE filings, and Screener snapshot |
| `recalculated` | 1 | 0.3% | Formulaically updated dynamic cells (e.g. implied share count = Market Cap / Current Price) |
| `generic_assumption` | 7 | 1.9% | Transparent macroeconomic / valuation assumptions (e.g. Rf=7.30%, Terminal Growth=5.00%, Kd=7.50%) |
| **Total** | **359** | **100.0%** | **Complete bounded adaptation** |

---

## 3. Coverage by Worksheet

| Worksheet | Adapted Cells | Primary Data Replaced |
|:---|:---:|:---|
| `Data Sheet` | 311 | P&L (FY18–FY25, 9 years incl TTM), Balance Sheet (Assets & Liabilities balanced Delta=0.00), Face Value, Price, Shares |
| `Cash Flow` | 32 | Historical Operating, Investing, and Financing Cash Flows across 8 years (FY18–FY25) tying to Net Cash Flow |
| `WACC-DCF` | 5 | Pre-tax borrowing rate (7.50%), Current borrowings (₹613.77 Cr), Market cap (₹105,108 Cr) |
| `Comps` | 3 | Ticker label, formula link to peer compdata row 10, table header |
| `DCF 1st way` | 4 | Market price (₹1,301.00), expected growth (10.0%), terminal growth (5.0%), reinvestment rate (45.0%) |
| `DCF 2nd way` | 2 | Company ticker header (`CIPLA`), terminal growth rate (5.0%) |
| `Ratio Analysis` | 1 | Sheet header text label (`Ratio Analysis of - Cipla Ltd.`) |
| `Beta` | 1 | Weekly returns title label (`Cipla past 2 year weekly returns`) |

---

## 4. Reconciled Financial Identities

1. **Balance Sheet Equilibrium (Data Sheet)**:
   - For every historical year FY18 through FY25: Total Assets - Total Liabilities = 0.00.
   - FY25 Total Assets (`Data Sheet!J66`) = ₹37,386.58 Cr.
   - FY25 Total Liabilities (`Data Sheet!J61`) = ₹37,386.58 Cr.
   - Balance Sheet Discrepancy = ₹0.00 Cr.

2. **Cash Flow Statement Equilibrium (Cash Flow & Data Sheet)**:
   - FY25 Cash from Operations (CFO): ₹5,005.00 Cr.
   - FY25 Cash from Investing (CFI): -₹3,682.00 Cr.
   - FY25 Cash from Financing (CFF): -₹1,293.00 Cr.
   - Net Cash Flow: 5,005 - 3,682 - 1,293 = +30.00 Cr (`Data Sheet!J85`, Delta = 0.00).

3. **Valuation & WACC Identity Checks (WACC-DCF)**:
   - Risk-Free Rate (Rf): 7.30%.
   - Market Return (Rm): 16.37% (Price return 15.15% + Dividend yield 1.22%).
   - Levered Beta (Hamada): 0.547.
   - Cost of Equity (Ke): 7.30% + 0.547 * (16.37% - 7.30%) = 12.26% (Identity holds, tolerance 0.0001).
   - Pre-tax Kd: 7.50%, Tax Rate: 30.00%, Post-tax Kd: 7.50% * (1 - 0.30) = 5.25%.
   - Target Capital Structure: Peer Median Equity Weight 97.76%, Debt Weight 2.24% (97.76% + 2.24% = 100.00%).
   - Weighted Average Cost of Capital (WACC): 97.76% * 12.26% + 2.24% * 5.25% = 12.10% (Identity holds).

---

## 5. Provenance Ledger & Audit Trail
The complete cell-by-cell log with original values, new values, classifications, and rationales is stored in machine-readable JSON format at [`docs/cipla_workbook_coverage.json`](file:///C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker/docs/cipla_workbook_coverage.json).
