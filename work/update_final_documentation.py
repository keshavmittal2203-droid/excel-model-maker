"""Update documentation files with final risk-free rate verification (6.51% Trading Economics proxy),
analyst assumption classifications (ERP 5.50% and Tax Rate 25.17%), valuation trace,
before-and-after valuation matrix, and new workbook SHA-256 hashes.
"""

from pathlib import Path

ROOT = Path("C:/Users/kesha/Documents/Codex/2026-10-16/https-x-com-painn-x-status/outputs/excel-model-maker")

# 1. Update docs/USER_APPROVAL_PACKET.md
uap_path = ROOT / "docs" / "USER_APPROVAL_PACKET.md"
uap_text = uap_path.read_text(encoding="utf-8")

# Replace hashes and sizes
uap_text = uap_text.replace(
    "`8b683e414e52a0e98165bbdda9d8435ef6cca0558338d087edc847227696727c`",
    "`a0763519d1be2094efbfd47111c03ea4e26d2bb3a935d38e02eb37286870cb15`"
)
uap_text = uap_text.replace(
    "`c8f8323b77a6610f79b77fa52dd9db612e8cb22f0f575d4ad6e0f79e8f50369a`",
    "`28f4dc2fa0b603d777af10cc6ad54960407cd6ee4071f7a04b0ea5894f56a6e7`"
)
uap_text = uap_text.replace(
    "`24797c0d82c914c5b0774ec100d8b7e623b831dfea5a255811def8e6f8554775`",
    "`6c0dcb8f392e4db4ba179e59e4f11d199668bdcc3960a16ddad4179083840931`"
)
uap_text = uap_text.replace(
    "`48d0b171bccbf5e41e0a3281b2b59c78acd1b0f6484a4fcedf9cfdfaf41b678b`",
    "`744399a6979abcc972c69597370e341b5320f2a8cfbb9dd292f25de2515a0f7c`"
)

# Replace old sizes
uap_text = uap_text.replace("(114,010 bytes)", "(114,551 bytes)")
uap_text = uap_text.replace("(114,103 bytes)", "(114,540 bytes)")
uap_text = uap_text.replace("(112,635 bytes)", "(112,992 bytes)")
uap_text = uap_text.replace("(116,118 bytes)", "(116,493 bytes)")

# Replace WACC and valuation values in matrix
uap_text = uap_text.replace(
    "| **LTIMindtree Limited** (`LTM`) | Full-Year Annual FCFF DCF Baseline | `E42` (`=E40/E41`) | `C9` (`=E42`) | **₹2,428.22** | ₹5,609.50 | **13.12%** | **0 errors** | `28f4dc2fa0b603d777af10cc6ad54960407cd6ee4071f7a04b0ea5894f56a6e7` |",
    "| **LTIMindtree Limited** (`LTM`) | Full-Year Annual FCFF DCF Baseline | `E42` (`=E40/E41`) | `C9` (`=E42`) | **₹2,497.80** | ₹5,609.50 | **12.86%** | **0 errors** | `28f4dc2fa0b603d777af10cc6ad54960407cd6ee4071f7a04b0ea5894f56a6e7` |"
)
uap_text = uap_text.replace(
    "| **Clean Science and Technology** (`CLEAN`) | Manufacturing FCFF DCF Baseline | `E42` (`=E40/E41`) | `C9` (`=E42`) | **₹468.61** | ₹1,041.40 | **11.91%** | **0 errors** | `6c0dcb8f392e4db4ba179e59e4f11d199668bdcc3960a16ddad4179083840931` |",
    "| **Clean Science and Technology** (`CLEAN`) | Manufacturing FCFF DCF Baseline | `E42` (`=E40/E41`) | `C9` (`=E42`) | **₹487.28** | ₹1,041.40 | **11.65%** | **0 errors** | `6c0dcb8f392e4db4ba179e59e4f11d199668bdcc3960a16ddad4179083840931` |"
)
uap_text = uap_text.replace(
    "| **Mahindra & Mahindra Limited** (`MM`) | Sum-of-the-Parts (SOTP) Valuation | `E42` (`=G60`) | `C9` (`=E42`) | **₹2,742.97** | ₹3,497.20 | 13.00% |",
    "| **Mahindra & Mahindra Limited** (`MM`) | Sum-of-the-Parts (SOTP) Valuation | `E42` (`=G60`) | `C9` (`=E42`) | **₹2,742.97** | ₹3,497.20 | **12.79%** |"
)
uap_text = uap_text.replace(
    "| **Edelweiss Financial Services** (`EDELWEISS`) | Financial Services SOTP (Industrial FCFF N/A) | `E42` (`=G60`) | `C9` (`=E42`) | **₹130.53** | ₹108.85 | 9.70% |",
    "| **Edelweiss Financial Services** (`EDELWEISS`) | Financial Services SOTP (Industrial FCFF N/A) | `E42` (`=G60`) | `C9` (`=E42`) | **₹130.53** | ₹108.85 | **9.60%** |"
)

# Add Before-and-After Valuation Table after deliverable paths
before_after_section = """
### Before-and-After Valuation Comparison (Impact of Evidenced 6.51% Risk-Free Rate Proxy)

| Target Company | Valuation Model | Rf (Before -> After) | WACC (Before -> After) | Fair Value / Sh (Before) | Fair Value / Sh (After) | Delta (₹/sh) | Rationale for Value Movement |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **MM** | SOTP | 6.78% -> **6.51%** | 13.00% -> **12.79%** | ₹2,742.97 | **₹2,742.97** | ₹0.00 | SOTP reflects FY25 segment EBITDA & listed stakes; independent of WACC |
| **LTM** | FCFF DCF | 6.78% -> **6.51%** | 13.12% -> **12.86%** | ₹2,428.22 | **₹2,497.80** | +₹69.57 | Lower discount rate on explicit FY26E–FY30E cash flows & Gordon TV |
| **CLEAN** | FCFF DCF | 6.78% -> **6.51%** | 11.91% -> **11.65%** | ₹468.61 | **₹487.28** | +₹18.67 | Lower discount rate on explicit FY26E–FY30E cash flows & Gordon TV |
| **EDELWEISS** | SOTP | 6.78% -> **6.51%** | 9.70% -> **9.60%** | ₹130.53 | **₹130.53** | ₹0.00 | SOTP reflects segment PAT, Book Value, and AUM; independent of WACC |
"""

if "### Before-and-After Valuation Comparison" not in uap_text:
    idx_marker = uap_text.find("## 3. Verified Excel Defect Remediation Summary")
    if idx_marker != -1:
        uap_text = uap_text[:idx_marker] + before_after_section + "\n---\n\n" + uap_text[idx_marker:]
        print("Before-and-after table added to USER_APPROVAL_PACKET.md")

# Replace Requirement 6 text with updated WACC and Rf details
old_req6_snippet = """### Requirement 6: Beta & WACC Blume Adjustment Linkage
- 104 weekly observations aligned with NIFTY 50 closing prices from Yahoo Finance / NSE closing prices embedded in `Beta` sheet (Rows 8 to 112).
- Raw beta calculated transparently via `=SLOPE(D9:D112, G9:G112)`.
- Blume adjustment implemented transparently as `=(K7*K8)+(K11*K10)` with `K8 = 2/3` raw beta weight and `K11 = 1/3` market beta weight (`K10 = 1.0`).
- WACC beta links directly to Beta sheet via `WACC-DCF!J32 = =Beta!K13` and `WACC-DCF!J24 = =J32`.
- Target company capital structure weights `E30 = =D30` and `E31 = =E32-E30` dynamically reflect actual company debt and equity.
- Evaluated WACC:
  - `MM`: 13.00% (Raw Beta 1.7638 -> Blume Beta 1.5092, Cost of Equity 15.08%)
  - `LTM`: 13.12% (Raw Beta 1.2569 -> Blume Beta 1.1712, Cost of Equity 13.22%, Cost of Debt after-tax 5.61%, actual capital structure 98.7% equity / 1.3% debt)
  - `CLEAN`: 11.91% (Raw Beta 0.9006 -> Blume Beta 0.9337, Cost of Equity 11.92%, Cost of Debt after-tax 5.99%, actual capital structure 99.98% equity / 0.02% debt)
  - `EDELWEISS`: 9.70% (Raw Beta 1.4830 -> Blume Beta 1.3220, Cost of Equity 14.05%, Cost of Debt after-tax 6.45%)"""

new_req6_snippet = """### Requirement 6: Beta, WACC Blume Adjustment Linkage & Evidenced Risk-Free Rate
- 104 weekly observations aligned with NIFTY 50 closing prices from Yahoo Finance / NSE closing prices embedded in `Beta` sheet (Rows 8 to 112).
- Raw beta calculated transparently via `=SLOPE(D9:D112, G9:G112)`.
- Blume adjustment implemented transparently as `=(K7*K8)+(K11*K10)` with `K8 = 2/3` raw beta weight and `K11 = 1/3` market beta weight (`K10 = 1.0`).
- WACC beta links directly to Beta sheet via `WACC-DCF!J32 = =Beta!K13` and `WACC-DCF!J24 = =J32`.
- Target company capital structure weights `E30 = =D30` and `E31 = =E32-E30` dynamically reflect actual company debt and equity.
- **Evidenced Risk-Free Rate ($R_f$)**:
  - Cell `'WACC-DCF'!J22`: Updated to **6.51%** (0.0651), replacing the unevidenced 6.78% figure.
  - Sourced directly from allowlisted provider **Trading Economics** (`https://tradingeconomics.com/india/government-bond-yield`), supported by its 14-Oct-2025 published reading ("India 10-Year Bond Yield Eases to Near 1-Week Low") reporting the 10-year benchmark yield just below 6.51% (6.51%) as the exact dated prior-session proxy preceding the 15-Oct-2025 cutoff.
- **Explicit Classification of Assumptions**:
  - **Equity Risk Premium (ERP, `'WACC-DCF'!J23` = 5.50%)**: Classified strictly as an **Analyst Macroeconomic Assumption**; not attributed to Trading Economics.
  - **Marginal Corporate Tax Rate (`'WACC-DCF'!E23` and `'Data Sheet'!K29` = 25.17%)**: Classified strictly as an **Analyst Modeling Assumption** (22% base + 10% surcharge + 4% cess); not called a verified statutory filing quote since allowlisted sites do not publish statutory rate tables.
- **Evaluated WACC (Updated with 6.51% Rf)**:
  - `MM`: **12.79%** (Raw Beta 1.7638 -> Blume Beta 1.5092, Cost of Equity 14.81%, Cost of Debt after-tax 5.99%)
  - `LTM`: **12.86%** (Raw Beta 1.2569 -> Blume Beta 1.1712, Cost of Equity 12.95%, Cost of Debt after-tax 5.61%, actual capital structure 98.7% equity / 1.3% debt)
  - `CLEAN`: **11.65%** (Raw Beta 0.9006 -> Blume Beta 0.9337, Cost of Equity 11.65%, Cost of Debt after-tax 5.99%, actual capital structure 99.98% equity / 0.02% debt)
  - `EDELWEISS`: **9.60%** (Raw Beta 1.4830 -> Blume Beta 1.3220, Cost of Equity 13.78%, Cost of Debt after-tax 7.11%)"""

if old_req6_snippet in uap_text:
    uap_text = uap_text.replace(old_req6_snippet, new_req6_snippet)
    print("Requirement 6 updated in USER_APPROVAL_PACKET.md")

# Replace Requirement 7 values for LTM and CLEAN
uap_text = uap_text.replace(
    "- **LTM**: Full-Year Annual FCFF DCF Baseline = **₹2,428.22 / share** (E42 = 2428.22, C9 = 2428.22, diff = 0.0043). Explicit FCFF PV ₹17,294.07 Cr, Terminal Value PV ₹43,383.93 Cr @ WACC 13.12% linked to company's own regression Blume Beta `Beta!K13` (1.1712) and terminal g 5.0% linked to `E20`, Cash ₹13,433 Cr, Debt ₹2,187 Cr, Shares 29.62 Cr.",
    "- **LTM**: Full-Year Annual FCFF DCF Baseline = **₹2,497.80 / share** (E42 = 2497.80, C9 = 2497.80, diff = 0.0029). Explicit FCFF PV ₹17,695.13 Cr, Terminal Value PV ₹45,043.62 Cr @ WACC 12.86% linked to company's own regression Blume Beta `Beta!K13` (1.1712) and terminal g 5.0% linked to `E20`, Cash ₹13,433 Cr, Debt ₹2,187 Cr, Shares 29.62 Cr."
)
uap_text = uap_text.replace(
    "- **CLEAN**: Full-Year Manufacturing FCFF DCF Baseline = **₹468.61 / share** (E42 = 468.61, C9 = 468.61, diff = 0.0026). Explicit FCFF PV ₹1,004.22 Cr, Terminal Value PV ₹3,578.08 Cr @ WACC 11.91% linked to company's own regression Blume Beta `Beta!K13` (0.9337) and terminal g 5.0% linked to `E20`, Cash ₹401 Cr, Debt ₹2 Cr, Shares 10.63 Cr.",
    "- **CLEAN**: Full-Year Manufacturing FCFF DCF Baseline = **₹487.28 / share** (E42 = 487.28, C9 = 487.28, diff = 0.0008). Explicit FCFF PV ₹1,014.85 Cr, Terminal Value PV ₹3,765.93 Cr @ WACC 11.65% linked to company's own regression Blume Beta `Beta!K13` (0.9337) and terminal g 5.0% linked to `E20`, Cash ₹401 Cr, Debt ₹2 Cr, Shares 10.63 Cr."
)

# Update Audited Provenance Ledger in USER_APPROVAL_PACKET.md
old_ledger = """| Target Company | Worksheet | Cell Coordinates | Sourced Financial Item | Value Basis / Observation | Allowed Source Name | Approved Source URL / Provenance | Observation Date / Period |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **MM** | `Main Financials` | `I11`, `I23`, `I38` | Revenue, EBITDA, PAT | ₹159,211 Cr / ₹30,518 Cr / ₹14,074 Cr | Screener.in | `https://www.screener.in/company/M&M/consolidated/` | 31-Mar-2025 (FY25) |
| **MM** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Share Closes & NIFTY 50 | ₹3,497.20 (15-Oct close), NIFTY 25,057.35 | NSE / Yahoo Finance | `https://www.nseindia.com/` / `finance.yahoo.com` | 15-Oct-2025 |
| **MM** | `WACC-DCF` | `D18`, `J22` | Risk-Free Rate, ERP | 6.78% (India 10Y Yield), 5.50% ERP | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 15-Oct-2025 |
| **MM** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹276,014 Cr (FY25) | Screener.in | `https://www.screener.in/company/M&M/consolidated/` | FY15–FY25 Statutory Actuals |
| **LTM** | `Main Financials` | `I11`, `I23`, `I38` | Revenue, EBITDA, PAT | ₹38,008 Cr / ₹6,495 Cr / ₹4,603 Cr | Screener.in | `https://www.screener.in/company/LTIM/consolidated/` | 31-Mar-2025 (FY25) |
| **LTM** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Share Closes & NIFTY 50 | ₹5,609.50 (15-Oct close), NIFTY 25,057.35 | NSE / Yahoo Finance | `https://www.nseindia.com/` / `finance.yahoo.com` | 15-Oct-2025 |
| **LTM** | `WACC-DCF` | `D18`, `J22` | Risk-Free Rate, ERP | 6.78% (India 10Y Yield), 5.50% ERP | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 15-Oct-2025 |
| **LTM** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹30,599 Cr (FY25) | Screener.in | `https://www.screener.in/company/LTIM/consolidated/` | FY15–FY25 Statutory Actuals |
| **CLEAN** | `Main Financials` | `I11`, `I23`, `I38` | Revenue, EBITDA, PAT | ₹967 Cr / ₹388 Cr / ₹264 Cr | Screener.in | `https://www.screener.in/company/CLEAN/consolidated/` | 31-Mar-2025 (FY25) |
| **CLEAN** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Share Closes & NIFTY 50 | ₹1,041.40 (15-Oct close), NIFTY 25,057.35 | NSE / Yahoo Finance | `https://www.nseindia.com/` / `finance.yahoo.com` | 15-Oct-2025 |
| **CLEAN** | `WACC-DCF` | `D18`, `J22` | Risk-Free Rate, ERP | 6.78% (India 10Y Yield), 5.50% ERP | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 15-Oct-2025 |
| **CLEAN** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹1,599 Cr (FY25) | Screener.in | `https://www.screener.in/company/CLEAN/consolidated/` | FY15–FY25 Statutory Actuals |
| **EDELWEISS** | `Main Financials` | `I11`, `I23`, `I38` | Revenue, Operating Profit, PAT | ₹9,393 Cr / ₹3,360 Cr / ₹536 Cr | Screener.in | `https://www.screener.in/company/EDELWEISS/consolidated/` | 31-Mar-2025 (FY25) |
| **EDELWEISS** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Share Closes & NIFTY 50 | ₹108.85 (15-Oct close), NIFTY 25,057.35 | NSE / Yahoo Finance | `https://www.nseindia.com/` / `finance.yahoo.com` | 15-Oct-2025 |
| **EDELWEISS** | `WACC-DCF` | `D18`, `J22` | Risk-Free Rate, ERP | 6.78% (India 10Y Yield), 5.50% ERP | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 15-Oct-2025 |
| **EDELWEISS** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹41,283 Cr (FY25) | Screener.in | `https://www.screener.in/company/EDELWEISS/consolidated/` | FY15–FY25 Statutory Actuals |"""

new_ledger = """| Target Company | Worksheet | Cell Coordinates | Financial Metric / Input | Value Basis / Observation | Classification | Allowed Source Provider | Approved Source URL / Evidence | Observation Date / Period |
| :--- | :--- | :--- | :--- | :--- | :---: | :--- | :--- | :--- |
| **MM** | `Main Financials` | `I11`, `I23`, `I38` | Sales, EBITDA, PAT | ₹159,211 Cr / ₹30,518 Cr / ₹14,074 Cr | Sourced Fact | Screener.in | `https://www.screener.in/company/M&M/consolidated/` | 31-Mar-2025 (FY25) |
| **MM** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Prices & NIFTY 50 | ₹3,497.20 (15-Oct close), NIFTY 25,057.35 | Sourced Fact | NSE / Yahoo Finance | `https://www.nseindia.com/` / `https://finance.yahoo.com/` | 15-Oct-2025 |
| **MM** | `WACC-DCF` | `J22` | Risk-Free Rate ($R_f$) | 6.51% (India 10Y Benchmark G-Sec) | Prior Proxy | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 14-Oct-2025 (Proxy) |
| **MM** | `WACC-DCF` | `J23` | Equity Risk Premium | 5.50% | Assumption | Analyst Modeling | Modeling Assumption (unsupported on allowlisted sites) | Forward |
| **MM** | `WACC-DCF` | `E23` | Corporate Tax Rate | 25.17% | Assumption | Analyst Modeling | Base 22% + 10% surcharge + 4% cess assumption | Forward |
| **MM** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹276,014 Cr (FY25) | Sourced Fact | Screener.in | `https://www.screener.in/company/M&M/consolidated/` | FY15–FY25 Statutory Actuals |
| **LTM** | `Main Financials` | `I11`, `I23`, `I38` | Sales, EBITDA, PAT | ₹38,008 Cr / ₹6,495 Cr / ₹4,603 Cr | Sourced Fact | Screener.in | `https://www.screener.in/company/LTIM/consolidated/` | 31-Mar-2025 (FY25) |
| **LTM** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Prices & NIFTY 50 | ₹5,609.50 (15-Oct close), NIFTY 25,057.35 | Sourced Fact | NSE / Yahoo Finance | `https://www.nseindia.com/` / `https://finance.yahoo.com/` | 15-Oct-2025 |
| **LTM** | `WACC-DCF` | `J22` | Risk-Free Rate ($R_f$) | 6.51% (India 10Y Benchmark G-Sec) | Prior Proxy | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 14-Oct-2025 (Proxy) |
| **LTM** | `WACC-DCF` | `J23` | Equity Risk Premium | 5.50% | Assumption | Analyst Modeling | Modeling Assumption (unsupported on allowlisted sites) | Forward |
| **LTM** | `WACC-DCF` | `E23` | Corporate Tax Rate | 25.17% | Assumption | Analyst Modeling | Base 22% + 10% surcharge + 4% cess assumption | Forward |
| **LTM** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹30,599 Cr (FY25) | Sourced Fact | Screener.in | `https://www.screener.in/company/LTIM/consolidated/` | FY15–FY25 Statutory Actuals |
| **CLEAN** | `Main Financials` | `I11`, `I23`, `I38` | Sales, EBITDA, PAT | ₹967 Cr / ₹388 Cr / ₹264 Cr | Sourced Fact | Screener.in | `https://www.screener.in/company/CLEAN/consolidated/` | 31-Mar-2025 (FY25) |
| **CLEAN** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Prices & NIFTY 50 | ₹1,041.40 (15-Oct close), NIFTY 25,057.35 | Sourced Fact | NSE / Yahoo Finance | `https://www.nseindia.com/` / `https://finance.yahoo.com/` | 15-Oct-2025 |
| **CLEAN** | `WACC-DCF` | `J22` | Risk-Free Rate ($R_f$) | 6.51% (India 10Y Benchmark G-Sec) | Prior Proxy | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 14-Oct-2025 (Proxy) |
| **CLEAN** | `WACC-DCF` | `J23` | Equity Risk Premium | 5.50% | Assumption | Analyst Modeling | Modeling Assumption (unsupported on allowlisted sites) | Forward |
| **CLEAN** | `WACC-DCF` | `E23` | Corporate Tax Rate | 25.17% | Assumption | Analyst Modeling | Base 22% + 10% surcharge + 4% cess assumption | Forward |
| **CLEAN** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹1,599 Cr (FY25) | Sourced Fact | Screener.in | `https://www.screener.in/company/CLEAN/consolidated/` | FY15–FY25 Statutory Actuals |
| **EDELWEISS** | `Main Financials` | `I11`, `I23`, `I38` | Revenue, Op Profit, PAT | ₹9,393 Cr / ₹3,360 Cr / ₹536 Cr | Sourced Fact | Screener.in | `https://www.screener.in/company/EDELWEISS/consolidated/` | 31-Mar-2025 (FY25) |
| **EDELWEISS** | `Beta` | `C8:C112`, `F8:F112` | 104-Wk Prices & NIFTY 50 | ₹108.85 (15-Oct close), NIFTY 25,057.35 | Sourced Fact | NSE / Yahoo Finance | `https://www.nseindia.com/` / `https://finance.yahoo.com/` | 15-Oct-2025 |
| **EDELWEISS** | `WACC-DCF` | `J22` | Risk-Free Rate ($R_f$) | 6.51% (India 10Y Benchmark G-Sec) | Prior Proxy | Trading Economics | `https://tradingeconomics.com/india/government-bond-yield` | 14-Oct-2025 (Proxy) |
| **EDELWEISS** | `WACC-DCF` | `J23` | Equity Risk Premium | 5.50% | Assumption | Analyst Modeling | Modeling Assumption (unsupported on allowlisted sites) | Forward |
| **EDELWEISS** | `WACC-DCF` | `E23` | Corporate Tax Rate | 25.17% | Assumption | Analyst Modeling | Base 22% + 10% surcharge + 4% cess assumption | Forward |
| **EDELWEISS** | `BS` | `C4:M15`, `C28:M37` | 11-Year Balance Sheet Lines | Components sum to ₹41,283 Cr (FY25) | Sourced Fact | Screener.in | `https://www.screener.in/company/EDELWEISS/consolidated/` | FY15–FY25 Statutory Actuals |"""

if old_ledger in uap_text:
    uap_text = uap_text.replace(old_ledger, new_ledger)
    print("Audited Provenance Ledger updated in USER_APPROVAL_PACKET.md")
else:
    print("Warning: old_ledger block not found exactly in USER_APPROVAL_PACKET.md")

uap_path.write_text(uap_text, encoding="utf-8")
print("Saved docs/USER_APPROVAL_PACKET.md")

# 2. Update work/progress.md with Section 19
prog_path = ROOT / "work" / "progress.md"
prog_text = prog_path.read_text(encoding="utf-8")

sec19 = """
## 19. Tightly Scoped Source & Valuation Verification: Evidenced 6.51% Risk-Free Rate Proxy, Assumption Classification & PBT Valuation Trace (2026-10-18 Asia/Kolkata)
- **User Mandate**:
  1. Verify the risk-free rate: inspect WACC input cell, source URL, date, and evidence for 6.78%. Replace unevidenced 6.78% with allowlisted observation/dated proxy.
  2. Classify assumptions correctly: 5.50% ERP is an analyst assumption (not Trading Economics); 25.17% tax rate is an analyst modeling assumption (not statutory rate quote).
  3. Trace PBT correction into valuation: explain precisely why fair values were unchanged despite PBT fix, prove formula chains.
  4. Recalculate all 4 workbooks in Excel COM, recheck formula errors, balance differences, WACC, valuation outputs, sensitivities, and update provenance ledger.
- **Root Cause & Technical Implementation**:
  1. **Risk-Free Rate Evidenced Proxy**:
     - Cell: `'WACC-DCF'!J22`.
     - Allowlisted source Trading Economics (`https://tradingeconomics.com/india/government-bond-yield`) reported on 14-Oct-2025 ("India 10-Year Bond Yield Eases to Near 1-Week Low") that India's benchmark 10-year yield eased to just below 6.51% (6.51%).
     - Updated $R_f$ from unevidenced 6.78% to **6.51%** (0.0651) as a dated prior-session proxy preceding the 15-Oct-2025 cutoff.
  2. **Assumption Classifications**:
     - ERP (`'WACC-DCF'!J23` = 5.50%): Explicitly labeled as **Analyst Macroeconomic Assumption** across workbook comments and documentation.
     - Tax Rate (`'WACC-DCF'!E23` and `'Data Sheet'!K29` = 25.17%): Explicitly labeled as **Analyst Modeling Assumption** (22% base + 10% surcharge + 4% cess).
  3. **PBT-to-Valuation Mathematical Proof**:
     - MM & EDELWEISS: Value is driven by SOTP schedules (`'DCF 1st way'!G60`), aggregating FY25 segment EBITDA/PAT and listed stakes. Consolidated forward PBT does not enter SOTP by methodological design.
     - LTM & CLEAN: Value is driven by unlevered FCFF DCF starting from Operating EBIT (`Row 5`), not levered PBT (`'Data Sheet'!K28`) or PAT (`'Data Sheet'!K30`). Interest is excluded from unlevered FCFF and subtracted at the equity value bridge.
  4. **Updated Valuations & WACCs with 6.51% Rf**:
     - **MM**: WACC = **12.79%** (down from 13.00%), SOTP = **₹2,742.97 / share** (unchanged).
     - **LTM**: WACC = **12.86%** (down from 13.12%), FCFF DCF = **₹2,497.80 / share** (before: ₹2,428.22, delta: +₹69.57).
     - **CLEAN**: WACC = **11.65%** (down from 11.91%), FCFF DCF = **₹487.28 / share** (before: ₹468.61, delta: +₹18.67).
     - **EDELWEISS**: WACC = **9.60%** (down from 9.70%), SOTP = **₹130.53 / share** (unchanged).
- **Verification Results**:
  - `python work/recalculate_template_v6_models.py`: Rebuilt in Excel COM 16.0; exactly **0 formula errors** across all 64 sheets; 8/8 dynamic sensitivities pass; 64 PDFs re-exported.
  - Balance Sheet Integrity: Net integration variance $\Delta A - \Delta L \equiv 0.0000$ Cr for all 4 targets; closing balance checks match opening Screener integer rounding carryovers (MM 0.00 Cr, LTM +1.00 Cr, CLEAN -1.13 Cr, EDELWEISS +1.00 Cr).
- **Fresh Deliverable Checksums (`template_v6`)**:
  - **MM**: `a0763519d1be2094efbfd47111c03ea4e26d2bb3a935d38e02eb37286870cb15` (114,551 bytes)
  - **LTM**: `28f4dc2fa0b603d777af10cc6ad54960407cd6ee4071f7a04b0ea5894f56a6e7` (114,540 bytes)
  - **CLEAN**: `6c0dcb8f392e4db4ba179e59e4f11d199668bdcc3960a16ddad4179083840931` (112,992 bytes)
  - **EDELWEISS**: `744399a6979abcc972c69597370e341b5320f2a8cfbb9dd292f25de2515a0f7c` (116,493 bytes)
- **Governance**: `ready: false`, `human_pack_approval: pending`. PR #3 open and unmerged.
"""

if "## 19. Tightly Scoped Source & Valuation Verification" not in prog_text:
    prog_text += sec19
    prog_path.write_text(prog_text, encoding="utf-8")
    print("Section 19 added to work/progress.md")

# 3. Update work/pr3_body.md with Section 16
pr3_path = ROOT / "work" / "pr3_body.md"
pr3_text = pr3_path.read_text(encoding="utf-8")

# Replace hashes in PR3 body
pr3_text = pr3_text.replace(
    "`8b683e414e52a0e98165bbdda9d8435ef6cca0558338d087edc847227696727c`",
    "`a0763519d1be2094efbfd47111c03ea4e26d2bb3a935d38e02eb37286870cb15`"
)
pr3_text = pr3_text.replace(
    "`c8f8323b77a6610f79b77fa52dd9db612e8cb22f0f575d4ad6e0f79e8f50369a`",
    "`28f4dc2fa0b603d777af10cc6ad54960407cd6ee4071f7a04b0ea5894f56a6e7`"
)
pr3_text = pr3_text.replace(
    "`24797c0d82c914c5b0774ec100d8b7e623b831dfea5a255811def8e6f8554775`",
    "`6c0dcb8f392e4db4ba179e59e4f11d199668bdcc3960a16ddad4179083840931`"
)
pr3_text = pr3_text.replace(
    "`48d0b171bccbf5e41e0a3281b2b59c78acd1b0f6484a4fcedf9cfdfaf41b678b`",
    "`744399a6979abcc972c69597370e341b5320f2a8cfbb9dd292f25de2515a0f7c`"
)
pr3_text = pr3_text.replace("(114,010 bytes", "(114,551 bytes")
pr3_text = pr3_text.replace("(114,103 bytes", "(114,540 bytes")
pr3_text = pr3_text.replace("(112,635 bytes", "(112,992 bytes")
pr3_text = pr3_text.replace("(116,118 bytes", "(116,493 bytes")

sec16_pr = """
### 16. Source Evidence & Valuation Alignment: Evidenced 6.51% Risk-Free Rate Proxy, Assumption Classification & PBT Valuation Trace

- **Risk-Free Rate Evidenced Proxy**:
  - Cell `'WACC-DCF'!J22`: Updated from unevidenced 6.78% to **6.51%** (0.0651).
  - Sourced from allowlisted provider **Trading Economics** (`https://tradingeconomics.com/india/government-bond-yield`), supported by its 14-Oct-2025 published reading ("India 10-Year Bond Yield Eases to Near 1-Week Low") reporting the 10-year benchmark yield just below 6.51% (6.51%) as the exact dated prior-session proxy preceding the 15-Oct-2025 cutoff.
- **Strict Classification of Modeling Assumptions**:
  - **Equity Risk Premium (ERP, `'WACC-DCF'!J23` = 5.50%)**: Classified as **Analyst Macroeconomic Assumption**; not attributed to Trading Economics.
  - **Marginal Corporate Tax Rate (`'WACC-DCF'!E23` and `'Data Sheet'!K29` = 25.17%)**: Classified as **Analyst Modeling Assumption** (22% base + 10% surcharge + 4% cess).
- **PBT-to-Valuation Mathematical Proof**:
  - **M&M & EDELWEISS**: Valuation is driven by Sum-of-the-Parts (SOTP) of business segments, listed subsidiaries, and standalone net debt/cash; consolidated forward 3-statement PBT does not enter the SOTP schedule by methodological design. MM SOTP = **₹2,742.97 / share**; EDELWEISS SOTP = **₹130.53 / share**.
  - **LTM & CLEAN**: Valuation is driven by unlevered FCFF DCF starting from Operating EBIT (`Row 5`), not levered PBT (`'Data Sheet'!K28`) or PAT (`'Data Sheet'!K30`). PBT includes interest expense, which must not be subtracted in unlevered FCFF.
  - Updating $R_f$ to 6.51% reduces WACC:
    - **LTM**: WACC 13.12% -> **12.86%**, FCFF DCF fair value moves from ₹2,428.22 to **₹2,497.80 / share** (+₹69.57).
    - **CLEAN**: WACC 11.91% -> **11.65%**, FCFF DCF fair value moves from ₹468.61 to **₹487.28 / share** (+₹18.67).
- **Final Deliverables in `companies/<TARGET>/template_v6/`**:
  - `MM`: `companies/MM/template_v6/MM-FullTemplate.xlsx` (114,551 bytes, SHA-256: `a0763519d1be2094efbfd47111c03ea4e26d2bb3a935d38e02eb37286870cb15`)
  - `LTM`: `companies/LTM/template_v6/LTM-FullTemplate.xlsx` (114,540 bytes, SHA-256: `28f4dc2fa0b603d777af10cc6ad54960407cd6ee4071f7a04b0ea5894f56a6e7`)
  - `CLEAN`: `companies/CLEAN/template_v6/CLEAN-FullTemplate.xlsx` (112,992 bytes, SHA-256: `6c0dcb8f392e4db4ba179e59e4f11d199668bdcc3960a16ddad4179083840931`)
  - `EDELWEISS`: `companies/EDELWEISS/template_v6/EDELWEISS-FullTemplate.xlsx` (116,493 bytes, SHA-256: `744399a6979abcc972c69597370e341b5320f2a8cfbb9dd292f25de2515a0f7c`)
- **Review Gate & Governance**:
  - Gate status: `ready: false`, pending gate: `["human_pack_approval"]`.
  - PR #3 remains open and unmerged pending human review.
"""

if "### 16. Source Evidence & Valuation Alignment" not in pr3_text:
    pr3_text += sec16_pr
    pr3_path.write_text(pr3_text, encoding="utf-8")
    print("Section 16 added to work/pr3_body.md")
else:
    pr3_path.write_text(pr3_text, encoding="utf-8")
    print("Updated work/pr3_body.md")
