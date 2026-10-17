# Locked Source Allowlist Derived from `originals/Marksans Keshav.xlsx`

> [!IMPORTANT]
> **STRICT SOURCE GOVERNANCE INVARIANT**: Only websites and data providers explicitly identified, hyperlinked, or cited within the master reference workbook `originals/Marksans Keshav.xlsx` (SHA-256 `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`) are permitted for target company data population. No external website, API, or downloaded dataset outside this allowlist is permitted.
>
> **MODELING METHODOLOGY REFERENCE EXCLUSION**: The YouTube course [Financial Modeling Course by Keshav](https://youtube.com/playlist?list=PL3uUjzLk6PulhRop_ffNeHyK0kprzO4cT&si=HtwfdWpNk-a0KSp7) is recognized solely as an analytical modeling-methodology reference; it is strictly **NOT** a numerical data source.

---

## 1. Inventory of Allowed Data Providers & In-Workbook Citations

| # | Data Provider / Website | Canonical URL / Domain | Exact Location in Template | Data Sourced in Model | Allowlist Classification |
|---|---|---|---|---|---|
| 1 | **Screener.in (Mittal Analytics Pvt Ltd)** | `https://www.screener.in/excel/`<br>`https://www.screener.in/company/<TICKER>/` | `'Data Sheet'!E1 (Hyperlink)`<br>`'Data Sheet'!Rows 16–31 (P&L)`<br>`'Data Sheet'!Rows 56–72 (Balance Sheet)`<br>`'BS'!Rows 4–37`<br>`'Cash Flow'!Rows 5–31` | Multi-year historical P&L, balance sheets, cash flows, CMP, market cap, equity face value, and segment sub-schedules. | **Secondary Data Compiler / Aggregator** (Subject to Screener compilation conventions; aggregate statements verified against statutory filings). |
| 2 | **Yahoo Finance** | `https://finance.yahoo.com` | `'WACC-DCF'!P10`: `(Yahoo Finance / NSE website)`<br>`'WACC-DCF'!T10`: `year returns - company + nse` | 105-week historical weekly adjusted closing prices for target equity and NIFTY 50 index for OLS regression beta calculation on `'Beta'`. | **Market Pricing & Weekly Returns Provider**. |
| 3 | **National Stock Exchange of India (NSE)** | `https://www.nseindia.com` | `'WACC-DCF'!P10`: `(Yahoo Finance / NSE website)`<br>`'WACC-DCF'!T10`: `year returns - company + nse`<br>`'WACC-DCF'!T13`: `div yeild nse` | Official unadjusted closing price on valuation snapshot date (15-Oct-2025), NIFTY 50 index historical returns, index dividend yield. | **Primary Stock Exchange Quotation & Benchmark Provider**. |
| 4 | **Investing.com** | `https://www.investing.com` | `'WACC-DCF'!P12`: `(investing.com)`<br>`'WACC-DCF'!T12`: `risk free rate` | India 10-Year Government Bond Yield (Risk-Free Rate $R_f$: 6.78% as of 15-Oct-2025). | **Sovereign Fixed Income Benchmark Provider**. |
| 5 | **Trading Economics** | `https://www.tradingeconomics.com` | `'WACC-DCF'!P9`: `(Trading Economics)` | India 10Y Benchmark Sovereign Yield and macroeconomic indicators. | **Macroeconomic Benchmark Provider**. |
| 6 | **Top Stock Research** | `https://www.topstockresearch.com` | `'WACC-DCF'!P11`: `( top stock research)`<br>`'WACC-DCF'!T11`: `beta` | Sector and peer beta reference benchmarks. | **Secondary Equity Analytics Provider**. |
| 7 | **Prime Investor** | `https://primeinvestor.in` | `'WACC-DCF'!P13`: `(prime investor.in)` | Equity risk premium, NIFTY historical annualized market returns ($R_m$), dividend yield benchmarks. | **Market Return & Risk Premium Provider**. |
| 8 | **Statutory Audited Filings (BSE / NSE)** | `https://www.bseindia.com`<br>`https://www.nseindia.com` | Referenced in statutory registries and verified in `'Data Sheet'!A3` | Audited statutory annual reports (FY19–FY25), note disclosures, auditor reports, and segment breakouts. | **Primary Statutory Audited Filing Provider** (Audited by B S R & Co. LLP / S.R. Batliboi & Co. LLP). |

---

## 2. In-Workbook Data Classification Rules

In accordance with the project instructions, every material numeric value in each target workbook is classified into one of three strict categories or marked unavailable:

1. **Website-Sourced Input**:
   - Directly ingested from one of the 8 allowlisted websites above.
   - Sourced cells contain cell comments or notes documenting: URL, retrieval date, reporting period, units, and accounting basis.
   - Screener figures are explicitly disclosed as secondary compiled data; aggregate primary numbers are cross-referenced to statutory BSE/NSE filings.
2. **Formula-Derived**:
   - Calculated dynamically using Excel formulas (`SUM`, `AVERAGE`, `SLOPE`, `INTERCEPT`, discounting present values, growth rates).
   - Zero hardcoded calculation results; formula cells point directly to inputs or dependent schedules.
3. **Explicit Assumption**:
   - Clearly labeled, editable inputs with economic rationale, units, and applicable forecast period.
   - Examples include: Expected 5Y Growth Rate (`'DCF 1st way'!E19`), Long-Term Terminal Growth Rate (`'DCF 1st way'!E20`), Reinvestment Rates (`'DCF 1st way'!F11:K11`), Valuation Multiples, and Pre-Tax Cost of Debt (`'WACC-DCF'!E22`).
4. **Unavailable / N/A**:
   - Where an allowlisted website does not report a specific sub-item (e.g. Clean Science pre-FY18 financials prior to public conversion, or IT services raw material costs), the item is left unpopulated or explicitly marked `0.0` as an additive component. No synthetic interpolation or arbitrary multiplier splits are permitted.
