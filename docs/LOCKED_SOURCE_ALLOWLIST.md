# Locked Source Allowlist

This Excel-only project may use company data only from websites named in the active master workbook `originals/Marksans Keshav.xlsx` or from Screener. Do not broaden the list based on links shown by those websites, generic web search results, filings, or model-provider suggestions.

The approved websites are:

1. **Screener** — `https://www.screener.in/` (company financial statements, ratios, and market data displayed on the relevant company page).
2. **Yahoo Finance** — `https://finance.yahoo.com/` (historical prices and market-index observations; the 2025-10-15 snapshot is explicitly historical).
3. **NSE India** — `https://www.nseindia.com/` (only where the master workbook explicitly identifies NSE data).
4. **Trading Economics** — `https://tradingeconomics.com/` (only the macroeconomic series identified in the master workbook; record the observation date and whether a cited value is a proxy).
5. **Top Stock Research** — `https://www.topstockresearch.com/` (only the beta reference identified in the master workbook).
6. **Investing.com** — `https://www.investing.com/` (only the risk-free-rate reference identified in the master workbook).
7. **Prime Investor** — `https://primeinvestor.in/` (only the market-return / dividend-yield references identified in the master workbook).

The YouTube financial-modeling course is a **methodology reference only**, not a source of company or market data: [Financial Modeling Course by Keshav](https://youtube.com/playlist?list=PL3uUjzLk6PulhRop_ffNeHyK0kprzO4cT).

## Evidence and modeling rules

- Record the exact source URL, retrieval date, observation/reporting date, units, fiscal period, and consolidated/standalone basis for every external input.
- Screener is an allowed source. Do not claim that its values were independently verified against BSE/NSE statutory filings unless those filings are later added to the allowlist by the user.
- APIs may be used only when they are an access method for an approved website (for example, Yahoo Finance historical prices); name the provider and retain the request parameters and response snapshot.
- A formula-derived value must link to evidenced inputs. A forecast or valuation driver that is not reported on an approved website must be labeled as an explicit, editable analyst assumption.
- If an approved source does not provide a required line item, leave the cell blank or mark the affected analysis unavailable. Do not substitute zeros, percentages of sales, guessed peer values, other providers, or standalone values for missing consolidated data.
- Never present a prior-session observation as the exact valuation-date observation. Label it as a dated proxy and show the source evidence.
- Preserve the workbook's existing source list and its structure. Do not add BSE, statutory filings, RBI, Damodaran, company investor-relations sites, or other sources without the user's explicit authorization.
