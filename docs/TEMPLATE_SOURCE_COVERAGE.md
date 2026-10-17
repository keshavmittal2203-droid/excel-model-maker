# Template-Referenced Source Inventory & Multi-Target Coverage Registry

> [!IMPORTANT]
> **PROVENANCE & METHODOLOGY NOTICE**: Every data source and website referenced in this registry derives directly from citations, hyperlinks, and data layouts established in the original uploaded workbook `originals/Marksans Keshav.xlsx` (SHA-256 `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`). All four target models preserve 100% of the 17 sheets without omission or combination.

## 1. Source Website Inventory & Template Integration

| Data Provider / Website | Template Location | Data Supported | Coverage Across 4 Targets |
|---|---|---|---|
| **[Screener.in (Mittal Analytics)](https://www.screener.in/excel/)** | `'Data Sheet'!E1 (Hyperlink)`<br>`'Data Sheet'!Rows 16-31 (P&L)`<br>`'Data Sheet'!Rows 56-72 (Balance Sheet)` | 10-year audited annual financial statements (P&L, Balance Sheet, Cash Flow), share counts, face value, capitalization ratios, peer database. | MM: Full coverage (Mar 2015 to Mar...<br>LTM: Full coverage (Mar 2015 to Mar...<br>CLEAN: Coverage from Mar 2018 to Mar ...<br>EDELWEISS: Full coverage (Mar 2015 to Mar... |
| **[National Stock Exchange of India (NSE)](https://www.nseindia.com/)** | `'WACC-DCF'!P10 (Citation text)`<br>`'WACC-DCF'!T10 ('year returns - company + nse')`<br>`'WACC-DCF'!T13 ('div yeild nse')` | Official unadjusted stock closing quotations on 15-Oct-2025, NIFTY 50 benchmark prices, 104-week weekly return time series, index dividend yields. | MM: Official close ₹3,497.20 on 15...<br>LTM: Official close ₹5,609.50 on 15...<br>CLEAN: Official close ₹1,041.40 on 15...<br>EDELWEISS: Official close ₹108.85 on 15-O... |
| **[Yahoo Finance (via API)](https://finance.yahoo.com/)** | `'WACC-DCF'!P10 (Citation text)`<br>`'Beta'!Cols B, C, F (Weekly closing prices)` | 105-week historical weekly prices and paired returns vs Nifty 50 (^NSEI) for empirical OLS regression. | MM: 105 weekly observations; OLS S...<br>LTM: 105 weekly observations; OLS S...<br>CLEAN: 105 weekly observations; OLS S...<br>EDELWEISS: 105 weekly observations; OLS S... |
| **[Investing.com & PrimeInvestor.in / Reserve Bank of India (RBI)](https://in.investing.com/ & https://primeinvestor.in/)** | `'WACC-DCF'!P12 ('(investing.com)')`<br>`'WACC-DCF'!P13 ('(prime investor.in)')`<br>`'WACC-DCF'!J22 (Risk Free Rate input)` | Indian 10-Year Government Securities (G-Sec) sovereign benchmark yield as of October 15, 2025. | MM: 6.78% (0.0678) 10-Year G-Sec y...<br>LTM: 6.78% (0.0678) 10-Year G-Sec y...<br>CLEAN: 6.78% (0.0678) 10-Year G-Sec y...<br>EDELWEISS: 6.78% (0.0678) 10-Year G-Sec y... |
| **[Damodaran Online (Aswath Damodaran, NYU Stern)](https://pages.stern.nyu.edu/~adamodar/)** | `'sectordata'!Rows 2-100 (Emerging Markets / India Sector Benchmark Dataset)` | India industry-specific benchmarks for CapEx/Depreciation, number of firms, reinvestment rates, and debt-to-capital ratios. | MM: Preserved intact; mapped to 'A...<br>LTM: Preserved intact; mapped to 'I...<br>CLEAN: Preserved intact; mapped to 'C...<br>EDELWEISS: Preserved intact; mapped to 'F... |
| **[BSE India & Statutory Audited Annual Reports](https://www.bseindia.com/)** | `Reconciliation Notes on 'DCF 1st way'!B47 and 'DCF 2nd way'!B95`<br>`'Data Sheet'!B7 (Face Value), B70 (Exact Shares)`<br>`Audited Standalone Net Cash / HoldCo Debt Bridges` | Statutory segment profitability, standalone vs consolidated debt schedules, attributable minority stakes. | MM: M&M 79th Annual Report Note 28...<br>LTM: LTIMindtree 28th Annual Report...<br>CLEAN: Clean Science 21st Annual Repo...<br>EDELWEISS: Edelweiss Financial Services 2... |

---

## 2. Input Classification by Worksheet and Company

| Worksheet | Input Parameter | Source Provider | MM Input | LTM Input | CLEAN Input | EDELWEISS Input | Classification |
|---|---|---|---|---|---|---|---|
| `'Data Sheet'!B1` | Company Name | Statutory Filings | MAHINDRA & MAHINDRA LTD | LTIMINDTREE LIMITED | CLEAN SCIENCE AND TECH LTD | EDELWEISS FINANCIAL SERVICES | Reported |
| `'Data Sheet'!B7` | Face Value (₹) | NSE / Screener | ₹5.00 | ₹1.00 | ₹1.00 | ₹1.00 | Reported |
| `'Data Sheet'!B8` | Official 15-Oct-2025 CMP | NSE Closing Quotation | ₹3,497.20 | ₹5,609.50 | ₹1,041.40 | ₹108.85 | Reported |
| `'Data Sheet'!B9` | Market Capitalization (₹ Cr) | NSE / Calculated | ₹4,34,701.96 | ₹1,66,153.39 | ₹11,070.08 | ₹10,286.33 | Reported |
| `'Data Sheet'!B6` | Implied Shares Cr | Formula `=IF(B9>0,B9/B8,0)` | 124.30 Cr | 29.62 Cr | 10.63 Cr | 94.50 Cr | Calculated |
| `'Data Sheet'!Rows 17-31` | Annual P&L History | Screener.in Consolidated | FY17–FY25 Audited | FY17–FY25 Audited | FY18–FY25 (FY17 N/A Pre-Inc) | FY17–FY25 Audited | Reported |
| `'Data Sheet'!Rows 57-72` | Annual Balance Sheet | Screener.in Consolidated | FY17–FY25 Audited | FY17–FY25 Audited | FY18–FY25 Audited | FY17–FY25 Audited | Reported |
| `'Data Sheet'!Rows 82-85` | Annual Cash Flow | Screener.in Consolidated | FY17–FY25 Audited | FY17–FY25 Audited | FY18–FY25 Audited | FY17–FY25 Audited | Reported |
| `'Beta'!Rows 8-112` | 104-Week Paired Returns | NSE / Yahoo Finance API | 105 Closes vs NIFTY | 105 Closes vs NIFTY | 105 Closes vs NIFTY | 105 Closes vs NIFTY | Reported |
| `'Beta'!O6` | OLS Regression Slope Beta | Formula `=SLOPE(D9:D112,G9:G112)` | 1.7638 | 1.2569 | 0.9006 | 1.4830 | Calculated |
| `'Beta'!K13` | Blume Adjusted Beta | Formula `=(K7*K8)+(K11*K10)` | 1.5729 | 1.1926 | 0.9254 | 1.3623 | Calculated |
| `'WACC-DCF'!J22` | Risk-Free Rate ($R_f$) | Investing.com / RBI 10Y G-Sec | 6.78% | 6.78% | 6.78% | 6.78% | Reported |
| `'WACC-DCF'!E22` | Pre-Tax Cost of Debt | Marginal Replacement Rate | 8.00% | 7.50% | 8.00% | 9.50% | Assumed |
| `'WACC-DCF'!E23` | Marginal Corporate Tax Rate | Statutory Ind AS Rate | 25.17% | 25.17% | 25.17% | 25.17% | Reported |
| `'WACC-DCF'!J24` | Beta in WACC | Formula `=J32` (Relevered) | 1.08 | 0.92 | 0.85 | 1.22 | Calculated |
| `'WACC-DCF'!J39` | Cost of Capital (WACC) | Formula `=SUMPRODUCT(I37:I38,J37:J38)` | 17.16% | 15.47% | 14.93% | 14.03% | Calculated |
| `'DCF 1st way'!E42` | Template DCF Value / Share | Formula `=E40/E41` | ₹686.81 | ₹1,506.84 | ₹181.97 | ₹71.91 | Calculated |
| `'DCF 1st way'!B47` | Institutional Reconciled Value | Statutory Segment / SOTP Notes | ₹2,742.97 / share | ₹2,811.09 / share | ₹480.89 / share | ₹130.53 / share | Calculated / SOTP |

---

## 3. Accounting Mappings and Unavailable Data Disclosures

1. **Pre-Incorporation Reporting Horizon for Clean Science (`CLEAN`)**:
   - Clean Science and Technology Limited was converted into a public limited company and published its first statutory public financials in FY18. Column B (Mar 2017) on `'Data Sheet'` is marked `0.00 / N/A (Pre-Incorporation)`. All dependent growth formulas wrap in `IFERROR(..., "-")` to prevent `#DIV/0!` errors.

2. **Zero-Debt Solvency Protection for Clean Science (`CLEAN`)**:
   - Clean Science carries virtually zero bank debt (₹2 Cr short-term facilities). On `'Ratio Analysis'!Row 34` (`CFO/Total Debt`), dividing operating cash flows by zero debt is guarded by `=IFERROR(..., "-")` to cleanly show debt-free status.

3. **Zero Dividend Periods for Edelweiss (`EDELWEISS`)**:
   - In periods where no equity dividend was declared during wholesale balance sheet cleanup, `'Ratio Analysis'!Row 9` (`Dividend Growth`) wraps in `=IFERROR(..., "-")`.

4. **Debt Perimeter & Sum-of-the-Parts Bridges**:
   - **M&M**: Consolidated borrowings of ₹129,025 Cr include ₹105,158 Cr of MMFS financial debt. Standalone Net Cash (+₹9,109 Cr) ring-fences parent auto equity, driving SOTP fair value of **₹2,742.97 / share**.
   - **EDELWEISS**: Inadmissibility of consolidated industrial FCFF is fully disclosed on `'DCF 1st way'!B47`. SOTP valuing Alternatives, ARC, Mutual Fund, Insurance, and Credit with Standalone HoldCo Net Debt (-₹3,067 Cr) drives fair value of **₹130.53 / share**.
