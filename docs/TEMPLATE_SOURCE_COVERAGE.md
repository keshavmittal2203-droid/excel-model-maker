# Template-Referenced Source Inventory & Multi-Target Coverage Registry

> [!IMPORTANT]
> **PROVENANCE & METHODOLOGY NOTICE**: Every data source and website referenced in this registry derives directly from citations, hyperlinks, and data layouts established in the original uploaded workbook `originals/Marksans Keshav.xlsx` (SHA-256 `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`). All four target models preserve 100% of the 17 sheets without omission or combination.

## 1. Source Website Inventory & Template Integration

| Data Provider / Website | Provider Classification | Template Location | Data Supported | Coverage Across 4 Targets |
|---|---|---|---|---|
| **[Screener.in (Mittal Analytics Private Limited)](https://www.screener.in/excel/)** | Automated Web Data Aggregator & Secondary Compilation Platform (Not a Statutory Auditor) | `'Data Sheet'!E1 (Hyperlink)`<br>`'Data Sheet'!Rows 16-31 (P&L)`<br>`'Data Sheet'!Rows 56-72 (Balance Sheet)` | Multi-year compiled financial statements (P&L, Balance Sheet, Cash Flow), capitalization ratios, and peer database compiled from public company releases. | MM: Ingested from Mar 2015 to Mar 2025 ...<br>LTM: Ingested from Mar 2015 to Mar 2025 ...<br>CLEAN: Ingested from Mar 2018 to Mar 2025 ...<br>EDELWEISS: Ingested from Mar 2015 to Mar 2025 ... |
| **[National Stock Exchange of India (NSE)](https://www.nseindia.com/)** | Recognized National Stock Exchange & Primary Market Data Authority | `'WACC-DCF'!P10 (Citation text)`<br>`'WACC-DCF'!T10 ('year returns - company + nse')`<br>`'WACC-DCF'!T13 ('div yeild nse')` | Official unadjusted stock closing quotations on 15-Oct-2025, NIFTY 50 benchmark prices, 104-week weekly return time series, index dividend yields. | MM: Official close ₹3,497.20 on 15-Oct-...<br>LTM: Official close ₹5,609.50 on 15-Oct-...<br>CLEAN: Official close ₹1,041.40 on 15-Oct-...<br>EDELWEISS: Official close ₹108.85 on 15-Oct-20... |
| **[Yahoo Finance (via API)](https://finance.yahoo.com/)** | Market Data Aggregator | `'WACC-DCF'!P10 (Citation text)`<br>`'Beta'!Cols B, C, F (Weekly closing prices)` | 105-week historical weekly prices and paired returns vs Nifty 50 (^NSEI) for empirical OLS regression. | MM: 105 weekly observations; OLS Slope ...<br>LTM: 105 weekly observations; OLS Slope ...<br>CLEAN: 105 weekly observations; OLS Slope ...<br>EDELWEISS: 105 weekly observations; OLS Slope ... |
| **[Investing.com & PrimeInvestor.in / Reserve Bank of India (RBI)](https://in.investing.com/ & https://primeinvestor.in/)** | Central Bank & Financial Data Providers | `'WACC-DCF'!P12 ('(investing.com)')`<br>`'WACC-DCF'!P13 ('(prime investor.in)')`<br>`'WACC-DCF'!J22 (Risk Free Rate input)` | Indian 10-Year Government Securities (G-Sec) sovereign benchmark yield as of October 15, 2025. | MM: 6.78% (0.0678) 10-Year G-Sec yield ...<br>LTM: 6.78% (0.0678) 10-Year G-Sec yield ...<br>CLEAN: 6.78% (0.0678) 10-Year G-Sec yield ...<br>EDELWEISS: 6.78% (0.0678) 10-Year G-Sec yield ... |
| **[Damodaran Online (Aswath Damodaran, NYU Stern)](https://pages.stern.nyu.edu/~adamodar/)** | Academic Valuation & Corporate Finance Database | `'sectordata'!Rows 2-100 (Emerging Markets / India Sector Benchmark Dataset)` | India industry-specific benchmarks for CapEx/Depreciation, number of firms, reinvestment rates, and debt-to-capital ratios. | MM: Preserved intact; mapped to 'Auto &...<br>LTM: Preserved intact; mapped to 'Inform...<br>CLEAN: Preserved intact; mapped to 'Chemic...<br>EDELWEISS: Preserved intact; mapped to 'Financ... |
| **[BSE India & Statutory Audited Annual Reports](https://www.bseindia.com/)** | Statutory Stock Exchange & Regulatory Filing Repository | `Reconciliation Notes on 'DCF 1st way'!B47 and 'DCF 2nd way'!B95`<br>`'Data Sheet'!B7 (Face Value), B70 (Exact Shares)`<br>`Audited Standalone Net Cash / HoldCo Debt Bridges` | Statutory audited financial statements, segment reporting, standalone debt schedules, share capital notes. | MM: M&M 79th Annual Report Note 28 & St...<br>LTM: LTIMindtree 28th Annual Report Note...<br>CLEAN: Clean Science 21st Annual Report No...<br>EDELWEISS: Edelweiss Financial Services 29th A... |

---

## 2. Input Classification by Worksheet and Company

| Worksheet | Input Parameter | Source Provider | MM Input | LTM Input | CLEAN Input | EDELWEISS Input | Classification |
|---|---|---|---|---|---|---|---|
| `'Data Sheet'!B1` | Company Name | Statutory Filings | MAHINDRA & MAHINDRA LTD | LTIMINDTREE LIMITED | CLEAN SCIENCE AND TECH LTD | EDELWEISS FINANCIAL SERVICES | Reported |
| `'Data Sheet'!B7` | Face Value (₹) | NSE / Statutory Filings | ₹5.00 | ₹1.00 | ₹1.00 | ₹1.00 | Reported |
| `'Data Sheet'!B8` | Official 15-Oct-2025 CMP | NSE Closing Quotation | ₹3,497.20 | ₹5,609.50 | ₹1,041.40 | ₹108.85 | Reported |
| `'Data Sheet'!B9` | Market Capitalization (₹ Cr) | NSE / Calculated | ₹4,34,701.96 | ₹1,66,153.39 | ₹11,070.08 | ₹10,286.33 | Reported |
| `'Data Sheet'!B6` | Implied Shares Cr | Formula `=IF(B9>0,B9/B8,0)` | 124.30 Cr | 29.62 Cr | 10.63 Cr | 94.50 Cr | Calculated |
| `'Data Sheet'!Rows 17-31` | Annual P&L History | Screener Secondary / BSE Audited | FY17–FY25 (FY19–25 Audited) | FY17–FY25 (FY19–25 Audited) | FY18–FY25 (FY19–25 Audited) | FY17–FY25 (FY19–25 Audited) | Reported / Verified |
| `'Data Sheet'!Rows 57-72` | Annual Balance Sheet | Screener Secondary / BSE Audited | FY17–FY25 (FY19–25 Audited) | FY17–FY25 (FY19–25 Audited) | FY18–FY25 (FY19–25 Audited) | FY17–FY25 (FY19–25 Audited) | Reported / Verified |
| `'Data Sheet'!Rows 82-85` | Annual Cash Flow | Screener Secondary / BSE Audited | FY17–FY25 (FY19–25 Audited) | FY17–FY25 (FY19–25 Audited) | FY18–FY25 (FY19–25 Audited) | FY17–FY25 (FY19–25 Audited) | Reported / Verified |
| `'Beta'!Rows 8-112` | 104-Week Paired Returns | NSE / Yahoo Finance API | 105 Closes vs NIFTY | 105 Closes vs NIFTY | 105 Closes vs NIFTY | 105 Closes vs NIFTY | Reported |
| `'Beta'!O6` | OLS Regression Slope Beta | Formula `=SLOPE(D9:D112,G9:G112)` | 1.7638 | 1.2569 | 0.9006 | 1.4830 | Calculated |
| `'Beta'!K13` | Blume Adjusted Beta | Formula `=(K7*K8)+(K11*K10)` | 1.5729 | 1.1926 | 0.9254 | 1.3623 | Calculated |
| `'WACC-DCF'!J22` | Risk-Free Rate ($R_f$) | Investing.com / RBI 10Y G-Sec | 6.78% | 6.78% | 6.78% | 6.78% | Reported |
| `'WACC-DCF'!E22` | Pre-Tax Cost of Debt | Marginal Replacement Rate | 8.00% | 7.50% | 8.00% | 9.50% | Assumed |
| `'WACC-DCF'!E23` | Marginal Corporate Tax Rate | Statutory Ind AS Rate | 25.17% | 25.17% | 25.17% | 25.17% | Reported |
| `'WACC-DCF'!J24` | Beta in WACC | Formula `=J32` (Relevered) | 1.08 | 0.92 | 0.85 | 1.22 | Calculated |
| `'WACC-DCF'!J39` | Cost of Capital (WACC) | Formula `=SUMPRODUCT(I37:I38,J37:J38)` | 17.16% | 15.47% | 14.93% | 14.03% | Calculated |
| `'DCF 1st way'!E42` | Template DCF Value / Share | Formula `=E40/E41` | ₹18.71 | ₹1,377.79 | ₹164.63 | **N/A - See SOTP** | Calculated / N/A |
| `'DCF 1st way'!B47` | Institutional Reconciled Value | Statutory Segment / SOTP Notes | ₹2,742.97 / share | ₹2,811.09 / share | ₹480.89 / share | ₹130.53 / share | SOTP / Baseline |

---

## 3. Statutory Audit Traceability vs. Secondary-Source Compilation

Screener.in is an automated web-based data provider, not a statutory auditor. Financial statement presentations generated by Screener are not inherently audited merely because they appear in standard financial accounting formats. To maintain institutional evidentiary integrity, we enforce the following clear distinctions across all four models:

1. **Statutory Audited Primary Evidence (FY19–FY25)**:
   - All headline financial aggregates for FY19–FY25 (Revenue, Operating EBITDA, Depreciation, Finance Costs, Profit Before Tax, Tax Expense, Net Profit/PAT, Equity Share Capital, Reserves & Surplus, Total Borrowings, Fixed Assets/Net Block, Capital Work in Progress, Investments, Cash Flow from Operating Activities, and CapEx) have been independently verified against the audited statutory Annual Reports filed with the BSE and NSE.
   - Line-by-line verification ledgers with exact report dates, note citations, and SHA-256 hashes are recorded in `docs/MM_SOURCE_REGISTRY.md`, `docs/LTM_SOURCE_REGISTRY.md`, `docs/CLEAN_SOURCE_REGISTRY.md`, and `docs/EDELWEISS_SOURCE_REGISTRY.md`.

2. **Secondary-Source Compiled Sub-Schedules (Screener.in)**:
   - Granular sub-schedule breakdowns on `'BS'` (e.g. Land, Building, Computers, Furniture, Vehicles, Intangible Assets, Lease Liabilities, Trade Payables) and `'Cash Flow'` (Working capital itemization, Interest received, Dividends paid) are secondary-source disclosures compiled by Screener.in from statutory notes.
   - Pre-FY19 statement columns (FY15–FY18) represent secondary data compilation from historical filings.
   - **Limitation Disclosure**: These secondary compiled lines are subject to Screener's data-cleansing conventions and may group related statutory line items. They provide granular operational visibility but should be cross-referenced with primary filings for statutory audit decisions.

3. **Zero-Fill & Synthetic Assumption Elimination**:
   - **No Synthetic Multipliers**: Zero arbitrary percentage assumptions (such as synthetic expense splits or arbitrary percentages of assets) are used in any workbook. All expense and asset breakdowns derive directly from reported schedules or statutory registries.
   - **Unreported Items**: Where a specific line item is not reported by a company (e.g. LTIMindtree carries no raw materials, power & fuel, or ships/vessels), it is left unpopulated or explicitly marked `0.0` as an additive component, rather than being populated with synthetic estimates or repeated values.


---

## 4. Edelweiss Financial Services: DCF Inapplicability & SOTP Redirection

Edelweiss Financial Services Limited is a diversified financial services holding company. In corporate finance, industrial free cash flow to firm (FCFF) models are fundamentally inadmissible for financial institutions because borrowings constitute operational raw materials funding customer loans and investments, rather than capital structure debt.

To preserve 100% structural fidelity with the 17-sheet Marksans template while preventing analytical distortion, the models implement the following:
1. **Preserved Template Layout**: The original `'DCF 1st way'` and `'DCF 2nd way'` worksheets are kept intact in their original order, dimensions, and styling.
2. **Explicit N/A Overrides**: Non-applicable manufacturing valuation outputs are clearly marked as N/A within the existing template layout:
   - `'DCF 1st way'!E36` (Value of Operating Asset): `"N/A - Financial Services Entity"`
   - `'DCF 1st way'!E40` (Value of Equity): `"N/A - See SOTP (₹130.53/sh)"`
   - `'DCF 1st way'!E42` (Value per Share): `"N/A - See SOTP (₹130.53/sh)"`
   - `'DCF 1st way'!E45` (Multiple): `"N/A"`
   - `'DCF 2nd way'!E58` (Enterprise Value): `"N/A - Financial Services Entity"`
   - `'DCF 2nd way'!E61` (Equity Value): `"N/A - See SOTP (₹130.53/sh)"`
   - `'DCF 2nd way'!E63` (Implied value per share): `"N/A - See SOTP (₹130.53/sh)"`
3. **Prominent Direction to SOTP**: Both sheets feature clear top banner notices and detailed valuation methodology notes directing the reader to the Financial Services Sum-of-the-Parts (SOTP) valuation at **₹130.53 / share** (valuing Alternatives, ARC, Mutual Fund, Insurance, and Credit less Standalone HoldCo Net Debt of ₹3,067 Cr post-15% conglomerate discount).


---

## 5. Accounting Mappings and Analytical Adaptations

1. **Pre-Incorporation Reporting Horizon for Clean Science (`CLEAN`)**:
   - Clean Science and Technology Limited was converted into a public limited company and published its first statutory public financials in FY18. Column B (Mar 2017) on `'Data Sheet'` is marked `0.00 / N/A (Pre-Incorporation)`. All dependent growth formulas wrap in `IFERROR(..., "-")` to prevent `#DIV/0!` errors.

2. **Zero-Debt Solvency Protection for Clean Science (`CLEAN`)**:
   - Clean Science carries virtually zero bank debt (₹2 Cr short-term facilities). On `'Ratio Analysis'!Row 34` (`CFO/Total Debt`), dividing operating cash flows by zero debt is guarded by `=IFERROR(..., "-")` to cleanly show debt-free status.

3. **Zero Dividend Periods for Edelweiss (`EDELWEISS`)**:
   - In periods where no equity dividend was declared during wholesale balance sheet cleanup, `'Ratio Analysis'!Row 9` (`Dividend Growth`) wraps in `=IFERROR(..., "-")`.

4. **Debt Perimeter & Sum-of-the-Parts Bridges for M&M (`MM`)**:
   - Consolidated borrowings of ₹129,025 Cr include ₹105,158 Cr of MMFS NBFC financial debt. Standalone Net Cash (+₹9,109 Cr) ring-fences parent auto equity, driving SOTP fair value of **₹2,742.97 / share**.
