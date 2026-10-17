# Marksans Package Financial Provenance Ledger

## 1. Valuation Snapshot & Period Cutoff Rules

- **Historical Financials Cutoff**: 2025-03-31 (FY18-FY25 (8 Fiscal Years ending 31 March 2025))
- **Market Retrieval Snapshot Date**: 2025-10-05 (Screener.in metadata in template records retrieval date 2025-10-05 (Sunday); preceding market session was Friday 2025-10-03. Primary trade confirmation logs are absent from originals/.)
- **Underlying Trading Session Observation Date**: Unverified (Friday 2025-10-03 inferred; exchange trade confirmation logs not in originals/) (Friday closing prices; market closed on retrieval date)
- **Forecast Horizon Start**: FY2026E (Period ending 2026-03-31)

### Valuation Methodology & Aggregation Notes

- **Nifty 50 Return Methodology**: Sheet RM!F10 (15.15%) is the 26-year (2000-2025) arithmetic average of annual percentage returns of the Nifty 50 Price Return (PR) index present in the template. Because the PR index excludes dividends, adding the 1.22% dividend yield assumption (RM!F11) yields Rm = 16.37% (arithmetic average). Primary NSE historical index time-series logs are not separately uploaded in originals/.
- **Target Capital Structure Methodology**: Target debt weight (2.24%) and equity weight (97.76%) are derived from the median capital structure of five listed peers (Sun Pharma, Divi's, Torrent, Zydus, Laurus) in WACC-DCF!H18. This represents an institutional modeling choice following Damodaran pure-play valuation guidelines rather than an absolute empirical fact.

### Deliberate Period Exceptions & Effects

#### Trailing Fiscal Year vs Market Pricing Cutoff
- **Rationale**: Balance sheet debt (Rs 343.38 Cr) and cash (Rs 704.21 Cr) are anchored at FY25 balance sheet audit date (2025-03-31), whereas Market Cap (Rs 14,002.74 Cr) reflects the October 2025 market price. This is institutional corporate finance standard practice (using trailing audited balance sheet with prevailing market value equity).
- **Effect**: Current debt/equity ratio (2.39% debt / 97.61% equity) reflects prevailing market capitalization against audited debt book value.

#### Target vs Current Capital Structure in DCF Discounting
- **Rationale**: Per Damodaran valuation methodology, the DCF model discounts future free cash flows at the target peer-median capital structure (2.24% debt / 97.76% equity) rather than current capital structure (2.39% debt / 97.61% equity).
- **Effect**: Target WACC is 12.11%, whereas current-weighted WACC would be 12.10% (0.01% / 1 bp difference).

## 2. Complete Provenance Ledger

| Metric | Source Cell / Part | Exact Input / Formula | Exact Value | Displayed | Category | Provenance Status | Verification Notes |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Risk-Free Rate (Rf)** | `WACC-DCF!J22` | `0.073` | `0.073` | `7.30%` | market_benchmark | `inherited_template_value` | Sourced market input present in template WACC-DCF!J22 tagged '(Latest)'; primary RBI/CCIL trade log not uploaded. |
| **Total Market Return (Rm)** | `RM!F12` | `=SUM(F10:F11)` | `0.1636923076923077` | `16.37%` | calculated_estimate | `calculated_estimate` | Hybrid: 26 annual return observations are historical facts; 1.22% dividend yield is an analyst constant assumption. |
| **Equity Risk Premium (ERP)** | `WACC-DCF!J23` | `=RM!F12-J22` | `0.0906923076923077` | `9.07%` | calculated_estimate | `calculated_estimate` | Exact model identity evaluated from RM!F12 and WACC-DCF!J22. |
| **Peer Comps Median Unlevered Beta** | `WACC-DCF!J18` | `=MEDIAN(J11:J15)` | `0.5378912172041949` | `0.538` | peer_aggregation | `calculated_estimate` | Exact peer aggregation derived from peerdata G5:G9 and E5:E9. |
| **Target Debt-to-Equity Ratio** | `WACC-DCF!E34` | `=E30/E31` | `0.02291100842442241` | `0.0229` | peer_aggregation | `calculated_estimate` | Derived from peer median Debt/Capital (2.2398%). |
| **Target Levered Beta (Hamada)** | `WACC-DCF!J24` | `=J32` | `0.5465177583503465` | `0.547` | calculated_estimate | `calculated_estimate` | Re-levered pure-play beta. Replaces single-stock raw regression beta (1.208) per Damodaran valuation standard. |
| **Cost of Equity (Ke)** | `WACC-DCF!J25` | `=J22+J24*(J23)` | `0.12256495669961989` | `12.26%` | calculated_estimate | `calculated_estimate` | Evaluated with unrounded floats. Discrepancy in earlier report was caused by manual rounded-string substitution (0.547 * 9.07% vs 0.546518 * 9.06923%). |
| **Pre-tax Cost of Debt** | `WACC-DCF!E22` | `0.08` | `0.08` | `8.00%` | analyst_assumption | `analyst_assumption` | Analyst assumption. Historical effective interest (Interest / Borrowings = 11.66 / 321.91 = 3.62%) reflects subsidized export credit, not marginal replacement borrowing cost. |
| **Marginal Corporate Tax Rate** | `WACC-DCF!E23` | `=F18` | `0.3` | `30.00%` | market_benchmark | `analyst_assumption` | Sourced statutory benchmark. WACC-DCF!B36 note 1 confirms country marginal tax rate. |
| **After-Tax Cost of Debt (Kd)** | `WACC-DCF!E24` | `=E22*(1-E23)` | `0.055999999999999994` | `5.60%` | calculated_estimate | `calculated_estimate` | Exact identity in WACC-DCF!E24. |
| **Total Debt / Borrowings (Current)** | `Data Sheet / WACC-DCF!Data Sheet!K59 / WACC-DCF!C30` | `343.38` | `343.38` | `₹343 Cr` | historical_fact | `inherited_template_value` | Total debt from template Data Sheet!K59 / BS!H20; primary statutory annual report filing not separately uploaded. |
| **Market Capitalisation (Current)** | `WACC-DCF!C31` | `14002.74` | `14002.74` | `₹14,003 Cr` | market_benchmark | `calculated_estimate` | Calculated as CMP (Rs 309.10) * 453.16M shares. CMP is an inherited template snapshot value from Sunday 2025-10-05 Screener capture. |
| **Target Weight of Debt (Peer Median)** | `WACC-DCF!E30` | `=H18` | `0.022397851069871626` | `2.24%` | peer_aggregation | `calculated_estimate` | Used directly in DCF WACC weighting. Represents an institutional modeling choice following Damodaran pure-play methodology, distinct from current book debt weight (2.39%). |
| **Target Weight of Equity (Peer Median)** | `WACC-DCF!E31` | `=E32-E30` | `0.9776021489301284` | `97.76%` | peer_aggregation | `calculated_estimate` | Used directly in DCF WACC weighting. Represents an institutional modeling choice following Damodaran pure-play methodology, distinct from current market equity weight (97.61%). |
| **WACC (Target-Weighted)** | `WACC-DCF!J39` | `=SUMPRODUCT(I37:I38,J37:J38)` | `0.12107404471298937` | `12.11%` | calculated_estimate | `calculated_estimate` | Exact model identity verified with zero rounding ambiguity. |
| **Geographic Revenue Mix (Unavailable — Pending Primary Filing)** | `Slide 1 Shape 25 (Chart 0)!N/A` | `None (Suppressed Shape 25)` | `None` | `Geographic revenue breakdown unavailable — primary filing required` | deck_visual_content | `unresolved_claim` | UNRESOLVED PROVENANCE: Missing primary document in repository originals. Chart is suppressed and unevidenced numeric payload removed from candidate deck. |
| **Historical Revenue (FY21-FY25)** | `Data Sheet!F17:J17 / F28:J28` | `Data Sheet annual revenue row` | `5 items` | `FY21: ₹1,376 Cr, FY22: ₹1,491 Cr, FY23: ₹1,852 Cr, FY24: ₹2,177 Cr, FY25: ₹2,623 Cr` | historical_fact | `inherited_template_value` | Template historical series. Chart retitled to 'Historical Revenue (₹ Crore)' from misleading 'Revenue by Segment'. |
| **Historical Balance Sheet Ratios (FY18-FY25)** | `Data Sheet & BS!Data Sheet rows 59, 60, 65, 68 and BS trade payables` | `Screener.in standard balance sheet formulas` | `{'cr': [3.18, 3.33, 2.54, 2.95, 3.73, 5.26, 5.03, 4.95], 'qr': [1.64, 1.5, 1.51, 1.63, 2.42, 3.75, 3.36, 3.07]}` | `CR: 3.18 to 4.95; QR: 1.64 to 3.07` | historical_fact | `inherited_template_value` | Fully mapped in Table 0 and 100% reconciled with native Slide 16 line charts. |
