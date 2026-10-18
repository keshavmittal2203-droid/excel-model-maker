# Template V4 Material Numeric Value Classification & Audit Ledger

> [!IMPORTANT]
> Every material numeric value across the four models is classified into one of four verified categories:
> 1. **website_sourced_input**: Entered from an allowlisted website (with documented URL, period, and basis).
> 2. **formula_derived**: Dynamically calculated from inputs or schedules (zero hardcoded values).
> 3. **explicit_assumption**: Labeled, editable input with rationale, units, and period.
> 4. **unavailable_na**: Non-applicable item or unpopulated reporting horizon.

| Company | Sheet | Cell | Value | Category | Source Provider | Allowed URL | Period | Units | Accounting Basis |
|---|---|---|---|---|---|---|---|---|---|
| **MM** | `Data Sheet` | `B7` | `5.0` | `website_sourced_input` | Screener.in / BSE Statutory Filings | https://www.screener.in/company/M&M/consolidated/ | As of 15-Oct-2025 | ₹/share | Statutory Face Value |
| **MM** | `Data Sheet` | `B8` | `3497.2` | `website_sourced_input` | National Stock Exchange of India (NSE) / Yahoo Finance | https://www.nseindia.com | 2025-10-15 | ₹/share | Official Unadjusted Close |
| **MM** | `Data Sheet` | `B9` | `434701.96` | `website_sourced_input` | Screener.in / NSE | https://www.screener.in/company/M&M/consolidated/ | 2025-10-15 | ₹ Cr | Market Capitalization |
| **MM** | `Data Sheet` | `B6` | `=IF(B9>0, B9/B8, 0)` | `formula_derived` | Template Master Formula | N/A | 2025-10-15 | Cr shares | Implied share count = Market Cap / CMP |
| **MM** | `Data Sheet` | `B17` | `83773.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2017 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `C17` | `92094.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2018 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `D17` | `104721.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2019 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `E17` | `75382.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2020 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `F17` | `74278.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2021 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `G17` | `90171.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2022 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `H17` | `121362.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2023 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `I17` | `139078.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2024 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `J17` | `159211.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2025 | ₹ Cr | Revenue from Operations |
| **MM** | `Data Sheet` | `K17` | `198639.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | Mar 2026 | ₹ Cr | Revenue from Operations |
| **MM** | `Beta` | `K8` | `0.75` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on raw empirical beta |
| **MM** | `Beta` | `K11` | `0.25` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on market beta (1.0) |
| **MM** | `WACC-DCF` | `C30` | `129025.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/M&M/consolidated/ | FY25 | ₹ Cr | Total Debt |
| **MM** | `WACC-DCF` | `C31` | `434701.96` | `website_sourced_input` | NSE / Screener.in | https://www.nseindia.com | 2025-10-15 | ₹ Cr | Market Capitalization |
| **MM** | `WACC-DCF` | `E22` | `0.08` | `explicit_assumption` | Weighted Average Borrowing Cost from Statutory Notes | https://www.bseindia.com | FY26E | Rate | Pre-tax cost of debt |
| **MM** | `WACC-DCF` | `E23` | `0.2517` | `explicit_assumption` | Indian Corporate Tax Law (Section 115BAA) | https://incometaxindia.gov.in | Statutory | Rate | Base 22% + 10% Surcharge + 4% Cess = 25.168% |
| **MM** | `WACC-DCF` | `J22` | `0.0678` | `website_sourced_input` | Investing.com / Trading Economics | https://www.investing.com | 2025-10-15 | Rate | India 10Y Benchmark G-Sec Yield |
| **MM** | `DCF 1st way` | `E19` | `0.12` | `explicit_assumption` | Research Forecast Driver | N/A | FY26E-FY30E | Growth Rate | Expected operational FCFF CAGR |
| **MM** | `DCF 1st way` | `E20` | `0.05` | `explicit_assumption` | Macroeconomic Terminal Growth Constraint | N/A | Terminal | Growth Rate | Long-term sustainable terminal growth rate capped below nominal GDP growth |
| **MM** | `DCF 1st way` | `E44` | `3497.2` | `website_sourced_input` | National Stock Exchange of India (NSE) | https://www.nseindia.com | 2025-10-15 | ₹/share | 15-Oct-2025 Official Close |
| **MM** | `DCF 1st way` | `F11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **MM** | `DCF 1st way` | `G11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **MM** | `DCF 1st way` | `H11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **MM** | `DCF 1st way` | `I11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **MM** | `DCF 1st way` | `J11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **MM** | `DCF 1st way` | `K11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **MM** | `DCF 1st way` | `E42` | `==G60` | `formula_derived` | Single Source-of-Truth Primary Valuation Output | N/A | 2025-10-15 | ₹/share | Direct link to verified primary valuation schedule |
| **LTM** | `Data Sheet` | `B7` | `1.0` | `website_sourced_input` | Screener.in / BSE Statutory Filings | https://www.screener.in/company/LTM/consolidated/ | As of 15-Oct-2025 | ₹/share | Statutory Face Value |
| **LTM** | `Data Sheet` | `B8` | `5609.5` | `website_sourced_input` | National Stock Exchange of India (NSE) / Yahoo Finance | https://www.nseindia.com | 2025-10-15 | ₹/share | Official Unadjusted Close |
| **LTM** | `Data Sheet` | `B9` | `166153.39` | `website_sourced_input` | Screener.in / NSE | https://www.screener.in/company/LTM/consolidated/ | 2025-10-15 | ₹ Cr | Market Capitalization |
| **LTM** | `Data Sheet` | `B6` | `=IF(B9>0, B9/B8, 0)` | `formula_derived` | Template Master Formula | N/A | 2025-10-15 | Cr shares | Implied share count = Market Cap / CMP |
| **LTM** | `Data Sheet` | `B17` | `6501.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2017 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `C17` | `7306.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2018 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `D17` | `9446.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2019 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `E17` | `10879.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2020 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `F17` | `12370.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2021 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `G17` | `26109.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2022 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `H17` | `33183.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2023 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `I17` | `35517.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2024 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `J17` | `38008.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2025 | ₹ Cr | Revenue from Operations |
| **LTM** | `Data Sheet` | `K17` | `42308.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | Mar 2026 | ₹ Cr | Revenue from Operations |
| **LTM** | `Beta` | `K8` | `0.75` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on raw empirical beta |
| **LTM** | `Beta` | `K11` | `0.25` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on market beta (1.0) |
| **LTM** | `WACC-DCF` | `C30` | `2187.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/LTM/consolidated/ | FY25 | ₹ Cr | Total Debt |
| **LTM** | `WACC-DCF` | `C31` | `166153.39` | `website_sourced_input` | NSE / Screener.in | https://www.nseindia.com | 2025-10-15 | ₹ Cr | Market Capitalization |
| **LTM** | `WACC-DCF` | `E22` | `0.075` | `explicit_assumption` | Weighted Average Borrowing Cost from Statutory Notes | https://www.bseindia.com | FY26E | Rate | Pre-tax cost of debt |
| **LTM** | `WACC-DCF` | `E23` | `0.2517` | `explicit_assumption` | Indian Corporate Tax Law (Section 115BAA) | https://incometaxindia.gov.in | Statutory | Rate | Base 22% + 10% Surcharge + 4% Cess = 25.168% |
| **LTM** | `WACC-DCF` | `J22` | `0.0678` | `website_sourced_input` | Investing.com / Trading Economics | https://www.investing.com | 2025-10-15 | Rate | India 10Y Benchmark G-Sec Yield |
| **LTM** | `DCF 1st way` | `E19` | `0.11` | `explicit_assumption` | Research Forecast Driver | N/A | FY26E-FY30E | Growth Rate | Expected operational FCFF CAGR |
| **LTM** | `DCF 1st way` | `E20` | `0.05` | `explicit_assumption` | Macroeconomic Terminal Growth Constraint | N/A | Terminal | Growth Rate | Long-term sustainable terminal growth rate capped below nominal GDP growth |
| **LTM** | `DCF 1st way` | `E44` | `5609.5` | `website_sourced_input` | National Stock Exchange of India (NSE) | https://www.nseindia.com | 2025-10-15 | ₹/share | 15-Oct-2025 Official Close |
| **LTM** | `DCF 1st way` | `F11` | `0.35` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **LTM** | `DCF 1st way` | `G11` | `0.35` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **LTM** | `DCF 1st way` | `H11` | `0.35` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **LTM** | `DCF 1st way` | `I11` | `0.35` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **LTM** | `DCF 1st way` | `J11` | `0.35` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **LTM** | `DCF 1st way` | `K11` | `0.35` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **LTM** | `DCF 1st way` | `E42` | `==G60` | `formula_derived` | Single Source-of-Truth Primary Valuation Output | N/A | 2025-10-15 | ₹/share | Direct link to verified primary valuation schedule |
| **CLEAN** | `Data Sheet` | `B7` | `1.0` | `website_sourced_input` | Screener.in / BSE Statutory Filings | https://www.screener.in/company/CLEAN/consolidated/ | As of 15-Oct-2025 | ₹/share | Statutory Face Value |
| **CLEAN** | `Data Sheet` | `B8` | `1041.4` | `website_sourced_input` | National Stock Exchange of India (NSE) / Yahoo Finance | https://www.nseindia.com | 2025-10-15 | ₹/share | Official Unadjusted Close |
| **CLEAN** | `Data Sheet` | `B9` | `11070.08` | `website_sourced_input` | Screener.in / NSE | https://www.screener.in/company/CLEAN/consolidated/ | 2025-10-15 | ₹ Cr | Market Capitalization |
| **CLEAN** | `Data Sheet` | `B6` | `=IF(B9>0, B9/B8, 0)` | `formula_derived` | Template Master Formula | N/A | 2025-10-15 | Cr shares | Implied share count = Market Cap / CMP |
| **CLEAN** | `Data Sheet` | `B17` | `0.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2017 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `C17` | `241.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2018 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `D17` | `393.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2019 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `E17` | `419.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2020 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `F17` | `512.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2021 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `G17` | `685.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2022 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `H17` | `936.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2023 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `I17` | `791.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2024 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `J17` | `967.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2025 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Data Sheet` | `K17` | `957.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | Mar 2026 | ₹ Cr | Revenue from Operations |
| **CLEAN** | `Beta` | `K8` | `0.75` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on raw empirical beta |
| **CLEAN** | `Beta` | `K11` | `0.25` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on market beta (1.0) |
| **CLEAN** | `WACC-DCF` | `C30` | `2.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/CLEAN/consolidated/ | FY25 | ₹ Cr | Total Debt |
| **CLEAN** | `WACC-DCF` | `C31` | `11070.08` | `website_sourced_input` | NSE / Screener.in | https://www.nseindia.com | 2025-10-15 | ₹ Cr | Market Capitalization |
| **CLEAN** | `WACC-DCF` | `E22` | `0.08` | `explicit_assumption` | Weighted Average Borrowing Cost from Statutory Notes | https://www.bseindia.com | FY26E | Rate | Pre-tax cost of debt |
| **CLEAN** | `WACC-DCF` | `E23` | `0.2517` | `explicit_assumption` | Indian Corporate Tax Law (Section 115BAA) | https://incometaxindia.gov.in | Statutory | Rate | Base 22% + 10% Surcharge + 4% Cess = 25.168% |
| **CLEAN** | `WACC-DCF` | `J22` | `0.0678` | `website_sourced_input` | Investing.com / Trading Economics | https://www.investing.com | 2025-10-15 | Rate | India 10Y Benchmark G-Sec Yield |
| **CLEAN** | `DCF 1st way` | `E19` | `0.14` | `explicit_assumption` | Research Forecast Driver | N/A | FY26E-FY30E | Growth Rate | Expected operational FCFF CAGR |
| **CLEAN** | `DCF 1st way` | `E20` | `0.05` | `explicit_assumption` | Macroeconomic Terminal Growth Constraint | N/A | Terminal | Growth Rate | Long-term sustainable terminal growth rate capped below nominal GDP growth |
| **CLEAN** | `DCF 1st way` | `E44` | `1041.4` | `website_sourced_input` | National Stock Exchange of India (NSE) | https://www.nseindia.com | 2025-10-15 | ₹/share | 15-Oct-2025 Official Close |
| **CLEAN** | `DCF 1st way` | `F11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **CLEAN** | `DCF 1st way` | `G11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **CLEAN** | `DCF 1st way` | `H11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **CLEAN** | `DCF 1st way` | `I11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **CLEAN** | `DCF 1st way` | `J11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **CLEAN** | `DCF 1st way` | `K11` | `0.45` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **CLEAN** | `DCF 1st way` | `E42` | `==G60` | `formula_derived` | Single Source-of-Truth Primary Valuation Output | N/A | 2025-10-15 | ₹/share | Direct link to verified primary valuation schedule |
| **EDELWEISS** | `Data Sheet` | `B7` | `1.0` | `website_sourced_input` | Screener.in / BSE Statutory Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | As of 15-Oct-2025 | ₹/share | Statutory Face Value |
| **EDELWEISS** | `Data Sheet` | `B8` | `108.85` | `website_sourced_input` | National Stock Exchange of India (NSE) / Yahoo Finance | https://www.nseindia.com | 2025-10-15 | ₹/share | Official Unadjusted Close |
| **EDELWEISS** | `Data Sheet` | `B9` | `10286.33` | `website_sourced_input` | Screener.in / NSE | https://www.screener.in/company/EDELWEISS/consolidated/ | 2025-10-15 | ₹ Cr | Market Capitalization |
| **EDELWEISS** | `Data Sheet` | `B6` | `=IF(B9>0, B9/B8, 0)` | `formula_derived` | Template Master Formula | N/A | 2025-10-15 | Cr shares | Implied share count = Market Cap / CMP |
| **EDELWEISS** | `Data Sheet` | `B17` | `6604.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2017 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `C17` | `8878.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2018 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `D17` | `11078.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2019 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `E17` | `9513.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2020 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `F17` | `10775.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2021 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `G17` | `7151.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2022 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `H17` | `8500.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2023 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `I17` | `9523.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2024 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `J17` | `9393.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2025 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Data Sheet` | `K17` | `10802.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | Mar 2026 | ₹ Cr | Revenue from Operations |
| **EDELWEISS** | `Beta` | `K8` | `0.75` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on raw empirical beta |
| **EDELWEISS** | `Beta` | `K11` | `0.25` | `explicit_assumption` | Blume (1975) Adjustment Convention | N/A | 2-Year Horizon | Weight | Mean-reversion weight on market beta (1.0) |
| **EDELWEISS** | `WACC-DCF` | `C30` | `18004.0` | `website_sourced_input` | Screener.in / BSE Filings | https://www.screener.in/company/EDELWEISS/consolidated/ | FY25 | ₹ Cr | Total Debt |
| **EDELWEISS** | `WACC-DCF` | `C31` | `10286.33` | `website_sourced_input` | NSE / Screener.in | https://www.nseindia.com | 2025-10-15 | ₹ Cr | Market Capitalization |
| **EDELWEISS** | `WACC-DCF` | `E22` | `0.095` | `explicit_assumption` | Weighted Average Borrowing Cost from Statutory Notes | https://www.bseindia.com | FY26E | Rate | Pre-tax cost of debt |
| **EDELWEISS** | `WACC-DCF` | `E23` | `0.2517` | `explicit_assumption` | Indian Corporate Tax Law (Section 115BAA) | https://incometaxindia.gov.in | Statutory | Rate | Base 22% + 10% Surcharge + 4% Cess = 25.168% |
| **EDELWEISS** | `WACC-DCF` | `J22` | `0.0678` | `website_sourced_input` | Investing.com / Trading Economics | https://www.investing.com | 2025-10-15 | Rate | India 10Y Benchmark G-Sec Yield |
| **EDELWEISS** | `DCF 1st way` | `E19` | `0.1` | `explicit_assumption` | Research Forecast Driver | N/A | FY26E-FY30E | Growth Rate | Expected operational FCFF CAGR |
| **EDELWEISS** | `DCF 1st way` | `E20` | `0.05` | `explicit_assumption` | Macroeconomic Terminal Growth Constraint | N/A | Terminal | Growth Rate | Long-term sustainable terminal growth rate capped below nominal GDP growth |
| **EDELWEISS** | `DCF 1st way` | `E44` | `108.85` | `website_sourced_input` | National Stock Exchange of India (NSE) | https://www.nseindia.com | 2025-10-15 | ₹/share | 15-Oct-2025 Official Close |
| **EDELWEISS** | `DCF 1st way` | `F11` | `0.5` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **EDELWEISS** | `DCF 1st way` | `G11` | `0.5` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **EDELWEISS** | `DCF 1st way` | `H11` | `0.5` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **EDELWEISS** | `DCF 1st way` | `I11` | `0.5` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **EDELWEISS** | `DCF 1st way` | `J11` | `0.5` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **EDELWEISS** | `DCF 1st way` | `K11` | `0.5` | `explicit_assumption` | Operating Model Reinvestment Ratio | N/A | Forecast | Ratio | CapEx + Working Capital reinvestment rate |
| **EDELWEISS** | `DCF 1st way` | `E42` | `==G60` | `formula_derived` | Single Source-of-Truth Primary Valuation Output | N/A | 2025-10-15 | ₹/share | Direct link to verified primary valuation schedule |
