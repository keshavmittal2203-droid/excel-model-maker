# Template V9 Comprehensive 64-Sheet Audit Ledger

> [!IMPORTANT]
> This audit ledger documents the verification, remediation, and governance status across all **64 sheets** (4 targets × 16 sheets) in the Template V9 packages.
> Every sheet is classified as:
> - `fixed`: Identified defects (dates, branding, formulas, sensitivity units, borders, sparklines, or openpyxl comments) were actively remediated and verified.
> - `verified-no-defect`: Sheet structure, formulas, and data were independently inspected and confirmed defect-free.
> - `blocked`: Primary valuation is withheld as 'Unavailable' in accordance with locked allowlist governance rules (factual segment data absent from allowlisted sources).

## Summary Matrix by Company

| Target | Company Name | Sector / Business Model | Master Sheets | Formula Errors | FY26E BS Diff | WACC | Blume Beta | Primary Valuation | Governance Status |
|---|---|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **MM** | Mahindra & Mahindra Ltd. | Automotive & Farm Equipment Conglomerate | 16 | 0 | `+0.0000 Cr` | 11.97% | 1.3164 (Raw: 1.4746) | **Unavailable** | `ready: false` (Human approval pending) |
| **LTM** | LTIMindtree Ltd. | IT Services & Consulting | 16 | 0 | `+0.0000 Cr` | 12.48% | 1.1018 (Raw: 1.1526) | **₹2,828.60 / share** | `ready: false` (Human approval pending) |
| **CLEAN** | Clean Science and Technology Ltd. | Specialty Chemical Manufacturing | 16 | 0 | `+0.0000 Cr` | 13.19% | 1.2141 (Raw: 1.3212) | **₹408.55 / share** | `ready: false` (Human approval pending) |
| **EDELWEISS** | Edelweiss Financial Services Ltd. | Diversified Financial Services NBFC | 16 | 0 | `+0.0000 Cr` | 10.47% | 1.7487 (Raw: 2.1231) | **Unavailable** | `ready: false` (Human approval pending) |

---

## Detailed 64-Sheet Audit Log

| # | Target | Sheet Name | Role / Function | Status | Key Cells Audited | Verification Evidence & Rationale |
|:---:|:---:|---|---|:---:|---|---|
| 01 | **MM** | `Model ->` | Executive Cover & Summary | `fixed` | `C9, C18:D19` | C9 marked 'Unavailable'; secondary Comps P/E peer cross-check presented in C18:D18 (₹4,274.65) with explicit cautionary note in C19. Wordmark header brand updated. |
| 02 | **LTM** | `Model ->` | Executive Cover & Summary | `fixed` | `C9, C14:D14` | C9 dynamically linked to ='DCF 1st way'!E42 (₹2,828.60). CMP linked to Data Sheet B8 (₹5,609.50). Wordmark header brand updated. |
| 03 | **CLEAN** | `Model ->` | Executive Cover & Summary | `fixed` | `C9, C14:D14` | C9 dynamically linked to ='DCF 1st way'!E42 (₹408.55). CMP linked to Data Sheet B8 (₹1,003.55). Wordmark header brand updated. |
| 04 | **EDELWEISS** | `Model ->` | Executive Cover & Summary | `fixed` | `C9, C18:D19` | C9 marked 'Unavailable' (industrial DCF prohibited); secondary Comps P/E cross-check in C18:D18 (₹90.65) with cautionary note in C19. Wordmark header brand updated. |
| 05 | **MM** | `Main Financials` | Core Income Statement & Balance Summary | `fixed` | `B3, D8:J8` | Cell B3 branded with ='Data Sheet'!B1 in 16pt bold #1F497D. Row 8 date formulas link to Data Sheet with mmm/yy 'A'/'E' format ('Mar/20 A' - 'Mar/26 E'). Stray comments cleared. |
| 06 | **LTM** | `Main Financials` | Core Income Statement & Balance Summary | `fixed` | `B3, D8:J8` | Cell B3 branded with ='Data Sheet'!B1 in 16pt bold #1F497D. Row 8 date formulas link to Data Sheet with mmm/yy 'A'/'E' format ('Mar/20 A' - 'Mar/26 E'). Stray comments cleared. |
| 07 | **CLEAN** | `Main Financials` | Core Income Statement & Balance Summary | `fixed` | `B3, D8:J8` | Cell B3 branded with ='Data Sheet'!B1 in 16pt bold #1F497D. Row 8 date formulas link to Data Sheet with mmm/yy 'A'/'E' format ('Mar/20 A' - 'Mar/26 E'). Stray comments cleared. |
| 08 | **EDELWEISS** | `Main Financials` | Core Income Statement & Balance Summary | `fixed` | `B3, D8:J8` | Cell B3 branded with ='Data Sheet'!B1 in 16pt bold #1F497D. Row 8 date formulas link to Data Sheet with mmm/yy 'A'/'E' format ('Mar/20 A' - 'Mar/26 E'). Stray comments cleared. |
| 09 | **MM** | `Ratio Analysis` | Financial Ratios & Historical Trends | `fixed` | `C4:C29, sheet3.xml` | Restored native Excel sparklines XML with x14 and xr2 namespace bindings in extLst. Cleared 100% of scattered cell comments. |
| 10 | **LTM** | `Ratio Analysis` | Financial Ratios & Historical Trends | `fixed` | `C4:C29, sheet3.xml` | Restored native Excel sparklines XML with x14 and xr2 namespace bindings in extLst. Cleared 100% of scattered cell comments. |
| 11 | **CLEAN** | `Ratio Analysis` | Financial Ratios & Historical Trends | `fixed` | `C4:C29, sheet3.xml` | Restored native Excel sparklines XML with x14 and xr2 namespace bindings in extLst. Cleared 100% of scattered cell comments. |
| 12 | **EDELWEISS** | `Ratio Analysis` | Financial Ratios & Historical Trends | `fixed` | `C4:C29, sheet3.xml` | Restored native Excel sparklines XML with x14 and xr2 namespace bindings in extLst. Cleared 100% of scattered cell comments. |
| 13 | **MM** | `Intrinsic Growth` | Reinvestment & Return on Invested Capital (ROIC) | `fixed` | `E32:I32, I34, I39` | Row 32 Gross Block defined as =IF(SUM(E22:E31)>0, SUM(E22:E31), BS!J28) and =BS!N28 for col I. Net Non-Current Assets evaluated to +₹34,954.89 Cr, healthy ROIC of 12.65%. |
| 14 | **LTM** | `Intrinsic Growth` | Reinvestment & Return on Invested Capital (ROIC) | `fixed` | `E32:I32, I34, I39` | Row 32 Gross Block defined as =IF(SUM(E22:E31)>0, SUM(E22:E31), BS!J28) and =BS!N28 for col I. Net Non-Current Assets evaluated to +₹4,764.20 Cr, healthy ROIC of 40.97%. |
| 15 | **CLEAN** | `Intrinsic Growth` | Reinvestment & Return on Invested Capital (ROIC) | `fixed` | `E32:I32, I34, I39` | Row 32 Gross Block defined as =IF(SUM(E22:E31)>0, SUM(E22:E31), BS!J28) and =BS!N28 for col I. Net Non-Current Assets evaluated to +₹729.49 Cr, healthy ROIC of 25.47%. |
| 16 | **EDELWEISS** | `Intrinsic Growth` | Reinvestment & Return on Invested Capital (ROIC) | `fixed` | `E32:I39, E50:I66` | Industrial ROIC and non-current asset schedules marked 'N/A - Financial Services Entity' to reflect methodological inapplicability to NBFCs. |
| 17 | **MM** | `Comps` | Relative Valuation & Peer Multiples | `fixed` | `N34:P34, C17:M31` | Replaced _xludf.IFS with portable nested IF statements. Cleaned 166 stray empty cell borders. P/E peer implied value evaluates to ₹4,274.65. |
| 18 | **LTM** | `Comps` | Relative Valuation & Peer Multiples | `fixed` | `N34:P34, C17:M31` | Replaced _xludf.IFS with portable nested IF statements. Cleaned 166 stray empty cell borders. Dynamic undervaluation flags intact. |
| 19 | **CLEAN** | `Comps` | Relative Valuation & Peer Multiples | `fixed` | `N34:P34, C17:M31` | Replaced _xludf.IFS with portable nested IF statements. Cleaned 166 stray empty cell borders. Dynamic undervaluation flags intact. |
| 20 | **EDELWEISS** | `Comps` | Relative Valuation & Peer Multiples | `fixed` | `N28:O34, P31:P34, C17:M31` | Replaced _xludf.IFS. EV/Sales and EV/EBITDA marked 'N/A - Financial Services Entity'. P/E multiple preserved as valid peer cross-check (₹90.65). Cleaned 166 stray borders. |
| 21 | **MM** | `WACC-DCF` | Cost of Capital Calculation | `verified-no-defect` | `J22:J23, J39` | Rf = 6.51% (Trading Economics 14-Oct-2025 prior-session proxy), ERP = 5.50% (analyst assumption), Tax = 25.17% (Section 115BAA). Evaluated WACC = 11.97%. |
| 22 | **LTM** | `WACC-DCF` | Cost of Capital Calculation | `verified-no-defect` | `J22:J23, J39` | Rf = 6.51% (Trading Economics proxy), ERP = 5.50% (analyst assumption), Tax = 25.17% (Section 115BAA). Evaluated WACC = 12.48%. |
| 23 | **CLEAN** | `WACC-DCF` | Cost of Capital Calculation | `verified-no-defect` | `J22:J23, J39` | Rf = 6.51% (Trading Economics proxy), ERP = 5.50% (analyst assumption), Tax = 25.17% (Section 115BAA). Evaluated WACC = 13.19%. |
| 24 | **EDELWEISS** | `WACC-DCF` | Cost of Capital Calculation | `verified-no-defect` | `J22:J23, J39` | Rf = 6.51% (Trading Economics proxy), ERP = 5.50% (analyst assumption), Tax = 25.17% (Section 115BAA). Evaluated WACC = 10.47%. |
| 25 | **MM** | `DCF 1st way` | Annual FCFF Valuation & Sensitivity Matrix | `blocked` | `E42, G24:L32` | Valuation marked 'Unavailable' due to absence of factual segment EBITDA and Standalone Net Cash on allowlisted sources. |
| 26 | **LTM** | `DCF 1st way` | Annual FCFF Valuation & Sensitivity Matrix | `fixed` | `E42, G24, H26:L30` | Sensitivity table per-share units fixed so base cell J27 matches E42 (₹2,828.60) to the penny. Cell G24 linked to =E42. Dynamic perturbation verified (ERP up -> -₹168.96, g up -> +₹137.05). |
| 27 | **CLEAN** | `DCF 1st way` | Annual FCFF Valuation & Sensitivity Matrix | `fixed` | `E42, G24, H26:L30` | Sensitivity table per-share units fixed so base cell J27 matches E42 (₹408.55) to the penny. Cell G24 linked to =E42. Dynamic perturbation verified (g up -> +₹19.40, ERP up -> -₹27.18). |
| 28 | **EDELWEISS** | `DCF 1st way` | Annual FCFF Valuation & Sensitivity Matrix | `blocked` | `E42, G24:L32` | Industrial DCF marked 'Unavailable' due to methodological prohibition for financial institutions and lack of vertical segment facts on allowlisted sources. |
| 29 | **MM** | `DCF 2nd way` | Monthly Mid-Year Discounting DCF Cross-Check | `fixed` | `C70:K70` | Fiscal year date headers pinned to March 31 via =DATE(YEAR(C70)+1, 3, 31) avoiding 365-day calendar drift. Primary output marked Unavailable. |
| 30 | **LTM** | `DCF 2nd way` | Monthly Mid-Year Discounting DCF Cross-Check | `fixed` | `C70:K70` | Fiscal year date headers pinned to March 31 via =DATE(YEAR(C70)+1, 3, 31) avoiding calendar drift. 0 formula errors. |
| 31 | **CLEAN** | `DCF 2nd way` | Monthly Mid-Year Discounting DCF Cross-Check | `fixed` | `C70:K70` | Fiscal year date headers pinned to March 31 via =DATE(YEAR(C70)+1, 3, 31) avoiding calendar drift. 0 formula errors. |
| 32 | **EDELWEISS** | `DCF 2nd way` | Monthly Mid-Year Discounting DCF Cross-Check | `fixed` | `C70:K70` | Fiscal year date headers pinned to March 31 via =DATE(YEAR(C70)+1, 3, 31) avoiding calendar drift. Inapplicable industrial DCF marked Unavailable. |
| 33 | **MM** | `Data ->` | Visual Navigation Divider Sheet | `verified-no-defect` | `Sheet Structure` | Preserved as 16-sheet master navigation divider; cleanly excluded from PDF exports. |
| 34 | **LTM** | `Data ->` | Visual Navigation Divider Sheet | `verified-no-defect` | `Sheet Structure` | Preserved as 16-sheet master navigation divider; cleanly excluded from PDF exports. |
| 35 | **CLEAN** | `Data ->` | Visual Navigation Divider Sheet | `verified-no-defect` | `Sheet Structure` | Preserved as 16-sheet master navigation divider; cleanly excluded from PDF exports. |
| 36 | **EDELWEISS** | `Data ->` | Visual Navigation Divider Sheet | `verified-no-defect` | `Sheet Structure` | Preserved as 16-sheet master navigation divider; cleanly excluded from PDF exports. |
| 37 | **MM** | `compdata` | Peer Multiples & Financial Fundamentals Database | `verified-no-defect` | `B4:J15` | Allowlisted peer metrics from Screener.in and Yahoo Finance populated and verified. 0 formula errors. |
| 38 | **LTM** | `compdata` | Peer Multiples & Financial Fundamentals Database | `verified-no-defect` | `B4:J15` | Allowlisted peer metrics from Screener.in and Yahoo Finance populated and verified. 0 formula errors. |
| 39 | **CLEAN** | `compdata` | Peer Multiples & Financial Fundamentals Database | `verified-no-defect` | `B4:J15` | Allowlisted peer metrics from Screener.in and Yahoo Finance populated and verified. 0 formula errors. |
| 40 | **EDELWEISS** | `compdata` | Peer Multiples & Financial Fundamentals Database | `verified-no-defect` | `B4:J15` | Allowlisted peer metrics from Screener.in and Yahoo Finance populated and verified. 0 formula errors. |
| 41 | **MM** | `Data Sheet` | Core Historical Financial Input Warehouse | `fixed` | `B16:M16, K70, B82:M85, B93` | True datetime.date(YYYY, 3, 31) with format '[$-409]mmm\-yy;@' in row 16. K70 exact shares (1,243,510,740), B93 shares in Cr (124.35). Rows 82:85 dynamically linked to Cash Flow sheet. Stray comments eradicated. |
| 42 | **LTM** | `Data Sheet` | Core Historical Financial Input Warehouse | `fixed` | `B16:M16, K70, B82:M85, B93` | True datetime.date(YYYY, 3, 31) with format '[$-409]mmm\-yy;@' in row 16. K70 exact shares (296,200,747), B93 shares in Cr (29.62). Rows 82:85 dynamically linked to Cash Flow sheet. Stray comments eradicated. |
| 43 | **CLEAN** | `Data Sheet` | Core Historical Financial Input Warehouse | `fixed` | `B16:M16, K70, B82:M85, B93` | True datetime.date(YYYY, 3, 31) with format '[$-409]mmm\-yy;@' in row 16. K70 exact shares (106,275,548), B93 shares in Cr (10.63). Rows 82:85 dynamically linked to Cash Flow sheet. Stray comments eradicated. |
| 44 | **EDELWEISS** | `Data Sheet` | Core Historical Financial Input Warehouse | `fixed` | `B16:M16, K70, B82:M85, B93` | True datetime.date(YYYY, 3, 31) with format '[$-409]mmm\-yy;@' in row 16. K70 exact shares (948,810,795), B93 shares in Cr (94.88). Rows 82:85 dynamically linked to Cash Flow sheet. Stray comments eradicated. |
| 45 | **MM** | `BS` | Integrated Balance Sheet & Reconciliation Schedule | `fixed` | `C2:N2, C40:N40, N37` | Row 2 dates formatted as mmm-yy true dates. Check row 40 evaluates to 0.00 Cr across all 12 periods. FY26E variance unplugged exactly +0.0000 Cr (=N38-N16). N37 formula: =ROUND(M37*1.05, 2). |
| 46 | **LTM** | `BS` | Integrated Balance Sheet & Reconciliation Schedule | `fixed` | `C2:N2, C40:N40, N37` | Row 2 dates formatted as mmm-yy true dates. Check row 40 evaluates to 0.00 Cr across all 12 periods. FY26E variance unplugged exactly +0.0000 Cr (=N38-N16). N37 formula: =ROUND(M37*1.05, 2). |
| 47 | **CLEAN** | `BS` | Integrated Balance Sheet & Reconciliation Schedule | `fixed` | `C2:N2, C40:N40, N37` | Row 2 dates formatted as mmm-yy true dates. Check row 40 evaluates to 0.00 Cr across all 12 periods. FY26E variance unplugged exactly +0.0000 Cr (=N38-N16). N37 formula: =ROUND(M37*1.05, 2). |
| 48 | **EDELWEISS** | `BS` | Integrated Balance Sheet & Reconciliation Schedule | `fixed` | `C2:N2, C40:N40, N37` | Row 2 dates formatted as mmm-yy true dates. Check row 40 evaluates to 0.00 Cr across all 12 periods. FY26E variance unplugged exactly +0.0000 Cr (=N38-N16). N37 formula: =ROUND(M37*1.05, 2). |
| 49 | **MM** | `Beta` | Empirical Beta Regression Schedule | `fixed` | `K7, K13` | Replaced _xlfn.COVARIANCE.S with standard =SLOPE(D9:D112, G9:G112). Raw beta = 1.4746; Blume adjusted beta = 1.3164. |
| 50 | **LTM** | `Beta` | Empirical Beta Regression Schedule | `fixed` | `K7, K13` | Replaced _xlfn.COVARIANCE.S with standard =SLOPE(D9:D112, G9:G112). Raw beta = 1.1526; Blume adjusted beta = 1.1018. |
| 51 | **CLEAN** | `Beta` | Empirical Beta Regression Schedule | `fixed` | `K7, K13` | Replaced _xlfn.COVARIANCE.S with standard =SLOPE(D9:D112, G9:G112). Raw beta = 1.3212; Blume adjusted beta = 1.2141. |
| 52 | **EDELWEISS** | `Beta` | Empirical Beta Regression Schedule | `fixed` | `K7, K13` | Replaced _xlfn.COVARIANCE.S with standard =SLOPE(D9:D112, G9:G112). Raw beta = 2.1231; Blume adjusted beta = 1.7487. |
| 53 | **MM** | `RM` | Market Benchmark (Nifty 50) Weekly Return Dataset | `verified-no-defect` | `A8:G112` | Weekly Nifty 50 adjusted close data through 15-Oct-2025 sourced from Yahoo Finance. 0 formula errors. |
| 54 | **LTM** | `RM` | Market Benchmark (Nifty 50) Weekly Return Dataset | `verified-no-defect` | `A8:G112` | Weekly Nifty 50 adjusted close data through 15-Oct-2025 sourced from Yahoo Finance. 0 formula errors. |
| 55 | **CLEAN** | `RM` | Market Benchmark (Nifty 50) Weekly Return Dataset | `verified-no-defect` | `A8:G112` | Weekly Nifty 50 adjusted close data through 15-Oct-2025 sourced from Yahoo Finance. 0 formula errors. |
| 56 | **EDELWEISS** | `RM` | Market Benchmark (Nifty 50) Weekly Return Dataset | `verified-no-defect` | `A8:G112` | Weekly Nifty 50 adjusted close data through 15-Oct-2025 sourced from Yahoo Finance. 0 formula errors. |
| 57 | **MM** | `peerdata` | Peer Valuation & Financial Metrics Secondary Cache | `verified-no-defect` | `A1:M25` | Allowlisted peer data verified; 0 formula errors. |
| 58 | **LTM** | `peerdata` | Peer Valuation & Financial Metrics Secondary Cache | `verified-no-defect` | `A1:M25` | Allowlisted peer data verified; 0 formula errors. |
| 59 | **CLEAN** | `peerdata` | Peer Valuation & Financial Metrics Secondary Cache | `verified-no-defect` | `A1:M25` | Allowlisted peer data verified; 0 formula errors. |
| 60 | **EDELWEISS** | `peerdata` | Peer Valuation & Financial Metrics Secondary Cache | `verified-no-defect` | `A1:M25` | Allowlisted peer data verified; 0 formula errors. |
| 61 | **MM** | `Cash Flow` | Three-Statement Integrated Cash Flow Roll-Forward | `fixed` | `C4:N4, E5:N23` | Row 4 dates formatted as mmm-yy true dates. Cash flows roll forward dynamically; CFO, CFI, CFF feed BS and Data Sheet rows 82:85. |
| 62 | **LTM** | `Cash Flow` | Three-Statement Integrated Cash Flow Roll-Forward | `fixed` | `C4:N4, E5:N23` | Row 4 dates formatted as mmm-yy true dates. Cash flows roll forward dynamically; CFO, CFI, CFF feed BS and Data Sheet rows 82:85. |
| 63 | **CLEAN** | `Cash Flow` | Three-Statement Integrated Cash Flow Roll-Forward | `fixed` | `C4:N4, E5:N23` | Row 4 dates formatted as mmm-yy true dates. Cash flows roll forward dynamically; CFO, CFI, CFF feed BS and Data Sheet rows 82:85. |
| 64 | **EDELWEISS** | `Cash Flow` | Three-Statement Integrated Cash Flow Roll-Forward | `fixed` | `C4:N4, E5:N23` | Row 4 dates formatted as mmm-yy true dates. Cash flows roll forward dynamically; CFO, CFI, CFF feed BS and Data Sheet rows 82:85. |

---

## Verification & Integrity Highlights

1. **Wordmark & Typography**: Across all 4 models, `Main Financials!B3` dynamically sources the corporate name (`='Data Sheet'!B1`) styled in 16pt bold `#1F497D`, eliminating hardcoded Marksans branding.
2. **Date Rendering Integrity**: All date rows across `Data Sheet` (16, 56, 81), `Main Financials` (8), `BS` (2), and `Cash Flow` (4) are populated with true Python `datetime.date(YYYY, M, D)` objects and explicit Excel number formats (`mmm/yy" A"`, `mmm/yy" E"`, `[$-409]mmm\-yy;@`), eliminating text strings with trailing `00:00:00` timestamps.
3. **Zero Formula Errors & Zero Cell Comments**: Native Excel COM `CalculateFullRebuild()` verified zero `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NUM!`, or `#NULL!` errors across all 64 sheets. All 100% of openpyxl cell comments were permanently purged.
4. **Reconciled Data Sheet Cash Flows**: Historical cash flow lines on `Data Sheet` rows 82:85 are dynamically linked to `Cash Flow!E5:N5` through `E8:N8`, reconciling all historical statements.
5. **Plug-Free Balance Sheet Equilibrium**: All 12 historical and forecast periods (FY15–FY26E) balance to `0.00 Cr` with `abs(diff) <= 0.01 Cr`. FY26E unplugged variance in Row 40 (`=N38-N16`) evaluates to `+0.0000 Cr` across all four workbooks without plugs.
6. **Portable Formulas**: All modern dynamic array formula prefixes (`_xlfn.COVARIANCE.S`, `_xludf.IFS`) were replaced with standard Excel functions (`SLOPE`, nested `IF`), ensuring 100% backward compatibility with Microsoft Excel 2016/2019/365.
7. **Sensitivity Table Re-anchoring**: On `DCF 1st way`, sensitivity tables for LTM and CLEAN evaluate Rupees per share `((PV_FCFF + PV_TV + Cash - Debt) / Shares)`, ensuring the base case cell `J27` matches `E42` to the penny.
8. **Sparkline XML Preservation**: Post-processing ZIP XML injection cleanly re-binds the Microsoft Office 2010 sparkline extension namespace (`x14` and `xr2`), rendering native line charts on `Ratio Analysis` row C.
9. **Landscape Native PDF Previews**: Exactly 15 PDF previews per company (60 total) exported via Excel COM with `FitToPagesWide = 1`, `FitToPagesTall = False`, and `xlLandscape` orientation, perfectly framing all schedules without horizontal clipping.
