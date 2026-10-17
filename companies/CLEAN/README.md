# Clean Science and Technology Limited (CLEAN) — Financial Valuation & Research Model Package

- **Target Entity**: Clean Science and Technology Limited
- **BSE Code**: `543318` | **NSE Ticker**: `CLEAN` | **ISIN**: `INE227W01023`
- **Industry Vertical**: Specialty Chemicals Manufacturing
- **Reporting Basis**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; Equity Share Capital increased from ₹1.06 Cr in FY20 to ₹10.62 Cr in FY21 via pre-IPO bonus issue).
- **Candidate Recalculated Workbook**: [`CLEAN-Recalculated.xlsx`](CLEAN-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Official NSE Snapshot Cutoff)

- **Methodology**: `Manufacturing FCFF DCF & CAPM WACC` (`supported`)
- **Implied Fair Value Per Share**: **₹480.89** (vs. 15-Oct-2025 NSE Official Close **₹1,041.40**, Implied `-53.82%`)
- **Prior vs. Current Reconciliation**: Held CFCL Unit 4 HALS installed capacity at verified 10,500 MTPA across FY24–FY30E in Base Case (replacing unsourced 12,000–15,000 MTPA step-ups with commercial utilization ramp from 15% in FY24 and 21% in FY25 to 88% in FY30E), replaced reverse-solved 0.904 beta (12.00% round WACC) with empirical 2-Year Weekly OLS Beta vs Nifty 50 (0.9006 -> Ke = 11.73%, WACC = 11.73%), and removed the +₹45 Cr hardcoded cash bridge offset (Cash & Liquid Investments = ₹20 Cr Cash + ₹381 Cr Investments = ₹401 Cr). Implied DCF value is ₹480.89/share vs verified 15-Oct-2025 NSE Official Close ₹1,041.40.

## 2. Reproducible Single-Command Build (`--excel-only` & Full Package)
```powershell
python build_package.py --request examples/clean-package.json --output-dir candidates/run_excel_clean --allow-proposal --excel-only
```