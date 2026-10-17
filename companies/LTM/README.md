# LTIMindtree Limited (LTM) — Financial Valuation & Presentation Package

- **Target Entity**: LTIMindtree Limited
- **BSE Code**: `540005` | **NSE Ticker**: `LTIM` | **ISIN**: `INE214T01019`
- **Industry Vertical**: IT Services & Digital Transformation
- **Reporting Basis**: FY19–FY21 reflect Pre-Merger Standalone/Consolidated LTI (Equity Cap ₹17–18 Cr); FY22 reflects Restated Combined Amalgamation Comparative (Equity Cap ₹30 Cr); FY23–FY25 reflect Audited Merged LTIMindtree (effective Nov 14, 2022).
- **Candidate Recalculated Workbook**: [`LTM-Recalculated.xlsx`](LTM-Recalculated.xlsx)
- **Executive Presentation**: [`../../candidates/run_native_ltm/LTM-Executive-Overview.pptx`](../../candidates/run_native_ltm/LTM-Executive-Overview.pptx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Snapshot Cutoff)

- **Methodology**: `Unlevered Free Cash Flow (FCFF) DCF & CAPM WACC` (`supported`)
- **Implied Fair Value Per Share**: **₹2,813.51** (vs. 15-Oct-2025 NSE Snapshot **₹5,568.00**, Implied `-49.5%`)
- **Prior vs. Current Reconciliation**: Removed look-ahead Mar 2026 actuals from FY26E and connected all 5 forecast years (FY26E–FY30E) to the USD Constant-Currency Revenue Growth (7.5%–9.0%) * Realized USD/INR Exchange Rate (₹85.50–₹87.50) and Utilization-linked margin schedule on 'IT Services Operations', plus an explicit CAPM WACC build-up (11.50%). Implied DCF value increased by +₹60.85/share (+2.21%) from ₹2,752.66 to ₹2,813.51.

## 2. Reproducible Single-Command Build
```powershell
python build_package.py --request examples/ltm-package.json --output-dir candidates/run_native_ltm --allow-proposal
```