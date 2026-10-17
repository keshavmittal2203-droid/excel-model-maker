# LTIMindtree Limited (LTM) — Financial Valuation & Research Model Package

- **Target Entity**: LTIMindtree Limited
- **BSE Code**: `540005` | **NSE Ticker**: `LTIM` | **ISIN**: `INE214T01019`
- **Industry Vertical**: IT Services & Digital Transformation
- **Reporting Basis**: FY19–FY21 reflect Pre-Merger Standalone/Consolidated LTI (Equity Cap ₹17–18 Cr); FY22 reflects Restated Combined Amalgamation Comparative (Equity Cap ₹30 Cr, +111.1% merger step); FY23–FY25 reflect Audited Merged LTIMindtree (effective Nov 14, 2022).
- **Candidate Recalculated Workbook**: [`LTM-Recalculated.xlsx`](LTM-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Official NSE Snapshot Cutoff)

- **Methodology**: `Unlevered Free Cash Flow (FCFF) DCF & CAPM WACC` (`supported`)
- **Implied Fair Value Per Share**: **₹2,811.09** (vs. 15-Oct-2025 NSE Official Close **₹5,609.50**, Implied `-49.89%`)
- **Prior vs. Current Reconciliation**: Eliminated double-counted operational growth (where +0*$I$5 ignored headcount/utilization/billing) so USD Revenue ($M) = Headcount * Utilization * Realized Billing per Billed FTE / 1,000,000 and INR Revenue = USD Revenue * FX / 10. Replaced reverse-solved 0.950 beta (11.50% round WACC) with 5-peer Tier-1 IT Hamada Relevered Beta (0.9405 -> Ke = 11.95%, WACC = 11.87%, alongside Direct 2Y Weekly OLS Beta 1.2569) and removed the +₹2,500 Cr hardcoded cash bridge offset (Cash & Treasury = ₹3,588 Cr Cash + ₹9,845 Cr Investments = ₹13,433 Cr). Implied DCF value is ₹2,811.09/share vs verified 15-Oct-2025 NSE Official Close ₹5,609.50.

## 2. Reproducible Single-Command Build (`--excel-only` & Full Package)
```powershell
python build_package.py --request examples/ltm-package.json --output-dir candidates/run_excel_ltm --allow-proposal --excel-only
```