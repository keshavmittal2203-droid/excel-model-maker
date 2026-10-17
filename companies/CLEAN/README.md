# Clean Science and Technology Limited (CLEAN) — Financial Valuation & Presentation Package

- **Target Entity**: Clean Science and Technology Limited
- **BSE Code**: `543318` | **NSE Ticker**: `CLEAN` | **ISIN**: `INE227W01023`
- **Industry Vertical**: Specialty Chemicals Manufacturing
- **Reporting Basis**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; Equity Share Capital increased from ₹1.06 Cr in FY20 to ₹10.62 Cr in FY21 via pre-IPO bonus issue).
- **Candidate Recalculated Workbook**: [`CLEAN-Recalculated.xlsx`](CLEAN-Recalculated.xlsx)
- **Executive Presentation**: [`../../candidates/run_native_clean/CLEAN-Executive-Overview.pptx`](../../candidates/run_native_clean/CLEAN-Executive-Overview.pptx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Snapshot Cutoff)

- **Methodology**: `Manufacturing FCFF DCF & CAPM WACC` (`supported`)
- **Implied Fair Value Per Share**: **₹415.69** (vs. 15-Oct-2025 NSE Snapshot **₹1,039.00**, Implied `-60.0%`)
- **Prior vs. Current Reconciliation**: Removed look-ahead Mar 2026 actuals from FY26E and explicitly connected Parent Kurkumbh Units 1–3 (29,900 MTPA) and Wholly-Owned Subsidiary Clean Fino-Chem Ltd (CFCL Unit 4 HALS capacity expanding from 10,500 MTPA at 55% utilization in FY25 to 15,000 MTPA at 90% utilization by FY30E) to consolidated revenue without double counting. Implied DCF value increased by +₹71.76/share (+20.86%) from ₹343.93 to ₹415.69.

## 2. Reproducible Single-Command Build
```powershell
python build_package.py --request examples/clean-package.json --output-dir candidates/run_native_clean --allow-proposal
```