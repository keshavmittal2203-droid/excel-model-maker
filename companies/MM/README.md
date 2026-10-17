# Mahindra & Mahindra Limited (MM) — Financial Valuation & Research Model Package

- **Target Entity**: Mahindra & Mahindra Limited
- **BSE Code**: `500520` | **NSE Ticker**: `M&M` | **ISIN**: `INE146A01017`
- **Industry Vertical**: Automotive, Farm Equipment & Financial Services
- **Reporting Basis**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; SOTP uses Audited FY25 M&M Standalone Core Operating Profit ₹18,416 Cr [Auto ₹11,700 Cr + Farm ₹6,716 Cr] and Audited FY25 Standalone Net Cash +₹9,109 Cr to prevent MMFS NBFC debt double-counting).
- **Candidate Recalculated Workbook**: [`MM-Recalculated.xlsx`](MM-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Official NSE Snapshot Cutoff)

- **Methodology**: `Sum-of-the-Parts (SOTP) Valuation` (`supported`)
- **Implied Fair Value Per Share**: **₹2,742.97** (vs. 15-Oct-2025 NSE Official Close **₹3,497.20**, Implied `-21.57%`)
- **Prior vs. Current Reconciliation**: Eliminated the backwards-solved ₹242,400 Cr target preservation (15,150 × 16x vs 20,200 × 12x) and replaced the unsourced -₹12,000 Cr standalone net debt plug with the Audited FY25 M&M Standalone Balance Sheet Bridge (Standalone Borrowings ₹1,682 Cr less Standalone Cash & Treasury ₹10,791 Cr = -₹9,109 Cr Net Debt / +₹9,109 Cr Standalone Net Cash). Connected FY26E Standalone Auto Core EBITDA (₹13,667 Cr @ 14.0x = ₹1,91,338 Cr) and Farm Core EBITDA (₹7,403 Cr @ 13.0x = ₹96,239 Cr) to Volume × ASP drivers, updated MMFS (52.16% of ₹36,305 Cr @ ₹293.85 15-Oct-2025 close × 80% = ₹15,149.35 Cr) and TechM (25.48% of ₹1,42,836 Cr @ ₹1,459.00 15-Oct-2025 close × 80% = ₹29,115.69 Cr), yielding ₹3,40,951.04 Cr Equity Value (₹2,742.97/share).

## 2. Reproducible Single-Command & Batch Build (`--excel-only`)
```powershell
# Single-company Excel build & validation:
python build_package.py --request examples/mm-package.json --output-dir candidates/run_excel_mm --allow-proposal --excel-only

# Batch build across all four priority targets:
python build_package.py --batch --allow-proposal --excel-only
```