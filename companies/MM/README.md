# Mahindra & Mahindra Limited (MM) — Financial Valuation & Presentation Package

- **Target Entity**: Mahindra & Mahindra Limited
- **BSE Code**: `500520` | **NSE Ticker**: `M&M` | **ISIN**: `INE146A01017`
- **Industry Vertical**: Automotive, Farm Equipment & Financial Services
- **Reporting Basis**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; SOTP uses Standalone Core Auto/Farm EBITDA to prevent MMFS NBFC debt double-counting).
- **Candidate Recalculated Workbook**: [`MM-Recalculated.xlsx`](MM-Recalculated.xlsx)
- **Executive Presentation**: [`../../candidates/run_native_mm/MM-Executive-Overview.pptx`](../../candidates/run_native_mm/MM-Executive-Overview.pptx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Snapshot Cutoff)

- **Methodology**: `Sum-of-the-Parts (SOTP) Valuation` (`supported`)
- **Implied Fair Value Per Share**: **₹2,929.98** (vs. 15-Oct-2025 NSE Snapshot **₹2,921.00**, Implied `+0.3%`)
- **Prior vs. Current Reconciliation**: M&M SOTP is anchored on audited FY25 Standalone Auto EBITDA (₹15,150 Cr @ 16.0x = ₹2,42,400 Cr) and Standalone Farm EBITDA (₹6,400 Cr @ 14.0x = ₹89,600 Cr) plus MMFS (52.2%) and TechM (25.5%) stakes at a 20% HoldCo discount less Standalone Net Debt (-₹12,000 Cr). Replacing look-ahead Mar 2026 revenue (₹1,98,639 Cr) with connected FY26E–FY30E segment formulas updated FY26E revenue to ₹1,76,874 Cr while keeping FY25-anchored SOTP at ₹2,929.98/share.

## 2. Reproducible Single-Command Build
```powershell
python build_package.py --request examples/mm-package.json --output-dir candidates/run_native_mm --allow-proposal
```