# Edelweiss Financial Services Limited (EDELWEISS) — Financial Valuation & Presentation Package

- **Target Entity**: Edelweiss Financial Services Limited
- **BSE Code**: `532922` | **NSE Ticker**: `EDELWEISS` | **ISIN**: `INE532F01054`
- **Industry Vertical**: Diversified Financial Services
- **Reporting Basis**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; Industrial FCFF is strictly prohibited because ₹18,004 Cr of FY25 borrowings fund NBFC/HFC/ARC assets).
- **Candidate Recalculated Workbook**: [`EDELWEISS-Recalculated.xlsx`](EDELWEISS-Recalculated.xlsx)
- **Executive Presentation**: [`../../candidates/run_native_edelweiss/EDELWEISS-Executive-Overview.pptx`](../../candidates/run_native_edelweiss/EDELWEISS-Executive-Overview.pptx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Snapshot Cutoff)

- **Methodology**: `Financial Services Sum-of-the-Parts (SOTP) Valuation` (`supported_sotp_fcff_withheld`)
- **Implied Fair Value Per Share**: **₹133.54** (vs. 15-Oct-2025 NSE Snapshot **₹108.00**, Implied `+23.6%`)
- **Prior vs. Current Reconciliation**: Edelweiss SOTP is anchored on audited FY25 vertical metrics on 'Segment Operations'!H6:H12 (EAAA PAT ₹230 Cr @ 22.0x = ₹5,060 Cr; EARCL Net Worth ₹2,800 Cr @ 1.20x = ₹3,360 Cr; Mutual Fund AAUM ₹1,41,000 Cr @ 4.0% = ₹5,640 Cr; Insurance EV/Book Proxy ₹1,500 Cr @ 1.00x = ₹1,500 Cr; Credit Allocated Book ₹3,200 Cr @ 0.80x = ₹2,560 Cr; less Corporate HoldCo Net Debt -₹5,500 Cr = ₹12,620 Cr / 94.50 Cr shares = ₹133.54/share). Replacing Mar 2026 look-ahead figures in FY26E with connected vertical revenue formulas updated FY26E revenue to ₹10,167 Cr while preserving the FY25-anchored SOTP at ₹133.54/share.

## 2. Reproducible Single-Command Build
```powershell
python build_package.py --request examples/edelweiss-package.json --output-dir candidates/run_native_edelweiss --allow-proposal
```