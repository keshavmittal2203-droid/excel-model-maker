# Edelweiss Financial Services Limited (EDELWEISS) — Financial Valuation & Research Model Package

- **Target Entity**: Edelweiss Financial Services Limited
- **BSE Code**: `532922` | **NSE Ticker**: `EDELWEISS` | **ISIN**: `INE532F01054`
- **Industry Vertical**: Diversified Financial Services
- **Reporting Basis**: Ind AS Consolidated (FY19–FY25 Audited, March 31 fiscal close; Industrial FCFF is strictly prohibited because ₹18,004 Cr of FY25 consolidated borrowings include ₹14,932 Cr of ring-fenced NBFC/HFC/ARC operating subsidiary debt; SOTP deducts Audited FY25 EFSL Standalone HoldCo Net Debt of ₹3,067 Cr).
- **Candidate Recalculated Workbook**: [`EDELWEISS-Recalculated.xlsx`](EDELWEISS-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Valuation & Financial Summary (15-Oct-2025 Official NSE Snapshot Cutoff)

- **Methodology**: `Financial Services Sum-of-the-Parts (SOTP) Valuation` (`supported_sotp_fcff_withheld`)
- **Implied Fair Value Per Share**: **₹130.53** (vs. 15-Oct-2025 NSE Official Close **₹108.85**, Implied `+19.92%`)
- **Prior vs. Current Reconciliation**: Built explicit vertical-specific financial services drivers on 'Segment Operations' (EAAA Earning AUM × Fee & Carry Yield; EARCL Recoveries × Fee Rate; Mutual Fund Equity+Alt AUM × Fee Rate; Life & General Insurance GWP & Actuarial Indian Embedded Value [IEV]; Credit Book × Net Interest & Fee Yield), disclosed EFSL attributable ownership stakes (EAAA 100%, EARCL 79.8% net of CDPQ 20.2% NCI, EAML 85% post-WestBridge 15%, Life 75% [Tokio Marine 25%] + General 100%, Credit 100%) with a 15% Conglomerate Discount (Gross Vertical Sum ₹15,402 Cr), and replaced the unsourced -₹5,500 Cr HoldCo debt assumption with the Audited FY25 EFSL Standalone Balance Sheet Bridge (Standalone Borrowings ₹3,072 Cr less Standalone Cash ₹5 Cr = -₹3,067 Cr Net Debt), yielding ₹12,335 Cr Equity Value (₹130.53/share) vs verified 15-Oct-2025 NSE Official Close ₹108.85.

## 2. Reproducible Single-Command & Batch Build (`--excel-only`)
```powershell
# Single-company Excel build & validation:
python build_package.py --request examples/edelweiss-package.json --output-dir candidates/run_excel_edelweiss --allow-proposal --excel-only

# Batch build across all four priority targets:
python build_package.py --batch --allow-proposal --excel-only
```