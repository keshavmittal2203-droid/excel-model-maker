# LTIMindtree Limited (LTM) — Financial Valuation Package

- **Target Entity**: LTIMindtree Limited
- **BSE Code**: `540005` | **NSE Ticker**: `LTIM` | **ISIN**: `INE214T01019`
- **Industry Vertical**: IT Services & Consulting
- **Reporting Standard**: Indian Accounting Standards (Ind AS) Consolidated
- **Statutory Auditor**: B S R & Co. LLP
- **Candidate Recalculated Workbook**: [`LTM-Recalculated.xlsx`](LTM-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Package Contents & Artifact Structure

| Artifact | Path | Purpose & Status |
|:---|:---|:---|
| **Recalculated Workbook** | `LTM-Recalculated.xlsx` | 100% evaluated via Microsoft Excel COM 16.0; zero formula errors; $\Delta = 0.00$ balance. |
| **Package Manifest** | `package.manifest.json` | Provenance manifest pinning SHA-256 hashes, COM engine details, and preview records. |
| **Source Registry** | `SOURCE_REGISTRY.md` | Primary statutory filing citations from audited annual reports and BSE/NSE records. |
| **Sheet PDF Previews** | `previews/*.pdf` | High-resolution landscape PDF exports of all 6 substantive worksheets. |
| **Canonical Request** | `ltm-package.json` | Standardized request format conforming to build pipeline. |
| **Model Pack** | `ltm.model-pack.json` | Declarative validation checks. Status: `proposal` (pending human approval). |

## 2. Financial Economics & Business Architecture
### LTIMindtree Limited
- **Merger Context**: Formed via amalgamation of Larsen & Toubro Infotech (LTI) and Mindtree Limited (effective Nov 14, 2022). Unified reporting commenced in FY23.
- **USD Scale & Constant Currency**: FY25 USD revenue stood at $4,492.5M (+5.0% constant currency growth; +7.0% INR to ₹38,008 Cr).
- **Cost & Delivery Structure**: Software professionals headcount: 84,307; utilization: 86.5%; operating EBIT margin: 14.5% (FY25 EBIT: ₹5,503 Cr).
- **Valuation**: Unlevered Free Cash Flow (FCFF) DCF discounting at 11.5% WACC and 5.0% Terminal Growth Rate.
- **Intrinsic Fair Value**: **₹2,752.66 per share** (vs Historical Snapshot Price ₹4,390.00 as of 15-Oct-2025 NSE Close).
- **Market Snapshot Nature**: Point-in-time historical snapshot dated October 15, 2025; explicitly NOT a current real-time valuation.

## 3. Fast Verification Commands
```powershell
# Recalculate and verify LTIMindtree Limited candidate workbook
python -c "from validate_workbook import scan_formula_errors_and_missing_caches; from modelmaker import Package; pkg = Package(open('companies/LTM/LTM-Recalculated.xlsx', 'rb').read()); errs, miss = scan_formula_errors_and_missing_caches(pkg); print('Errors:', len(errs), 'Missing:', len(miss))"
```