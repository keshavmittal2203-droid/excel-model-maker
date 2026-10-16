# Clean Science and Technology Limited (CLEAN) — Financial Valuation Package

- **Target Entity**: Clean Science and Technology Limited
- **BSE Code**: `543318` | **NSE Ticker**: `CLEAN` | **ISIN**: `INE227W01023`
- **Industry Vertical**: Specialty Chemicals
- **Reporting Standard**: Indian Accounting Standards (Ind AS) Consolidated
- **Statutory Auditor**: B S R & Co. LLP
- **Candidate Recalculated Workbook**: [`CLEAN-Recalculated.xlsx`](CLEAN-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Package Contents & Artifact Structure

| Artifact | Path | Purpose & Status |
|:---|:---|:---|
| **Recalculated Workbook** | `CLEAN-Recalculated.xlsx` | 100% evaluated via Microsoft Excel COM 16.0; zero formula errors; $\Delta = 0.00$ balance. |
| **Package Manifest** | `package.manifest.json` | Provenance manifest pinning SHA-256 hashes, COM engine details, and preview records. |
| **Source Registry** | `SOURCE_REGISTRY.md` | Primary statutory filing citations from audited annual reports and BSE/NSE records. |
| **Sheet PDF Previews** | `previews/*.pdf` | High-resolution landscape PDF exports of all 6 substantive worksheets. |
| **Canonical Request** | `clean-package.json` | Standardized request format conforming to build pipeline. |
| **Model Pack** | `clean.model-pack.json` | Declarative validation checks. Status: `proposal` (pending human approval). |

## 2. Financial Economics & Business Architecture
### Clean Science and Technology Limited
- **Specialty Chemical Manufacturing**: 100% catalytic, eco-friendly green chemistry manufacturing.
- **Verticals**: Performance Chemicals (MEHQ, BHA, AP, HALS series) ~70%; Pharma Intermediates (Guaiacol, DCC) ~18%; FMCG Chemicals (4-MAP, Anisole) ~12%.
- **Clean Fino-Chem (CFCL)**: Unit 4 greenfield expansion at Kurkumbh (34 acres) commercializing 10,500 MTPA HALS capacity.
- **Balance Sheet**: Zero net debt, ROCE of 20.7%, EBITDA margin ~37%–40%.
- **Valuation**: Manufacturing FCFF DCF reflecting HALS capacity utilization ramp-up and normalized CapEx.
- **Intrinsic Fair Value**: **₹343.93 per share** (vs Historical Snapshot Price ₹784.00 as of 15-Oct-2025 NSE Close).
- **Market Snapshot Nature**: Point-in-time historical snapshot dated October 15, 2025; explicitly NOT a current real-time valuation.

## 3. Fast Verification Commands
```powershell
# Recalculate and verify Clean Science and Technology Limited candidate workbook
python -c "from validate_workbook import scan_formula_errors_and_missing_caches; from modelmaker import Package; pkg = Package(open('companies/CLEAN/CLEAN-Recalculated.xlsx', 'rb').read()); errs, miss = scan_formula_errors_and_missing_caches(pkg); print('Errors:', len(errs), 'Missing:', len(miss))"
```