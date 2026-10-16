# Mahindra & Mahindra Limited (MM) — Financial Valuation Package

- **Target Entity**: Mahindra & Mahindra Limited
- **BSE Code**: `500520` | **NSE Ticker**: `M&M` | **ISIN**: `INE146A01017`
- **Industry Vertical**: Automotive & Farm Equipment
- **Reporting Standard**: Indian Accounting Standards (Ind AS) Consolidated
- **Statutory Auditor**: B S R & Co. LLP
- **Candidate Recalculated Workbook**: [`MM-Recalculated.xlsx`](MM-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Package Contents & Artifact Structure

| Artifact | Path | Purpose & Status |
|:---|:---|:---|
| **Recalculated Workbook** | `MM-Recalculated.xlsx` | 100% evaluated via Microsoft Excel COM 16.0; zero formula errors; $\Delta = 0.00$ balance. |
| **Package Manifest** | `package.manifest.json` | Provenance manifest pinning SHA-256 hashes, COM engine details, and preview records. |
| **Source Registry** | `SOURCE_REGISTRY.md` | Primary statutory filing citations from audited annual reports and BSE/NSE records. |
| **Sheet PDF Previews** | `previews/*.pdf` | High-resolution landscape PDF exports of all 6 substantive worksheets. |
| **Canonical Request** | `mm-package.json` | Standardized request format conforming to build pipeline. |
| **Model Pack** | `mm.model-pack.json` | Declarative validation checks. Status: `proposal` (pending human approval). |

## 2. Financial Economics & Business Architecture
### Mahindra & Mahindra Limited
- **Automotive Sector**: Market leader in SUVs (22.5% revenue market share) and Light Commercial Vehicles (<3.5T). FY25 revenue: ₹90,825 Cr (+19% YoY); volumes: 941,115 units.
- **Farm Equipment Sector**: Global market leader in tractors. FY25 revenue: ₹35,375 Cr (+6% YoY); volumes: 424,641 units.
- **Financial Services**: 52.2% stake in Mahindra & Mahindra Financial Services Limited (MMFS).
- **Valuation**: Sum-of-the-Parts (SOTP) approach valuing Automotive Core (12x EV/EBITDA), Farm Equipment Core (14x EV/EBITDA), MMFS stake (20% HoldCo discount), and Tech Mahindra (25.5% stake, 20% discount).
- **Intrinsic Fair Value**: **₹2,929.98 per share** (vs Historical Snapshot Price ₹2,921.00 as of 15-Oct-2025 NSE Close).
- **Market Snapshot Nature**: Point-in-time historical snapshot dated October 15, 2025; explicitly NOT a current real-time valuation.

## 3. Fast Verification Commands
```powershell
# Recalculate and verify Mahindra & Mahindra Limited candidate workbook
python -c "from validate_workbook import scan_formula_errors_and_missing_caches; from modelmaker import Package; pkg = Package(open('companies/MM/MM-Recalculated.xlsx', 'rb').read()); errs, miss = scan_formula_errors_and_missing_caches(pkg); print('Errors:', len(errs), 'Missing:', len(miss))"
```