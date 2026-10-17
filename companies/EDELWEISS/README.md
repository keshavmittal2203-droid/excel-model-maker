# Edelweiss Financial Services Limited (EDELWEISS) — Financial Valuation Package

- **Target Entity**: Edelweiss Financial Services Limited
- **BSE Code**: `532922` | **NSE Ticker**: `EDELWEISS` | **ISIN**: `INE532F01054`
- **Industry Vertical**: Diversified Financial Services
- **Reporting Standard**: Indian Accounting Standards (Ind AS) Consolidated
- **Statutory Auditor**: S.R. Batliboi & Co. LLP
- **Candidate Recalculated Workbook**: [`EDELWEISS-Recalculated.xlsx`](EDELWEISS-Recalculated.xlsx)
- **Package Manifest**: [`package.manifest.json`](package.manifest.json)
- **Source Registry**: [`SOURCE_REGISTRY.md`](SOURCE_REGISTRY.md) & [`source_registry.json`](source_registry.json)

## 1. Package Contents & Artifact Structure

| Artifact | Path | Purpose & Status |
|:---|:---|:---|
| **Recalculated Workbook** | `EDELWEISS-Recalculated.xlsx` | 100% evaluated via Microsoft Excel COM 16.0; zero formula errors; $\Delta = 0.00$ balance. |
| **Package Manifest** | `package.manifest.json` | Provenance manifest pinning SHA-256 hashes, COM engine details, and preview records. |
| **Source Registry** | `SOURCE_REGISTRY.md` | Primary statutory filing citations from audited annual reports and BSE/NSE records. |
| **Sheet PDF Previews** | `previews/*.pdf` | High-resolution landscape PDF exports of all 6 substantive worksheets. |
| **Canonical Request** | `edelweiss-package.json` | Standardized request format conforming to build pipeline. |
| **Model Pack** | `edelweiss.model-pack.json` | Declarative validation checks. Status: `proposal` (pending human approval). |

## 2. Financial Economics & Business Architecture
### Edelweiss Financial Services Limited
- **Diversified Conglomerate**: Alternatives (EAAA), Asset Reconstruction (EARCL), Mutual Fund (EAML), Insurance, and Credit/Mortgages.
- **Key Segment Earnings**: Alternative Asset Management FY25 PAT ₹230 Cr (up from ₹175 Cr in FY24); Asset Reconstruction FY25 PAT ₹385 Cr (up from ₹355 Cr in FY24); Mutual Fund AAUM ₹1.41 Lakh Cr.
- **Valuation Methodology**: Sum-of-the-Parts (SOTP) using vertical-appropriate metrics: EAAA @ 22x P/E, EARCL @ 1.2x P/B, Mutual Fund @ 4% of AUM, Insurance @ 1.0x Net Worth proxy, Credit @ 0.8x P/B, less HoldCo Net Debt (-₹5,500 Cr).
- **Intrinsic Fair Value**: **₹133.54 per share** (vs CMP ₹135.00).
- **Methodological Notice**: Industrial FCFF DCF is strictly withheld and blocked as invalid for financial institutions.

## 3. Fast Verification Commands
```powershell
# Recalculate and verify Edelweiss Financial Services Limited candidate workbook
python -c "from validate_workbook import scan_formula_errors_and_missing_caches; from modelmaker import Package; pkg = Package(open('companies/EDELWEISS/EDELWEISS-Recalculated.xlsx', 'rb').read()); errs, miss = scan_formula_errors_and_missing_caches(pkg); print('Errors:', len(errs), 'Missing:', len(miss))"
```