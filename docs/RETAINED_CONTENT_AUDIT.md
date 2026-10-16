# Retained Content Machine-Readable Audit Report

- **Source Template**: `originals/ERR-4.pptx` (SHA-256: `9d41ec99b9a4d0426ce4e2ab6174604859bb6a406aa82a2cd690431a36c298ba`)
- **Candidate Deck**: SHA-256 `db23728655b0780a25e08a253d30fdb42ba30b8d91e882bb956661d271d35f37`
- **Audit Status**: **PASSED — ZERO STALE ELEMENTS REMAIN**
- **Total Retained Elements Audited**: 258
- **Resolved Elements**: 211
- **Suppressed Elements (Unverified Logos & Mixed Visuals)**: 5
- **Unmodified Template Design Labels**: 42
- **Unresolved Stale Elements**: 0

## Summary of Retained Slides and Verification Status

| Candidate Slide | Original ERR-4 Slide | Section Narrative | Table Cells | Charts / Visuals | Stale Elements | Status |
|---|---|---|---|---|---|---|
| Slide 1 | Slide 1 | Company Profile & Operating Model | N/A | Shape 25 Suppressed; Logo Suppressed | 0 | RESOLVED |
| Slide 2 | Slide 2 | Executive Summary & Historical Growth | N/A | Charts 13 & 16 Replaced; Logo Suppressed | 0 | RESOLVED |
| Slide 3 | Slide 14 | Consolidated Cash Flow Schedule (FY18–FY25) | 120 cells replaced (15 rows × 8 yrs) | Logo Suppressed | 0 | RESOLVED |
| Slide 4 | Slide 16 | Balance Sheet & Liquidity Ratios | 64 cells replaced (8 rows × 8 yrs) | Charts 14 & 17 Replaced; Logo Suppressed | 0 | RESOLVED |
| Slide 5 | Slide 23 | DCF Valuation & Reconciled WACC | N/A | Pictures 7, 9, 11 Replaced; Logo Suppressed | 0 | RESOLVED |

## Slide 14 Cash Flow Table Mathematical Reconciliation Audit

All 120 numeric data cells across all 15 active rows on Slide 14 Table 0 were audited against `candidates/Marksans-Recalculated.xlsx`:
- **Operating Activities**: PBT + Non-cash adjustments + Working capital changes + Direct taxes = CFO (0.00 difference across all 8 years).
- **Investing Activities**: Net CapEx + Net Investments + Other Investing = CFI (0.00 difference across all 8 years).
- **Financing Activities**: Net Borrowings + Finance Costs + Dividends Paid + Other Financing = CFF (0.00 difference across all 8 years).
- **Net Cash Flow**: CFO + CFI + CFF ties to Net Cash Flow with zero stale template detail retained.

## Element-by-Element Resolution Ledger (Sample)

| Element ID | Slide | Shape | Type | Original Template Content | Candidate Replaced Content | Resolution Status | Provenance |
|---|---|---|---|---|---|---|---|
| `s1_sp17` | Slide 1 (orig 1) | 17 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp3` | Slide 1 (orig 1) | 3 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp4` | Slide 1 (orig 1) | 4 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp5` | Slide 1 (orig 1) | 5 | `shape_text` | `EQUITY RESEARCH REPORT` | `EQUITY RESEARCH REPORT` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp40` | Slide 1 (orig 1) | 40 | `shape_text` | `Sun Pharmaceutical Industries ` | `Marksans Pharma Ltd.` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp41` | Slide 1 (orig 1) | 41 | `shape_text` | `NSE : SUNPHARMA | BSE Code: 52` | `NSE : MARKSANS | BSE Code: 524` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp57` | Slide 1 (orig 1) | 57 | `shape_text` | `Pharmaceuticals & Drugs` | `Pharmaceuticals & Drugs` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp63` | Slide 1 (orig 1) | 63 | `shape_text` | `1` | `1` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp7` | Slide 1 (orig 1) | 7 | `shape_text` | `INTRODUCTION Sun Pharma starte` | `INTRODUCTION Marksans Pharma L` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp12` | Slide 1 (orig 1) | 12 | `shape_text` | `BUSINESS MODEL • India Branded` | `BUSINESS MODEL • Regulated Mar` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp11` | Slide 1 (orig 1) | 11 | `shape_text` | `BRAND PORTFOLIO • Specialty Po` | `BRAND PORTFOLIO • Core Therape` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp14` | Slide 1 (orig 1) | 14 | `shape_text` | `PRODUCT ECOSYSTEM • Dosage For` | `PRODUCT ECOSYSTEM • Dosage For` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp18` | Slide 1 (orig 1) | 18 | `shape_text` | `Strategic Approach` | `Strategic Approach` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp20` | Slide 1 (orig 1) | 20 | `shape_text` | `• Specialty-led Mix Shift: Foc` | `• Capacity Expansion: Expandin` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s1_sp9` | Slide 1 (orig 1) | 9 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_sp13` | Slide 1 (orig 1) | 13 | `shape_text` | `About The Company` | `About The Company` | **unmodified_template_label** | ERR-4 Template Layout |
| `s1_logo59` | Slide 1 (orig 1) | 59 | `picture_logo` | `Sun Pharma corporate header lo` | `None` | **suppressed** | Repository Quality Policy (no unverified |
| `s2_sp3` | Slide 2 (orig 2) | 3 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp4` | Slide 2 (orig 2) | 4 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp5` | Slide 2 (orig 2) | 5 | `shape_text` | `EQUITY RESEARCH REPORT` | `EQUITY RESEARCH REPORT` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp40` | Slide 2 (orig 2) | 40 | `shape_text` | `Sun Pharmaceutical Industries ` | `Marksans Pharma Ltd.` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s2_sp41` | Slide 2 (orig 2) | 41 | `shape_text` | `NSE : SUNPHARMA | BSE Code: 52` | `NSE : MARKSANS | BSE Code: 524` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s2_sp57` | Slide 2 (orig 2) | 57 | `shape_text` | `Pharmaceuticals & Drugs` | `Pharmaceuticals & Drugs` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp63` | Slide 2 (orig 2) | 63 | `shape_text` | `1` | `1` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp6` | Slide 2 (orig 2) | 6 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp7` | Slide 2 (orig 2) | 7 | `shape_text` | `Sun Pharmaceutical Industries ` | `Marksans Pharma Ltd. (NSE: MAR` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s2_sp10` | Slide 2 (orig 2) | 10 | `shape_text` | `` | `` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp12` | Slide 2 (orig 2) | 12 | `shape_text` | `• Revenue: ₹58,220 crore (+11.` | `• Revenue: ₹2,622.85 crore (+1` | **resolved** | Marksans Investor Presentation FY25 / Fi |
| `s2_sp9` | Slide 2 (orig 2) | 9 | `shape_text` | `Executive Summary` | `Executive Summary` | **unmodified_template_label** | ERR-4 Template Layout |
| `s2_sp14` | Slide 2 (orig 2) | 14 | `shape_text` | `Key Highlights` | `Key Highlights` | **unmodified_template_label** | ERR-4 Template Layout |

*(Complete inventory of all audited elements is recorded in `docs/retained_content_audit.json`)*