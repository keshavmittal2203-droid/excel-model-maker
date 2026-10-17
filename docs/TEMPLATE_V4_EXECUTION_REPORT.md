# Template V4 Execution, Recalculation & Perturbation Verification Report

## 1. Executive Summary
All four 17-sheet equity research models have been rebuilt directly from `originals/Marksans Keshav.xlsx`,
fully recalculated using native Microsoft Excel COM 16.0 (`CalculateFullRebuild`), tested for dynamic perturbation/restoration,
verified for 0 formula errors across 68 sheets, and exported to preview PDFs using native master print settings.

## 2. Valuation Outputs & Single Source-of-Truth Cells

| Company | Ticker | Primary Valuation Method | Source-of-Truth Cell | Cover Cell Link | Implied Value (₹/share) | 15-Oct-2025 CMP (₹) | Formula Errors |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| Mahindra & Mahindra | `M&M.NS` | SOTP Valuation | `'DCF 1st way'!E42` | `'Model ->'!C9` | **₹2742.97** | ₹3,497.20 | 0 |
| LTIMindtree | `LTM.NS` | Full-Year Annual FCFF DCF | `'DCF 1st way'!E42` | `'Model ->'!C9` | **₹2811.09** | ₹5,609.50 | 0 |
| Clean Science | `CLEAN.NS` | Manufacturing FCFF DCF | `'DCF 1st way'!E42` | `'Model ->'!C9` | **₹480.89** | ₹1,041.40 | 0 |
| Edelweiss Financial Services | `EDELWEISS.NS` | Financial Services SOTP | `'DCF 1st way'!E42` | `'Model ->'!C9` | **₹130.53** | ₹116.30 | 0 |

## 3. Perturbation & Restoration Test Results

| Company | Tested Input Parameter | Input Cell | Baseline Value | Perturbed Value | Perturbed Output | Restored Output | Restoration Match |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| MM | Auto Core EBITDA Multiple | `'DCF 1st way'!F50` | 14.0 | 15.0 | ₹2852.92 | ₹2742.97 | **100% Exact Match** |
| LTM | Terminal Growth Rate | `'DCF 1st way'!E20` | 0.05 | 0.055 | ₹2963.14 | ₹2811.09 | **100% Exact Match** |
| CLEAN | Terminal Growth Rate | `'DCF 1st way'!E20` | 0.05 | 0.055 | ₹510.62 | ₹480.89 | **100% Exact Match** |
| EDELWEISS | Alternatives Target Multiple | `'DCF 1st way'!F50` | 18.7 | 20.0 | ₹133.69 | ₹130.53 | **100% Exact Match** |

## 4. Master Template Integrity
- Immutable Master Template: `originals/Marksans Keshav.xlsx`
- Master Template SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a` (Byte-for-byte verified identical before and after run)
- Native Print Settings: Unmodified, inheriting native master layout without forced scaling or landscape distortion.
- Release Gate: `ready: false`, `human_pack_approval: pending`. PR #3 remains unmerged.
