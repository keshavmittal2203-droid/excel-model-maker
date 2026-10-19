# Template V5 Cell-Level Formula & Label Adaptation Log

> [!NOTE]
> This log records every intentional cell label, title, formula guard, single source-of-truth link, and SOTP/FCFF schedule adapted from `originals/Marksans Keshav.xlsx` in template_v5, eliminating formula literals and establishing direct cell references.

## Target Entity: MM (MAHINDRA & MAHINDRA LIMITED)

- **Primary Valuation Method**: Sum-of-the-Parts (SOTP) Valuation with Standalone Net Cash Ring-Fencing
- **Primary Fair Value per Share**: **₹2,742.97**
- **Single Source-of-Truth Cell**: `'DCF 1st way'!E42`
- **Executive Cover Link**: `'Model ->'!C9` links dynamically to `='DCF 1st way'!E42`

| Sheet | Cell | Original Template Value | Target Adapted Value | Nature of Change | Rationale / Explanation |
|---|---|---|---|---|---|
| `Data Sheet` | `B1` | `MARKSANS PHARMA LTD` | `MAHINDRA & MAHINDRA LIMITED` | label | Update company name |
| `DCF 1st way` | `E42` | `18.71` | `=G60` | reconciliation_formula | Link primary valuation output to SOTP schedule G60 |
| `Model ->` | `C9` | `None` | `='DCF 1st way'!E42` | executive_summary_link | Dynamically link cover primary value to source-of-truth cell 'DCF 1st way'!E42 |

## Target Entity: LTM (LTIMINDTREE LIMITED)

- **Primary Valuation Method**: Full-Year Annual Free Cash Flow to Firm (FCFF) DCF Baseline
- **Primary Fair Value per Share**: **₹2,811.09**
- **Single Source-of-Truth Cell**: `'DCF 1st way'!E42`
- **Executive Cover Link**: `'Model ->'!C9` links dynamically to `='DCF 1st way'!E42`

| Sheet | Cell | Original Template Value | Target Adapted Value | Nature of Change | Rationale / Explanation |
|---|---|---|---|---|---|
| `Data Sheet` | `B1` | `MARKSANS PHARMA LTD` | `LTIMINDTREE LIMITED` | label | Update company name |
| `WACC-DCF` | `E30` | `=H18` | `=D30` | capital_structure_target | Set target capital structure equal to company actual capital structure |
| `DCF 1st way` | `E42` | `1377.79` | `=E40/E41` | reconciliation_formula | Connect native DCF formula E40/E41 directly to forecast FCFF and empirical WACC |
| `Model ->` | `C9` | `None` | `='DCF 1st way'!E42` | executive_summary_link | Dynamically link cover primary value to source-of-truth cell 'DCF 1st way'!E42 |

## Target Entity: CLEAN (CLEAN SCIENCE AND TECHNOLOGY LIMITED)

- **Primary Valuation Method**: Full-Year Manufacturing FCFF DCF Baseline (CFCL Unit 4 Commercialization)
- **Primary Fair Value per Share**: **₹480.89**
- **Single Source-of-Truth Cell**: `'DCF 1st way'!E42`
- **Executive Cover Link**: `'Model ->'!C9` links dynamically to `='DCF 1st way'!E42`

| Sheet | Cell | Original Template Value | Target Adapted Value | Nature of Change | Rationale / Explanation |
|---|---|---|---|---|---|
| `Data Sheet` | `B1` | `MARKSANS PHARMA LTD` | `CLEAN SCIENCE AND TECHNOLOGY LIMITED` | label | Update company name |
| `WACC-DCF` | `E30` | `=H18` | `=D30` | capital_structure_target | Set target capital structure equal to company actual capital structure |
| `DCF 1st way` | `E42` | `164.63` | `=E40/E41` | reconciliation_formula | Connect native DCF formula E40/E41 directly to CFCL ramp FCFF and empirical WACC |
| `Model ->` | `C9` | `None` | `='DCF 1st way'!E42` | executive_summary_link | Dynamically link cover primary value to source-of-truth cell 'DCF 1st way'!E42 |

## Target Entity: EDELWEISS (EDELWEISS FINANCIAL SERVICES LIMITED)

- **Primary Valuation Method**: Financial Services Sum-of-the-Parts (SOTP) Valuation (Industrial FCFF Prohibited)
- **Primary Fair Value per Share**: **₹130.53**
- **Single Source-of-Truth Cell**: `'DCF 1st way'!E42`
- **Executive Cover Link**: `'Model ->'!C9` links dynamically to `='DCF 1st way'!E42`

| Sheet | Cell | Original Template Value | Target Adapted Value | Nature of Change | Rationale / Explanation |
|---|---|---|---|---|---|
| `Data Sheet` | `B1` | `MARKSANS PHARMA LTD` | `EDELWEISS FINANCIAL SERVICES LIMITED` | label | Update company name |
| `DCF 1st way` | `E42` | `N/A` | `=G60` | reconciliation_formula | Link primary valuation output to Financial Services SOTP schedule G60 |
| `Model ->` | `C9` | `None` | `='DCF 1st way'!E42` | executive_summary_link | Dynamically link cover primary value to source-of-truth cell 'DCF 1st way'!E42 |

