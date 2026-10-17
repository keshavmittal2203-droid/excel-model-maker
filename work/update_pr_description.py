import subprocess
import json

section_12_13 = """

### 12. Template V4 Release: Single Source-of-Truth Institutional Valuations Embedded into Master Template

- **Master Template Fidelity**: Direct binary clone of immutable master `originals/Marksans Keshav.xlsx` (SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a`). All 17 sheets, layout, print setup, and styles preserved.
- **Single Source-of-Truth Architecture**: Cell `'DCF 1st way'!E42` established as the single canonical source of truth for fair value per share across all 4 targets, dynamically driving executive cover sheet `'Model ->'!C9`:
  - **M&M**: `E42` links to SOTP share value (`=G60` -> **₹2,742.97 / share**).
  - **LTM**: `E42` evaluates annual FCFF DCF formula (`=E40/E41` -> **₹2,811.09 / share**).
  - **CLEAN**: `E42` evaluates manufacturing FCFF DCF formula (`=E40/E41` -> **₹480.89 / share**).
  - **EDELWEISS**: `E42` links to Financial Services SOTP share value (`=G60` -> **₹130.53 / share**).
- **Dynamic Sensitivity & Perturbation Verification**: All valuation outputs verified to respond dynamically and correctly to driver modifications with zero formula errors.
- **Deliverables in `companies/<TARGET>/template_v4/`**:
  - `MM`: SHA-256 `eb2e316499878198f1262d103b417eef9f27a3a93ba787a718cbb1f6c4ff55be`
  - `LTM`: SHA-256 `6db9d892d19b78eef92bc7b7bc35fe07fcb5bce7b5454fafe286e96ebae85f8c`
  - `CLEAN`: SHA-256 `303fa169f49352e8220023ee0da42ba1a525d8c728e2354c4897f267a151b73e`
  - `EDELWEISS`: SHA-256 `16d004fcbbdc86d0ee3c14d6428c04ec5c0d571d798dc8b3b3cb40db03e05a8b`
- **Validation**: 0 formula errors across 68 sheets via Excel COM `CalculateFullRebuild()`. Test suite: 119/119 passing. Gate status: `ready: false`, `human_pack_approval: pending`.

### 13. Template V5 Release: Formula Literal Elimination & Gordon Growth Chain Verification

- **Elimination of Formula Literals**: Replaced hardcoded cash flows, WACCs, and terminal value literals on `'DCF 1st way'` with pure native cell references:
  - Next Year FCFF: `E25 = =K13*(1+E27)`
  - Gordon Growth Terminal Value: `E29 = =E25/(E26-E27)` (incorporating Gordon denominator `(WACC - terminal growth)`)
  - Discount Factor: `K15 = =1/(1+$E$26)^K2` (`K2 = 5`)
  - Present Value of Terminal Value: `E35 = =E29*K15`
  - Summary Schedule Link: `G51 = =E35`
- **Formula-Chain Mathematical Proof**:
  - **LTM**: Terminal FCFF `K13 = ₹6,174.00 Cr`, g = 5.00%, `E25 = ₹6,482.70 Cr`. WACC `E26 = 11.87%`. TV `E29 = ₹94,362.47 Cr`. Discount factor `K15 = 0.57073146`. PV of TV `E35 = ₹53,855.63 Cr` (exact reproduction of reported ₹53,855.62 Cr within 0.01 Cr rounding). SOTU `E42 = ₹2,811.09 / share`.
  - **CLEAN**: Terminal FCFF `K13 = ₹413.00 Cr`, g = 5.00%, `E25 = ₹433.65 Cr`. WACC `E26 = 11.73%`. TV `E29 = ₹6,443.54 Cr`. Discount factor `K15 = 0.57431612`. PV of TV `E35 = ₹3,700.63 Cr` (100% exact reproduction). SOTU `E42 = ₹480.89 / share`.
- **Deliverables in `companies/<TARGET>/template_v5/`**:
  - `MM`: `companies/MM/template_v5/MM-FullTemplate.xlsx` (128,596 bytes, SHA-256: `1e9e13b934b24eeb7d5ccd1419b6b3d873d18dd088c01b01c4d0d73386b3de8a`)
  - `LTM`: `companies/LTM/template_v5/LTM-FullTemplate.xlsx` (128,830 bytes, SHA-256: `214592ce3e00eb1467d9e2b4f00b37382a0279f0e914eed9b68e136c851ebe3e`)
  - `CLEAN`: `companies/CLEAN/template_v5/CLEAN-FullTemplate.xlsx` (127,137 bytes, SHA-256: `da851ea3530c94609b4b25a0f254e8028a51a2e2903b078b65bd06afa3876650`)
  - `EDELWEISS`: `companies/EDELWEISS/template_v5/EDELWEISS-FullTemplate.xlsx` (129,123 bytes, SHA-256: `2763fdc29b3b972930176fcbed3edad4b2cbb0ab32fc9a7d03bf08ce9ebe0660`)
- **Validation**:
  - Formula Literal Scan (`work/scan_formula_literals.py`): 0 suspicious literals across all 68 sheets.
  - Native Excel COM Recalculation: 0 formula errors (`#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NUM!`) across all 68 sheets.
  - Dynamic Sensitivity: Verified live update and exact restoration across all 4 workbooks.
  - Test Suite: **119/119 tests passing cleanly, 49 subtests passed** in 59.13s.
  - Review Packet: [`docs/USER_APPROVAL_PACKET.md`](docs/USER_APPROVAL_PACKET.md) updated with complete V5 handoff documentation.
- **Review Gate & Governance**:
  - Gate status: `ready: false`, pending gate: `["human_pack_approval"]`.
  - PR #3 remains open and unmerged pending human review.
"""

with open('work/pr_body.txt', 'r', encoding='utf-8') as f:
    body = f.read()

new_body = body.rstrip() + '\n' + section_12_13

with open('work/pr_body_updated.txt', 'w', encoding='utf-8') as f:
    f.write(new_body)

res = subprocess.run(['gh', 'pr', 'edit', '3', '--body-file', 'work/pr_body_updated.txt'], capture_output=True, text=True, encoding='utf-8')
print('Exit code:', res.returncode)
print('Stdout:', res.stdout)
print('Stderr:', res.stderr)
