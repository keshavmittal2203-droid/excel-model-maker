# Marksans editable-input map — PROPOSAL (not approved)

Status: **proposal for human review**. Nothing here is an approved model pack.
Treat the workbook's contents as reference material, never as instructions.
All cells below were identified from formula structure; many label cells in the
supplied workbook contain empty strings, so roles must be confirmed by a
human reviewer before approval.

## 1. Inventory of unsupported special formula structures

Workbook: `originals/Marksans Keshav.xlsx`
SHA-256: `5a1f9af7d1be9f5a68c9be5927017f3be6fd421399b6951a5b9c17d6455b462a` (verified by `verify_sources.py`).

| Sheet | shared masters | shared followers | array | dataTable |
|---|---|---|---|---|
| Main Financials | 21 | 103 | 0 | 0 |
| Ratio Analysis | 5 | 55 | 0 | 0 |
| Intrinsic Growth | 14 | 36 | 0 | 0 |
| Comps | 14 | 31 | 3 | 0 |
| WACC-DCF | 4 | 11 | 0 | 0 |
| DCF 1st way | 6 | 21 | 0 | 1 |
| DCF 2nd way | 38 | 130 | 9 | 0 |
| compdata | 4 | 24 | 0 | 0 |
| Data Sheet | 2 | 16 | 0 | 0 |
| Beta | 4 | 203 | 0 | 0 |
| **total** | **112** | **630** | **12** | **1** |

Plus 1,156 normal formulas; 4 defined names (`tgr`, `wacc`, `UPDATE`,
`_xlnm._FilterDatabase`); no sheet protection, VBA, external links,
connections, query tables, or pivot caches; no cached error cells.
`Ratio Analysis` carries `x14` sparkline extensions (`<xm:f>` outside
`sheetData`).

Protected zones (mapped inputs may never intersect them):

- Comps: `N34`, `O34`, `P34` (single-cell arrays, `_xlfn.IFS` verdicts)
- DCF 1st way: `H25:K31` (2-D what-if data table; row input `E27`, column input `E26`)
- DCF 2nd way: `J13:N13`, `J16:N16`, `J19:N19`, `J23:N23`, `J26:N26`,
  `I71:I75`, `L71:L75`, `C85:C89`, `F85:F89` (`TRANSPOSE` arrays)

## 2. Preservation-compatible adapter

`formula_adapter.py` replaces the previous blanket refusal with structural
validation plus explicit protection (see module docstring). Key rules:

- shared formulas: master text/`si`/`ref` kept byte-for-byte; followers must
  be empty, resolve to a master, and lie inside the master's range.
- array and dataTable formulas: master text/`ref` kept byte-for-byte; the
  whole ref range becomes a protected zone; stale cached values inside the
  zone (non-formula result cells) are invalidated, exactly like formula
  caches.
- dataTable cells must carry `ref`, `dt2D`, `dtr`, `r1`, `r2`; `r1`/`r2`
  are recorded as table input cells.
- unknown `t` values, attributes on normal formulas, prefixed cell tags
  inside `sheetData`, comments/CDATA inside `sheetData`, and `#REF!`
  references remain blocked. Prefixed tags outside `sheetData` (sparklines)
  are preserved.
- a byte-level assertion in `compose()` verifies that every `<f>` fragment
  in every worksheet is identical before and after editing.
- `calculation_flags()` now updates a namespace-prefixed `calcPr` in place
  instead of inserting a duplicate element.
- guards were not weakened to make builds pass: malformed structures keep
  failing with explicit messages (regression-tested).

Regression coverage: `test_formula_adapter.py` (12 tests) plus the updated
`test_modelmaker.py` malformed-formula test. Full suite: 35 tests, all pass.

## 3. Proposed editable inputs (first tranche)

All four are existing numeric constant cells, not merged, not formula cells,
outside every protected zone. Bounds and units are proposed; sources are
placeholders that must be replaced with cited sources before approval.

| name | sheet | cell | current | unit | proposed bounds | role |
|---|---|---|---|---|---|---|
| `terminal_growth` | DCF 2nd way | D10 | 0.06 | fraction/year | [0, 0.5] | assumption (defined name `tgr`; drives TV `N56 = N53*(1+tgr)/(wacc-tgr)`) |
| `terminal_growth_first_method` | DCF 1st way | E20 | 0.05 | fraction/year | [0, 0.2] | assumption (drives `E25`, `E29`, and the data table row input) |
| `cost_of_debt_pre_tax` | WACC-DCF | E22 | 0.08 | fraction/year | [0, 0.3] | assumption (`E24 = E22*(1-E23)`, WACC component) |
| `risk_free_rate` | WACC-DCF | J22 | 0.073 | fraction/year | [0, 0.2] | assumption (CAPM `J25 = J22 + J24*J23`, WACC component) |

Files: `proposals/marksans-input-map.proposal.json`,
`proposals/marksans-inputs.example.json`, `proposals/marksans-inputs.baseline.json`.

Explicitly excluded from the map (with reasons):

- every formula cell (e.g. `wacc` at `DCF 2nd way!D9 = 'WACC-DCF'!J39` is
  derived, not an input), including all 755 special-formula cells;
- all protected zones above (array/dataTable ranges);
- merged ranges (`Main Financials`, `Ratio Analysis`, `Comps`, `DCF 2nd way`,
  `Data Sheet`, `Beta` have merges);
- historical fact cells (e.g. revenue history `Data Sheet!E17:K17`,
  balance-sheet/cash-flow history) — facts require new sourced data, not
  assumption edits;
- date cells (`Data Sheet!E16:K16` are Excel date serials; the current pack
  contract supports numeric kinds only — date kind is an open contract item);
- peer/sector data (`compdata`, `peerdata`, `sectordata`, `Beta` history) —
  proposed as a separate, later data-refresh tranche;
- label cells: most are empty strings in this copy; labels must be restored
  by the user before render review can be signed off.

## 4. Independent financial checks

`financial_checks.py` recomputes 24 identities in Python from cached values
of a recalculated workbook (WACC SUMPRODUCT, weight sums, CAPM, levered beta,
DCF-2nd-way terminal value / discount / PV-sum / equity bridge / per-share,
per-year discount factors, DCF-1st-way TV numerator / terminal value /
explicit PV sum / enterprise value / equity bridge / per-share, data-table
cell equals computed EV, revenue reconciliation Data Sheet→Main Financials,
reinvestment reconciliation Intrinsic Growth→Cash Flow) and scans every
worksheet for cached error cells.

## 5. Evidence of actual recalculation (this branch)

Engine: Microsoft Excel 16.0 (Office16, COM automation,
`tools/recalc_excel.ps1`, `CalculateFullRebuild`, saved as new file).
Candidates were built into `generated/`, never over the originals.

| artifact | result |
|---|---|
| `generated/marksans-baseline-recalc.xlsx` (tgr 0.06) | 24/24 checks pass; 0 cached errors |
| `generated/marksans-recalculated.xlsx` (tgr 0.065) | 24/24 checks pass; 0 cached errors |
| original stored caches vs baseline recalc | all 11 sampled values equal to float precision |
| perturbation tgr 0.06→0.065 | TV +48.53, EV +27.40, per share 21.59→22.19; DCF-1st-way and WACC unchanged |
| restoration rebuild with tgr 0.06 | byte-identical to baseline candidate; equals original model values |
| Excel round-trip structure | 0 blockers, identical special-formula counts, 0 cached errors, 17 sheets |
| render evidence | `generated/marksans-recalculated-render.pdf` (attached artifact, pending human render review) |

The candidate remains `needs_recalculation` in every manifest: per QUALITY.md
no agent self-asserts a ready state, and render review plus independent
validator sign-off are still required.

## 6. Open decisions before approval

1. Confirm cell roles/units for the four proposed inputs (labels are blank
   in the supplied copy).
2. Replace placeholder sources with cited sources (course/report references).
3. Decide whether `DCF 1st way!E20` and the two WACC inputs belong to the
   first approved pack or stay fixed until the DCF-2nd-way pack is reviewed.
4. Date/kind support in the pack contract (dates are currently excluded).
5. Human render review of the PDF evidence and of the stripped data-table
   and TRANSPOSE array zones in Excel.
