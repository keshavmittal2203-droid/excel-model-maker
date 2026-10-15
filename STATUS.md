# Current shared status

Updated: 2026-10-16, Asia/Calcutta (second agent session, adapter + recalc work).

## Implemented

- Dependency-free Python XLSX inspector and conservative existing-numeric-input composer.
- Template/approval hashes, exact input mapping, numeric bounds/precision checks, source references, original/output overwrite protection.
- Formula-cache invalidation, calculation-chain cleanup, recalculation flags, semantic worksheet-change guard, and unchanged native package-part preservation.
- **Special-formula compatibility adapter** (`formula_adapter.py`): structural validation and byte-preservation for shared formulas (112 masters + 630 followers), array formulas (12), and dataTable (1) in the Marksans workbook; protected zones for array/dataTable ranges; stale-cache stripping inside zones; prefixed-tag handling inside/outside `sheetData` (sparklines preserved); prefixed `calcPr` updated without duplication. Malformed structures stay blocked with explicit messages — no guard was bypassed.
- **Independent financial checks** (`financial_checks.py`): 24 Python-recomputed identities (WACC chain, CAPM, levered beta, both DCF methods, data-table-vs-EV, cross-sheet revenue and reinvestment reconciliations) plus a cached-error scan, run against actually recalculated files only.
- **Excel recalculation tool** (`tools/recalc_excel.ps1`): headless COM recalculation with Microsoft Excel 16.0 and save-as to a new file; `tools/compare_values.py` for perturbation/restoration evidence.
- Explicit-but-unapproved editable-input map proposal (`docs/editable-input-map.md`, `proposals/marksans-input-map.proposal.json` + example/baseline inputs).
- 38 focused tests pass locally (23 original + 15 new adapter/checks tests).

## Verified progress (this agent, on this branch)

- `python verify_sources.py`: all six originals verified (hashes match `source-catalog.json`).
- `python -m unittest -v`: **Ran 38 tests, OK** (Python 3.14.8).
- Marksans workbook now passes the adapter gate with **no unsupported features**; special formula inventory locked by regression tests (742 shared, 12 array, 1 dataTable across 10 sheets).
- Real recalculation performed with Excel 16.0: baseline (tgr 0.06) and perturbed (tgr 0.065) candidates recalculated; **24/24 financial checks pass on both, 0 cached error cells**; original stored caches match the baseline recalc to float precision; perturbation moves DCF-2nd-way value (per share 21.59 → 22.19) while DCF-1st-way and WACC stay fixed; restoration rebuild is byte-identical to the baseline candidate. Render PDF exported for human review (`generated/marksans-recalculated-render.pdf`, local artifact).
- Candidates remain `needs_recalculation` in every manifest: no ready transition, per QUALITY.md.

## Next steps

1. Human review of `docs/editable-input-map.md`: confirm the four proposed inputs' roles/units/bounds, replace placeholder sources with cited references, and sign off (or amend) the pack before it becomes an approved revision.
2. Human render review of the recalculated candidate and the stripped array/dataTable zones.
3. Merge this branch's PR after review; keep originals immutable.
4. Run the independent AI smoke checks in AI_START_HERE.md on other assistants.
5. Connect Supabase under keshavmittal2203@gmail.com and check project cost before provisioning.
6. Implement/deploy shared file storage, task/job database, HTTP/MCP service, and compatible calculation worker.

## Current limits / blockers

- Repository access is resolved: the repo was made public for read and the agent account was granted write. This work is pushed on branch `agent/deepseek-continue` (commit `997887a` on top of main `6d10f4e8d8017f4fd6b73f71a210e4cde8b661ea`) with **PR #1 open**: https://github.com/keshavmittal2203-droid/excel-model-maker/pull/1. The PR is not merged, per the no-merge rule; human review is required.
- No deployed API/MCP service, shared runtime database, calculation worker, or approved real model pack yet. The desktop-Excel recalculation path is local evidence only, not a cloud worker.
- Many label cells in the supplied workbook are empty strings; cell roles were derived from formula structure and need human confirmation.
- Date-valued inputs and peer/sector data refresh are out of scope of the current pack contract (documented in the map proposal).
