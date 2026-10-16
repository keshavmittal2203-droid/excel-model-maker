# Current shared status

Updated: 2026-10-16, Asia/Calcutta (`agent/argon-continue`).

## Implemented

- **Dependency-free Python XLSX Inspector, Composer & Validator (`modelmaker.py` + `validate_workbook.py`)**:
  - Template/approval hashes, exact input mapping, numeric bounds/precision checks, source references, and original/output overwrite protection.
  - Formula-cache invalidation, calculation-chain cleanup, recalculation flags, semantic worksheet-change guard, and unchanged native package-part preservation.
  - **Special formula preservation adapter**:
    - Supports `shared` formulas (master + follower groups), single/multi-cell `array` formulas, and 2-variable sensitivity `dataTable` structures across all 17 worksheets of `originals/Marksans Keshav.xlsx` (10 sheets with special formulas: 112 shared groups, 620 followers, 12 array formulas, 1 2D `dataTable` at `DCF 1st way!H25:K31`, and `<x14:sparklines>` on `Ratio Analysis`).
    - Formula range collision guard strictly rejects any mapped input intersecting normal formulas, shared groups, array ranges, `dataTable` ranges, or merged cells.
  - **Comprehensive 93-Input Equity Research Model Pack (`marksans.model-pack.json` & `marksans.inputs.json`)**:
    - Covers all 10 editable input sheets of `originals/Marksans Keshav.xlsx`: `Data Sheet` (company meta, 5Y historical P&L, balance sheet, and cash flow drivers), `BS` (working capital & balance sheet totals), `Cash Flow` (operating cash flow & 5Y CapEx), `peerdata` (5 peer companies' market cap & debt), `compdata` (comparable company valuation multiples & TTM financials), `sectordata` (Damodaran pharma & healthcare reinvestment rates), `WACC-DCF` (peer tax rates, levered betas, cost of debt, risk-free rate, capital structure), `Intrinsic Growth` (marginal tax rates), `DCF 1st way` (explicit growth, reinvestment rates, terminal growth), and `DCF 2nd way` (terminal growth `tgr`).
  - **Recalculation & Independent Financial Check Runner (`validate_workbook.py` / `python modelmaker.py validate`)**:
    - Recalculates candidate `.xlsx` workbooks via installed Microsoft Excel COM (`Excel.Application` 16.0 headless `CalculateFullRebuild()`) or the built-in Python formula evaluator for standard workbooks.
    - Scans all worksheets for Excel formula errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`) and missing formula caches (`<v>`).
    - Evaluates 19 model-pack checks and **29 independent Python-recomputed accounting & valuation identities** (WACC chain, CAPM, Hamada levered beta, DCF 1st way & 2nd way terminal value / PV / EV / equity bridge / per-share value, 2D `dataTable` center tie to EV, Balance Sheet `Assets == Liabilities` across all historical years, cross-sheet revenue and reinvestment reconciliations, and Comps EV identity).
    - Automates **input perturbation and restoration sensitivity checks (`--perturb-check`)**, confirming that perturbing `dcf2_terminal_growth_rate` (`0.06 -> 0.065`) moves DCF 2nd way implied value per share (`21.59 -> 22.19`) while preserving 0 errors and all 29 identities, and that restoring the original input produces a byte-for-byte identical candidate (`sha256` match).
- **Template-Preserving PowerPoint Inspector & Composer (`pptmaker.py`, `deck-pack.schema.json`, `err4.deck-pack.json`, `err4.inputs.json`)**:
  - Inspects all 5 reference `.pptx` decks reconstructed by `verify_sources.py` in `originals/` (182 slides total across `ERR-4.pptx`, the three `GMCC 2023` decks, and `Healthcare and Pharma Industry Primer Final.pptx`).
  - Supports `shape_text`, `shape_token`, and `table_cell` replacements while preserving run-level formatting (`<a:rPr>`), paragraph properties (`<a:pPr>`), table cell properties (`<a:tcPr>`), slide layouts, masters, charts, media, and unedited slides byte-for-byte.
  - Enforces provenance citations (`source`) and optional `workbook_binding` verification against recalculated `.xlsx` cells (`candidates/Marksans-Recalculated.xlsx` -> `candidates/ERR-4-Marksans-Candidate.pptx`).
- **Universal AI Onboarding Playbook (`AI_PLAYBOOK.md`)**:
  - End-to-end instructions for any AI agent to inspect templates, map researched company/industry data into `inputs.json`, build and recalculate `.xlsx` models, compose `.pptx` presentations bound to `.xlsx` outputs, and run the 35-test verification suite.

## Verified progress

- `python verify_sources.py`: all 6 original files in `originals/` verified against `source-catalog.json`.
- `python -m unittest -v`: **Ran 35 tests, OK** (28 in `test_modelmaker.py` + 7 in `test_pptmaker.py`).
- Real Excel 16.0 COM recalculation and `--perturb-check` validation verified on `candidates/Marksans-Candidate.xlsx` -> `candidates/Marksans-Recalculated.xlsx` (`status: validated`, `ready: true`, `0` formula errors, `19/19` pack checks passed, `29/29` independent identities passed).
- Real PowerPoint deck composition verified on `originals/ERR-4.pptx` -> `candidates/ERR-4-Marksans-Candidate.pptx` (`status: composed`, `ready: true`, 18 slots updated across 5 slides, 227 package parts preserved byte-for-byte).

## Next steps

1. Human review and merge of open PRs (`#1`, `#2`, and the `agent/argon-continue` PR) into `main`.
2. Connect Supabase / versioned cloud object storage under `keshavmittal2203@gmail.com` after checking project cost.
3. Deploy shared HTTPS/MCP service and cloud recalculation worker wrapping `modelmaker.py`, `validate_workbook.py`, and `pptmaker.py`.

## Current limits / blockers

- GitHub repository access is active and verified (`public` visibility, push access enabled).
- Cloud database/storage and hosted HTTP/MCP server deployment remain future infrastructure tasks; local CLI and desktop Excel COM recalculation are fully operational.
