# Current Shared Status

## Codex Critical Safety Patch Integrated & Verified

Integrated Codex PR #7 (`1395b247d4a1e87ef83b4938c47a88be200b8b4f`) into `agent/argon-continue`.
- Rejects request-controlled output paths with directory traversal (`../`), absolute paths, Windows device names (`CON`), or alternate data streams (`file:stream`) before any writes.
- Protects existing output files (never replaced) and prohibits non-empty output directory reuse, even with `--force`.
- Office COM capability detection probes Windows registry directly without launching or quitting user Office applications.
- Full process isolation: Excel automation uses dedicated `DispatchEx` instance with links/macros disabled; PowerPoint export closes only its own presentation without calling application `Quit()`.
- Run-local image containment: dynamic visual assets generate exclusively inside `candidates/<run_dir>/assets/` with zero writes to `work/assets/`.

Updated: 2026-10-16, Asia/Calcutta (`agent/argon-continue`).

## Implemented & Verified

- **Full Test Suite & Originals Integrity**:
  - `python verify_sources.py`: Verified all 6 original template SHA-256 hashes byte-for-byte (`Marksans Keshav.xlsx`, `ERR-4.pptx`, three `GMCC 2023` decks, and `Healthcare and Pharma Industry Primer Final.pptx`).
  - `python -m unittest discover -v`: **69/69 tests passing** in 26s across `test_package_safety.py` (8), `test_build_package.py` (10), `test_modelmaker.py` (33), `test_pptmaker.py` (11), and `test_reportmaker.py` (7).

- **Repeatable Package Build Orchestrator (`build_package.py`)**:
  - End-to-end repeatable build entry point with formal request schema (`package-request.schema.json`) and shipped configurations (`examples/marksans-package.json`, `examples/scenario-b-package.json`).
  - Strict destination directory safety: requires fresh output directories; enforces source workbook (`5a1f9af7...`) and presentation template (`9d41ec99...`) SHA-256 hashes prior to execution.
  - Generates comprehensive `package.manifest.json` recording execution timestamp, source hashes, COM capabilities (`excel_com`, `powerpoint_com`), recalculated workbook validation status, financial reconciliation table, and 13 generated artifact hashes (with zero self-referential manifest hash loops).
  - Native Windows COM integration: exports 1080x1440 PNG slide previews via PowerPoint COM and landscape `FitToPagesWide = 1` multi-column schedule PDFs via Excel COM.
  - Successfully executed fresh native builds:
    - Baseline: `candidates/run_native_baseline/` (13 artifacts, WACC 12.11%)
    - Scenario B: `candidates/run_native_scenario_b/` (13 artifacts, WACC 12.43%)

- **Defensible Financial Provenance Ledger (`docs/provenance_ledger.json` & `docs/PROVENANCE_LEDGER.md`)**:
  - Full machine-readable and markdown provenance ledger covering 18 key financial metrics, strictly categorized into: `externally_verified_fact`, `inherited_template_value`, `calculated_estimate`, `analyst_assumption`, or `unresolved_claim`.
  - Valuation snapshot:
    - Historical financials cutoff: Audited FY25 financials cutoff (`2025-03-31`).
    - Market retrieval date: `2025-10-05` (Sunday Screener snapshot).
    - Trading observation date: `2025-10-03` (Friday closing CMP ₹309.10, Market Cap ₹14,002.74 Cr).
    - Forecast horizon start: `FY2026E`.
  - Methodology notes:
    - Nifty 50: 26-year arithmetic average of Price Return (PR) index returns ($15.15\%$) + $1.22\%$ dividend yield assumption $= 16.37\%$ (no dividend double-counting).
    - Capital structure: 5-peer median debt/capital ($2.2398\%$, Laurus Labs) yielding target debt weight $2.24\%$, equity weight $97.76\%$, target D/E $0.022911$ following Damodaran pure-play methodology.
  - Geographic Revenue Evidence Gap Resolved:
    - Slide 1 Chart 0 figures `[43.1, 41.7, 9.8, 5.4]` have no supporting primary filing in `originals/`.
    - Retitled chart to `"Geographic Revenue Mix (Unavailable — Pending Primary Filing)"`, series to `"Unverified (Primary Filing Not Uploaded)"`, and classified as `unresolved_claim`.

## Current Readiness State

Both model and deck outputs maintain `ready: false` with explicit pending gates:
- Excel Model: `ready: false`, `pending_gates: ["human_pack_approval"]`.
- PowerPoint Deck: `ready: false`, `pending_gates: ["human_pack_approval", "render_review"]`.

## Next Steps

1. Human review and approval of proposals: `proposals/marksans.model-pack.proposal.json` and `proposals/err4.deck-pack.proposal.json`.
2. Sourcing primary investor presentation PDF for Marksans geographic segment reporting.
3. Sourcing verified vector SVG/PNG asset for Marksans corporate logo.
