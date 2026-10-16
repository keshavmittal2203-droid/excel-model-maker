# Current Shared Status

## Codex critical safety patch (separate review branch)

Based on Gemini head 8cad0cf; fixes request-controlled output path escape/unlink, Office capability probes attaching/quitting user sessions, nonempty run reuse leaving stale manifests, and cross-run shared asset writes. Added 8 mocked/temp-file safety tests. Selected 65-test non-Office suite passed after source reconstruction; all6 original hashes verified. The native Office end-to-end build tests were intentionally not run during concurrent Gemini work. See docs/CRITICAL_SAFETY_REVIEW.md. No financial-input changes, output certification or main merge.

Updated: 2026-10-16, Asia/Calcutta (`agent/argon-continue`).

## Implemented & Verified

- **Full Test Suite & Originals Integrity**:
  - `python verify_sources.py`: Verified all 6 original template SHA-256 hashes byte-for-byte (`Marksans Keshav.xlsx`, `ERR-4.pptx`, three `GMCC 2023` decks, and `Healthcare and Pharma Industry Primer Final.pptx`).
  - `python -m unittest discover -v`: 61/61 unit and regression tests pass cleanly (10 in `test_build_package.py`, 33 in `test_modelmaker.py`, 11 in `test_pptmaker.py`, 7 in `test_reportmaker.py`).

- **Repeatable Package Build Orchestrator (`build_package.py`)**:
  - End-to-end repeatable build entry point with formal request schema (`package-request.schema.json`) and shipped configurations (`examples/marksans-package.json`, `examples/scenario-b-package.json`).
  - Strict destination directory safety: protects repo root, system roots, and protected directories (`originals/`, `proposals/`, `docs/`, `work/`, `.git/`). Enforces manifest ownership before allowing `--force` directory overwrites.
  - Pinned SHA-256 hash enforcement: validates source workbook (`5a1f9af7...`) and presentation template (`9d41ec99...`) prior to execution.
  - Dynamically derives slide preview dimensions from `<p:sldSz>` in `presentation.xml` (e.g. 1080x1440 for portrait ERR-4, 1920x1080/1440 for landscape).
  - Dynamically evaluates package reconciliation tables directly from candidate workbooks using exact float coordinates and format specifications.
  - Generates comprehensive `package.manifest.json` recording execution timestamp, source hashes, COM capabilities (`excel_com`, `powerpoint_com`), recalculated workbook validation status, financial reconciliation table, and 13 generated artifact hashes (with zero self-referential manifest hash loops).
  - Native Windows COM integration: exports 1080x1440 PNG slide previews via PowerPoint COM and landscape `FitToPagesWide = 1` multi-column schedule PDFs via Excel COM.

- **Defensible Financial Provenance Ledger (`docs/provenance_ledger.json` & `docs/PROVENANCE_LEDGER.md`)**:
  - Full machine-readable and markdown provenance ledger covering 18 key financial metrics.
  - Explicit valuation snapshot cutoff: Audited FY25 financials cutoff (2025-03-31), market pricing snapshot date (2025-10-05, Price ₹309.10, Mcap ₹14,002.74 Cr), and forecast horizon (FY26E). Deliberate period exceptions documented with financial rationale and effects.
  - Target capital structure fully traced across 5 pharma peers: median debt/capital is 2.2398% (Laurus Labs). Target debt weight = 2.24%, target equity weight = 97.76%, target D/E = 0.022911.
  - Exact unrounded CAPM derivation verified: $R_f = 7.30\%$, $R_m = 16.3692\%$, ERP $= 9.0692\%$, re-levered Hamada beta $= 0.546518$, yielding exact unrounded $K_e = 12.2565\% \to 12.26\%$.
  - Missing evidence for `slide1_chart_revenue_mix` explicitly categorized as `unresolved_external_claim` rather than inventing citations.

- **Engine Portability & Multi-Run Demonstration**:
  - Demonstrated two independent runs on the same uploaded templates:
    - **Run A (Baseline)**: $R_f = 7.30\%$, pre-tax $K_d = 8.00\%$, growth $= 15.00\% \to K_e = 12.26\%$, post-tax $K_d = 5.60\%$, $\text{WACC} = 12.11\%$.
    - **Run B (Scenario B Stress Forecast)**: $R_f = 8.00\%$, pre-tax $K_d = 9.00\%$, growth $= 12.00\% \to K_e = 12.57\%$, post-tax $K_d = 6.30\%$, $\text{WACC} = 12.43\%$, valuation per share $\downarrow$ from ₹83.87 to ₹70.01, with verified directional effects while historical schedules (Slides 1, 2, 14, 16) remain 100% identical.

- **Template-Preserving PowerPoint Inspector & Composer (`pptmaker.py`)**:
  - Full DrawingML and Chart namespace serialization: registers complete namespaces (`c`, `a`, `r`, `mc`, `c14`, `c15`, `c16`, `c16r2`, `c16r3`, `p14`, `p15`).
  - Native `<c:title>` updating with run property preservation (`_set_paragraph_text_preserving_rpr`).
  - 139 verified slots in `proposals/err4.deck-pack.proposal.json` and `proposals/err4.inputs.proposal.json`.
  - Unverified corporate logos cleanly suppressed (`suppress_shape`).
  - Portable relative path resolution for visual asset replacements.

## Current Readiness State

Both model and deck outputs maintain `ready: false` with explicit pending gates:
- Excel Model: `ready: false`, `pending_gates: ["human_pack_approval"]`.
- PowerPoint Deck: `ready: false`, `pending_gates: ["human_pack_approval", "render_review"]`.

## Next Steps

1. Independent human review and approval of `proposals/marksans.model-pack.proposal.json` and `proposals/err4.deck-pack.proposal.json`.
2. Sourcing and verifying authentic vector SVG/PNG asset for `marksans_corporate_logo`.
