# Excel Model Maker

Excel-only tools for building professional financial workbooks from the user's supplied Excel template. In accordance with **GitHub Issue #8**, **Mahindra & Mahindra (`MM`)** is the sole active generated company model for current delivery and review.

## Product rules

- Start from the active 16-sheet master workbook and preserve its sheets, layout, formulas, styles, widths, merged cells, logos, and other workbook structure. Do not create a replacement layout unless the user explicitly asks for a universal template.
- Use only the websites named in the master workbook plus Screener. The exact allowlist and evidence rules are in [docs/LOCKED_SOURCE_ALLOWLIST.md](docs/LOCKED_SOURCE_ALLOWLIST.md).
- Each factual input must be traceable to an allowed source; calculated values must be formulas; forecasts and valuation drivers that are not reported facts must be clearly labeled assumptions.
- Missing evidence remains blank or N/A. Never use another provider, an unexplained hardcode, or a plug to make a model look complete or balanced.
- Candidates remain unapproved until a human reviews them. Successful export, automated tests, or an agent's statement does not close the human approval gate.
- PowerPoint generation is retired. Historical `.pptx` reference files, if present in `originals/`, are inert archives and are not read by the Excel product.

## Excel workflows

### Build Single-Target Mahindra & Mahindra Template V17 Model (Issue #8)
A repeatable command builds the repaired Mahindra & Mahindra workbook directly from the unchanged 16-sheet master template (`originals/Marksans Keshav.xlsx`), executes native Microsoft Excel COM 16.0 recalculation, verifies zero circular references and zero formula errors across all 16 sheets, validates dynamic `BS!C40:M40` balance ties (0.00 Cr across all 11 periods), marks FY26E forecast explicitly blocked without plugs, and exports 15 landscape PDF previews:

```powershell
python work/build_template_v17_mm.py --clean
# OR via package builder:
python build_package.py --template-v17 --clean
```

For environments without Microsoft Excel COM, add `--skip-com` to perform the structural 16-sheet openpyxl build and formula checks without activating Office:
```powershell
python work/build_template_v17_mm.py --clean --skip-com
# OR via package builder:
python build_package.py --template-v17 --clean --skip-com
```

### Reproduce Four Template V16 Models (Archive Milestone)
To reproduce the historical four-target baseline (MM, LTM, CLEAN, EDELWEISS):
```powershell
python work/build_template_v16_models.py --output-dir staging/v16_reproducibility
# OR via package builder:
python build_package.py --template-v16 --output-dir staging/v16_reproducibility
```

### Environment & Tool Requirements
- **OS**: Windows (required for Microsoft Excel COM automation and PDF export).
- **Python**: Python 3.11 or later (Python 3.14.8 verified).
- **Core Dependencies**: `openpyxl`, `pywin32` (`win32com.client`), `pytest`.
- **Microsoft Excel**: Office / Microsoft 365 desktop Excel 16.0 installed and licensed.
- **Verification Suite**:
  ```powershell
  python verify_sources.py
  python -m unittest test_mm_repaired_v17.py
  ```

### Active Model Status & Issue #8 Repairs
The active generated model preserves the authentic 16-sheet template structure and allowlisted sources only (Screener.in + 6 original web sources):
- **Mahindra & Mahindra (MM)**:
  - 16 sheets in exact master sequence; exactly 0 circular references and 0 formula errors.
  - 11 historical balance ties verified: `BS!C40:M40 = 0.00 Cr` (`={col}38-{col}16`).
  - FY26E balance sheet marked `BLOCKED` (`BS!N40 = "BLOCKED"`) without artificial balancing plugs. Consolidated Other Assets includes the systemic NBFC loan book of MMFS (>₹1.3L Cr) without a published funding bridge.
  - All 5 circular formula loops in `compdata` (`F19:F25` ↔ `K4:K10`) eliminated; formula reverse-solving of EBITDA removed.
  - Valuation marked `Unavailable` in `DCF 1st way!E42` and cover `Model ->!C9`. SOTP segment facts are unavailable on approved sources; industrial FCFF DCF is prohibited for conglomerates.
  - Cost of capital: Raw OLS beta 1.4746, Blume adjusted beta 1.3164 (recomputed from Yahoo Finance 104-week return series). WACC marked `Unavailable` in `WACC-DCF!J39`.
  - Native `<x14:sparklineGroups>` XML restored in `Ratio Analysis` with data ranges `C6:G6` through `C26:G26`.
- **LTM, CLEAN, EDELWEISS**: Retired from active generated deliverables pursuant to Issue #8. Prior four-company delivery safely archived at `deliverables/archive/V16_user_handoff_four_company_archive.zip`. Source snapshots and registries remain preserved in repo.

Governance: All packages remain `ready: false` with `human_pack_approval: pending`. PR #8 remains open and unmerged.

## Project files

- `originals/Marksans Keshav.xlsx` — active master template (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`); preserve byte-for-byte.
- `work/sources/*_template_v16_delta.json` — committed source inputs and formulas for clean-checkout reproduction.
- `deliverables/V16_user_handoff.zip` — packaged user delivery archive containing the 4 validated workbooks and review documents; preserved intact.
- `docs/LOCKED_SOURCE_ALLOWLIST.md` — locked source and evidence policy (7 original sites + Screener.in).
- `docs/TEMPLATE_V16_SHEET_AUDIT_LEDGER.md` — 64-sheet verification ledger across all 4 targets.
- `STATUS.md` — authoritative verified state and governance tracking.
- `work/progress.md` — resumable engineering checkpoint and historical log.

The original workbook's exact structure is authoritative. The course link in the allowlist is a methodology reference only; it is not a data source.
