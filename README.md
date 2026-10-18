# Excel Model Maker

Excel-only tools for building professional financial workbooks from the user's supplied Excel template. The four current priority targets are Mahindra & Mahindra (MM), LTIMindtree (LTM), Clean Science (CLEAN), and Edelweiss Financial Services (EDELWEISS).

## Product rules

- Start from the active 16-sheet master workbook and preserve its sheets, layout, formulas, styles, widths, merged cells, logos, and other workbook structure. Do not create a replacement layout unless the user explicitly asks for a universal template.
- Use only the websites named in the master workbook plus Screener. The exact allowlist and evidence rules are in [docs/LOCKED_SOURCE_ALLOWLIST.md](docs/LOCKED_SOURCE_ALLOWLIST.md).
- Each factual input must be traceable to an allowed source; calculated values must be formulas; forecasts and valuation drivers that are not reported facts must be clearly labeled assumptions.
- Missing evidence remains blank or N/A. Never use another provider, an unexplained hardcode, or a plug to make a model look complete or balanced.
- Candidates remain unapproved until a human reviews them. Successful export, automated tests, or an agent's statement does not close the human approval gate.
- PowerPoint generation is retired. Historical `.pptx` reference files, if present in `originals/`, are inert archives and are not read by the Excel product.

## Excel workflows

### Reproduce Four Template V16 Models from Committed Inputs
A single repeatable command builds the four priority models (MM, LTM, CLEAN, EDELWEISS) directly from the unchanged 16-sheet master template (`originals/Marksans Keshav.xlsx`) and committed delta inputs in `work/sources/`, executes native Microsoft Excel COM 16.0 recalculation, verifies zero formula errors across all 64 sheets, validates dynamic `BS!C40:N40` balance checks and valuation outputs, and exports 15 landscape PDF previews per company (60 total):

```powershell
python work/build_template_v16_models.py --output-dir staging/v16_reproducibility
# OR via package builder:
python build_package.py --template-v16 --output-dir staging/v16_reproducibility
```

For environments without Microsoft Excel COM, add `--skip-com` to perform the structural 16-sheet openpyxl build and formula checks without activating Office:
```powershell
python work/build_template_v16_models.py --output-dir staging/v16_reproducibility --skip-com
```

### Environment & Tool Requirements
- **OS**: Windows (required for Microsoft Excel COM automation and PDF export).
- **Python**: Python 3.11 or later (Python 3.14.8 verified).
- **Core Dependencies**: `openpyxl`, `pywin32` (`win32com.client`), `pytest`.
- **Microsoft Excel**: Office / Microsoft 365 desktop Excel 16.0 installed and licensed.
- **Verification Suite**:
  ```powershell
  python verify_sources.py
  pytest
  ```

### Target Models, Balance Checks & Valuation Status
All four models preserve the authentic 16-sheet template structure and allowlisted sources only (Screener.in + 6 original web sources). Dynamic balance checks in `BS!C40:N40` (`={col}38-{col}16`) are exposed without balancing plugs:
- **Mahindra & Mahindra (MM)**: Historical model complete (`FY15–FY25 = 0.00 Cr`). FY26E balance check exposes authentic conglomerate gap of `+₹7,677.49 Cr` (`BS!N40`) because Screener consolidates the MMFS NBFC loan book (>₹1.3L Cr) into Other Assets without sub-schedules; forecast is **BLOCKED**. Valuation is **Unavailable** (`N/A`) because required segment EBITDA and standalone net cash disclosures are absent from permitted sources.
- **LTIMindtree (LTM)**: Historical model reconciles (`FY15–FY25 = 0.00 Cr`). FY26E balance check is `+₹1.00 Cr` (`BS!N40`), reflecting Screener's integer-rounded sub-schedule delta preserved without plugs (**UNRECONCILED**). Valuation is supported at **₹2,620.24 / share** (WACC 12.48%, adjusted beta 1.1018).
- **Clean Science (CLEAN)**: Historical model reconciles (`FY15–FY25 = 0.00 Cr`). FY26E balance check is `−₹1.13 Cr` (`BS!N40`), reflecting Screener's integer/decimal sub-schedule delta preserved without plugs (**UNRECONCILED**). Valuation is supported at **₹393.03 / share** (WACC 13.19%, adjusted beta 1.2141).
- **Edelweiss Financial Services (EDELWEISS)**: Historical model reconciles (`FY15–FY25 = 0.00 Cr`). FY26E balance check is `+₹1.00 Cr` (`BS!N40`), reflecting Screener's integer-rounded sub-schedule delta preserved without plugs (**UNRECONCILED**). Valuation is **Unavailable** (`N/A`) because industrial DCF is prohibited for financial institutions and vertical SOTP facts are absent from permitted sources.

Governance: All packages remain `ready: false` with `human_pack_approval: pending`. PR #3 remains open and unmerged.

## Project files

- `originals/Marksans Keshav.xlsx` — active master template (SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`); preserve byte-for-byte.
- `work/sources/*_template_v16_delta.json` — committed source inputs and formulas for clean-checkout reproduction.
- `deliverables/V16_user_handoff.zip` — packaged user delivery archive containing the 4 validated workbooks and review documents; preserved intact.
- `docs/LOCKED_SOURCE_ALLOWLIST.md` — locked source and evidence policy (7 original sites + Screener.in).
- `docs/TEMPLATE_V16_SHEET_AUDIT_LEDGER.md` — 64-sheet verification ledger across all 4 targets.
- `STATUS.md` — authoritative verified state and governance tracking.
- `work/progress.md` — resumable engineering checkpoint and historical log.

The original workbook's exact structure is authoritative. The course link in the allowlist is a methodology reference only; it is not a data source.
