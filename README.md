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

The repository contains a template-preserving XLSX composer (`modelmaker.py`), a package builder (`build_package.py`), company-specific workbook adapters under `work/`, Excel validation (`validate_workbook.py`), and a JSON-driven styled Excel report builder (`cli.py`).

Typical commands from the repository root:

```powershell
python verify_sources.py
python build_package.py --request examples/mm-package.json --output-dir candidates/mm_run
python build_package.py --batch
python -m unittest discover -v
```

Use a new output directory for every build. Never edit an original workbook in place. For financial-model adapters, Microsoft Excel COM recalculation and a rendered review are required before an output can be described as verified. See [QUALITY.md](QUALITY.md).

## Project files

- `originals/Marksans Keshav.xlsx` — active master template; preserve it byte-for-byte.
- `work/sources/` — versioned source snapshots and extraction evidence; only allowlisted sources may supply values.
- `companies/` — company-specific candidate workbooks, source registries, and previews.
- `docs/LOCKED_SOURCE_ALLOWLIST.md` — source and evidence policy.
- `STATUS.md` — current verified state and remaining blockers.
- `work/progress.md` — resumable engineering checkpoint.

The original workbook's exact structure is authoritative. The course link in the allowlist is a methodology reference only; it is not a data source.
