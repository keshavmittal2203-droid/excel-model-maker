# Excel Model Maker

## Package build safety

Use a new output directory for each package run. Output names must be plain `.xlsx` / `.pptx` filenames. Existing files and nonempty run directories are never overwritten, including with the legacy `--force` flag. COM capability checks do not launch Office; PowerPoint export closes only its own presentation, and replacement images remain isolated per run. [Safety review and actual test results](docs/CRITICAL_SAFETY_REVIEW.md).

Build financial models inside approved Excel templates, with traceable inputs and explicit validation evidence. Designed for a shared cloud workspace usable by OpenAI, Gemini, DeepSeek, Claude, and other clients that support HTTP tools or MCP.

It also ships a standalone styled report builder: turn a JSON spec into a professional Excel report with `python cli.py --input data.json --output report.xlsx`.

## Current working slice

The Python core inspects XLSX and PPTX files and fills **existing numeric input cells** (in `.xlsx`) and **approved text/table slots** (in `.pptx`) from explicit template packs:
- **Excel Composer & Validator (`modelmaker.py` + `validate_workbook.py`)**: Preserves all unrelated ZIP package parts, formula expressions (including shared formulas, array formulas, and 2D what-if `dataTable` structures), formatting, sparklines, charts, and workbook layout. Strips stale formula caches (`<v>`) on `build` and transitions a candidate from `needs_recalculation` to `validated` only after actual engine recalculation, a full formula-error scan, 29 independent accounting/valuation checks, and perturbation/restoration verification.
- **PowerPoint Inspector & Composer (`pptmaker.py`)**: Inspects `.pptx` templates (182 slides across the 5 reference decks in `originals/`) and builds candidate presentations by replacing mapped shape text, tokens, and table cells while preserving `<a:rPr>` run formatting, slide layouts, masters, charts, and media byte-for-byte. Enforces source citations and optional cell-level bindings to recalculated `.xlsx` workbooks.

Six user source files are catalogued in [SOURCES.md](SOURCES.md). Run `python verify_sources.py` to restore and verify the original files after cloning.

## Styled report builder

`cli.py` builds a new, standalone presentation workbook from a JSON spec: colored bold headers, frozen header row, banded rows, number/currency/percent/date formats, estimated column widths, SUM totals as real Excel formulas, conditional formatting, and an embedded chart. It never edits templates and never evaluates formulas; Excel computes the totals when the file is opened.

    pip install -r requirements.txt
    python cli.py --input report.example.json --output report.xlsx

The command refuses to overwrite an existing file unless `--force` is passed. The spec format, rules, and a full example are in [AGENTS.md](AGENTS.md); a working spec ships as [report.example.json](report.example.json).

## Environment & Software Dependencies

The repository is self-contained with respect to templates, references, and test fixtures (verified by `python verify_sources.py`). Runtime execution has standard software dependencies:
- **Python**: Python 3.11 or later (`openpyxl`, `Pillow`).
- **Microsoft Office (Windows)**: Excel and PowerPoint (Office 16.0+) for live workbook recalculation, PDF schedule export, and slide PNG preview rendering via `pywin32`.
- When Office COM is absent or non-Windows, core OOXML inspection and composition run with `--skip-com`.

## Package Build Command

To generate a complete, reconciled financial model and 5-slide executive presentation:

    python verify_sources.py
    python build_package.py --request examples/marksans-package.json --output-dir dist/marksans-run-01 --allow-proposal

**Important**: Each package run requires a **fresh, non-existent output directory** (e.g. `dist/marksans-run-01`, `dist/marksans-run-02`). Per safety rules, existing output files and non-empty directories are never overwritten, even with `--force`.

## Individual Tool Commands

    python modelmaker.py inspect "originals/Marksans Keshav.xlsx"
    python modelmaker.py build --template "originals/Marksans Keshav.xlsx" --pack proposals/marksans.model-pack.proposal.json --inputs marksans.inputs.json --output dist/Marksans-Candidate.xlsx
    python modelmaker.py validate dist/Marksans-Candidate.xlsx --pack proposals/marksans.model-pack.proposal.json --recalculate --recalculated-output dist/Marksans-Recalculated.xlsx --perturb-check --template "originals/Marksans Keshav.xlsx" --inputs marksans.inputs.json
    python pptmaker.py inspect "originals/ERR-4.pptx"
    python pptmaker.py build --template "originals/ERR-4.pptx" --pack proposals/err4.deck-pack.proposal.json --inputs proposals/err4.inputs.proposal.json --workbook dist/Marksans-Recalculated.xlsx --output dist/Marksans-Pharma-Overview.pptx
    python -m unittest discover -v

## Cloud source of truth

- Code, instructions, contracts, and reviewed model specifications: this GitHub repository.
- Original Excel files, courses, extracted source references, generated workbooks, and reports: versioned cloud object storage.
- Agent tasks, model-pack revisions, build runs, leases, and audit events: shared PostgreSQL database.
- Generation and recalculation: cloud workers operating on immutable snapshots.
- AI access: the same authenticated HTTPS API, OpenAPI contract, and MCP tools for every provider.

Agents can work simultaneously on separate branches, tasks, and build runs. Publishing a revision uses compare-and-swap checks; writing one workbook uses a lease. A binary XLSX is never merged by concurrent writers.

## Read next

- [Universal AI Playbook (Excel + PPTX)](AI_PLAYBOOK.md)
- [Architecture and AI access](ARCHITECTURE.md)
- [Workbook quality policy](QUALITY.md)
- [Agent collaboration rules](AGENTS.md)
- [Roadmap and required inputs](ROADMAP.md)
- [Excel model-pack contract](model-pack.schema.json)
- [PowerPoint deck-pack contract](deck-pack.schema.json)
- [Input contract](inputs.example.json)
- [API contract](openapi.json)

## Deployment status

The core is a development foundation, not a deployed cloud app. Database/storage provisioning, HTTP/MCP implementations, Excel recalculation workers, and model/course onboarding remain open. Runtime cloud ownership requested: keshavmittal2203@gmail.com. Repository owner: keshavmittal2203-droid. No paid resources are required to run the core checks.
