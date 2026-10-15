# Excel Model Maker

Build financial models inside approved Excel templates, with traceable inputs and explicit validation evidence. Designed for a shared cloud workspace usable by OpenAI, Gemini, DeepSeek, Claude, and other clients that support HTTP tools or MCP.

## Current working slice

The Python core inspects XLSX files and fills **existing numeric input cells** from an approved template map. It preserves other ZIP package parts, formula expressions, formatting, charts, and workbook layout. It removes stale formula caches and requests full recalculation. A preservation-compatible adapter (`formula_adapter.py`) now supports well-formed shared, array, and dataTable formula structures with protected zones, and `financial_checks.py` provides independent checks that run only against actually recalculated files. Every generated workbook is a **candidate requiring recalculation**. No production release or error-free claim is implemented.

Six user source files are catalogued in [SOURCES.md](SOURCES.md); the Marksans workbook passes the adapter gate, and an explicit-but-unapproved editable-input map proposal lives in [docs/editable-input-map.md](docs/editable-input-map.md). Run `python verify_sources.py` to restore and verify the original files after cloning.

## Run

Python 3.11 or later; no third-party runtime dependencies.

    python modelmaker.py inspect template.xlsx
    python modelmaker.py build --template template.xlsx --pack pack.json --inputs inputs.json --output candidate.xlsx
    python -m unittest -v

Build prints a JSON manifest containing template/input/output hashes, changed inputs, source references, and status. Save that manifest beside the workbook in cloud storage. Inspect is read-only. Build refuses to overwrite either the template or an existing output.

## Cloud source of truth

- Code, instructions, contracts, and reviewed model specifications: this GitHub repository.
- Original Excel files, courses, extracted source references, generated workbooks, and reports: versioned cloud object storage.
- Agent tasks, model-pack revisions, build runs, leases, and audit events: shared PostgreSQL database.
- Generation and recalculation: cloud workers operating on immutable snapshots.
- AI access: the same authenticated HTTPS API, OpenAPI contract, and MCP tools for every provider.

Agents can work simultaneously on separate branches, tasks, and build runs. Publishing a revision uses compare-and-swap checks; writing one workbook uses a lease. A binary XLSX is never merged by concurrent writers.

## Read next

- [Architecture and AI access](ARCHITECTURE.md)
- [Workbook quality policy](QUALITY.md)
- [Agent collaboration rules](AGENTS.md)
- [Roadmap and required inputs](ROADMAP.md)
- [Model-pack contract](model-pack.schema.json)
- [Input contract](inputs.example.json)
- [API contract](openapi.json)

## Deployment status

The core is a development foundation, not a deployed cloud app. Database/storage provisioning, HTTP/MCP implementations, Excel recalculation workers, and model/course onboarding remain open. Runtime cloud ownership requested: keshavmittal2203@gmail.com. Repository owner: keshavmittal2203-droid. No paid resources are required to run the core checks.
