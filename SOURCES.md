# Source assets

The Excel product consumes the active master template `originals/Marksans Keshav.xlsx` and the user-approved websites listed in [docs/LOCKED_SOURCE_ALLOWLIST.md](docs/LOCKED_SOURCE_ALLOWLIST.md). The workbook template is an immutable source asset; create outputs in new candidate directories.

`source-catalog.json` also tracks five PowerPoint files supplied earlier in the project history. They are retained solely as inert archival assets for byte-integrity and provenance. The Excel Model Maker does not inspect, build, export, or depend on them.

Run `python verify_sources.py` to verify the immutable source catalog. Source verification confirms file identity only; it does not establish financial correctness or approve a workbook.
