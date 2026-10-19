# Mahindra & Mahindra — active Template V17 model

The sole active M&M candidate is `template_v17/MM-FullTemplate.xlsx`. It is built from the unchanged 16-sheet Marksans master and the V17 builder, with historical input cutoffs no later than 15 October 2025.

## Current verified state

- Excel COM 16.0 full recalculation; 0 formula errors and 0 circular references.
- The 11 historical BS reported-total checks tie. Source-rounded component subtotals can differ by INR 1 crore; the difference is displayed without a plug.
- FY26E is blocked because approved sources do not provide the consolidated MMFS funding bridge.
- WACC and intrinsic fair value are unavailable. ROIC and intrinsic growth are unavailable because Screener does not split the operating asset base from aggregate Other Assets/MMFS lending balances.
- No cutoff-dated diluted share count or consolidated CapEx detail is present in the approved snapshots; the old share-count and CapEx claims have been withdrawn. The Comps sheet shows only supported historical inputs and a formula-derived P/E cross-check.
- Source policy: only websites named in the original master plus Screener. Active V17 provenance is in `SOURCE_REGISTRY.md` and `source_registry.json`.

## Legacy files

The pre-V17 `mm-package.json`, `mm.inputs.json`, and `mm.model-pack.json` files are archived in `archive/v16/`; they are not inputs to the active V17 builder and must not be treated as verified V17 evidence. The old source-coverage report and V8 validation report have been replaced with superseded notices. Generated PowerPoint previews are retired; original presentation templates remain inert archives only.

Build and review: run `python work/build_template_v17_mm.py --clean` from the repository root. The build keeps `ready: false` and human approval pending.
