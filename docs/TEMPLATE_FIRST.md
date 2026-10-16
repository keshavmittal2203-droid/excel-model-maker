# Template-first delivery contract

## User's requested product

Make Excel models in the style and structure of the supplied workbook. Make presentations using the supplied decks as the only design sources, placing new company information into their existing layouts. The generic styled report builder is not the default route for either request.

This is a reusable template and rules system, not model-weight training. All AI clients must follow the same contracts, source references and acceptance examples.

## Excel modes

1. **Exact template (default):** start from an immutable, hash-verified copy of `originals/Marksans Keshav.xlsx`. Retain sheet order/names, formulas, formatting, print areas, hidden sheets, charts, named ranges, validations and native features. Change only reviewed mapped inputs. Current composer supports numeric inputs; company names and fiscal-year label updates require a separately reviewed text-input adapter. Do not deliver a model with a new company's figures beneath old company labels.
2. **Universal derivative (requested future capability):** derive a reusable model family from the uploaded workbook and course, preserving its modelling methodology and recognisable schedule structure. Document every structural change, separate historical inputs from forecast assumptions, and version the derivative as a new template. Pharma, project businesses and financial firms need different operating schedules; unsupported families block rather than silently using pharma assumptions. This mode is not yet implemented by the planner.

Default output is a candidate copy plus provenance and validation evidence. The uploaded Marksans workbook is a reference, not proof that its financial formulas are all correct.

## PowerPoint modes

Default: select **one of the five uploaded decks** and fill approved existing text/table/chart destinations. Existing themes, masters, backgrounds, logos, placement, typography, shapes and relationships are the design source. No external templates or arbitrary new slide designs.

Plan each slide before composition: narrative purpose, original deck SHA and source slide number, shape/table/chart target, new content, evidence, numeric workbook binding, and content budget. A useful research narrative normally covers thesis, company/business, industry, operating drivers, financials, valuation, catalysts, risks and sources; select matching uploaded layouts rather than fabricating a new layout for each heading.

Reusing/reordering/duplicating slides from uploaded decks is allowed in the intended product, **after a dedicated assembly adapter** preserves layout/master/media/chart dependencies and records source slide lineage. The current planner supports existing text shapes in a single unchanged deck; it does not implement slide assembly, chart replacement or table-cell planning. Do not imply otherwise.

Every retained slide must be reviewed for old company names, figures, dates, logos, footnotes, hidden content, speaker notes and embedded chart workbooks. Unmapped old content is not automatically relevant. Replace it through reviewed slots, explicitly retain it as relevant, or block the affected slide. Never squeeze text by arbitrarily shrinking fonts. Edit the content or use another approved uploaded layout; confirm fit by rendering.

## Runnable preflight

`python template_plan.py --request template-workflow/excel-request.example.json`

`python template_plan.py --request template-workflow/ppt-request.example.json`

Both examples are labelled fictional planning data. The planner validates the catalog hash, original bytes, existing destinations, duplicate mappings, basic provenance fields and PPT text budgets. It prints a reviewable plan to stdout and does not modify originals or create documents. Exit 2 means blocked. Source strings are not independently verified; character limits do not prove visual fit. It cannot replace the engine's input protections or QUALITY.md.

## End-to-end acceptance

- Each factual value has a precise source and location/date; assumptions have explicit rationale and bounds. Preserve units, periods, currencies and sign conventions.
- Build using reviewed mappings; preserve template bytes outside permitted changes. Numeric deck claims bind to the recalculated workbook hash, sheet, cell, scale and rounding.
- Recalculate the actual Excel output in a compatible engine. Independently reconcile accounting identities, valuation, and meaningful perturbation/restoration. Unsupported formulas block, never become zero or a copied cached result.
- Render substantive sheets and every changed slide. Inspect clipping, overflow, stale content, chart labels, footnotes, alignment and consistent templates. Package preservation alone is not visual verification.
- Record exact source/output hashes, engine/version, commands, actual checks and missing capabilities. Never set `ready: true` merely because a ZIP parses or a file exports.
- Changes go through a separate branch and reviewed PR. Gemini's engine work in PR #3 and Codex's planning work remain separate until integration review.

## Integration task for Gemini Flash

Read this file and COURSE_REFERENCE.md. Preserve current main's report builder, but route financial-model requests to the template composer. Wire the planning contract to reviewed model/deck packs; add exact-template text labels safely, then bounded slide reuse with lineage, numeric bindings and rendered overflow checks. Fix PR #3's unconditional PPT `ready: true`. Do not treat its built-in partial Python evaluator as Excel-equivalent. Add appropriate regression coverage in your implementation PR, report exact results and unmet gates, and leave main and originals unchanged. A universal derivative is a separate model-family onboarding task, never an automatic fallback.
