# Workbook release policy

The goal is dependable models with visible evidence. No system can promise that an unspecified financial model is universally error-free. Every supported model family must have an approved template map and independently reviewed acceptance cases.

## Required release evidence

1. Exact source template and approved pack revision identified by SHA-256; original assets retained.
2. Complete factual input provenance; assumptions explicitly labeled; correct units, dates, fiscal periods, and sign conventions.
3. Only allowed cell edits. Formulas, named ranges, styles, merges, validations, charts, hidden sheets, and other native parts preserved or deliberately accounted for.
4. Supported calculation engine recalculated and saved the actual output; evidence binds engine/version, run ID, and output hash. Cached values or fullCalcOnLoad flags do not establish this.
5. Every formula cache checked for Excel errors; required formulas have calculated results. External data availability and iterative/circular calculation settings reviewed.
6. Model-specific independent checks, including reconciliations, timing, control totals, and valuation/sensitivity outputs, meet explicit tolerances. Checks must not simply repeat the same implementation.
7. Material input perturbation confirms dependent outputs move; restoration recovers the baseline. Each sensitivity point is independently verified where relevant.
8. Rendered substantive sheets reviewed against the template. Visual evidence is attached to the exact output version.
9. All mandatory checks passed and no unresolved blockers. A trusted validator publishes the report; an agent cannot self-assert a ready state through the public API.

## States

queued -> composing -> needs_recalculation -> validating -> ready

Any stage may produce blocked or failed. Ready publication belongs to the trusted validation service. A candidate remains needs_recalculation even when package/structural checks pass. The current development core has no ready transition.

## Initial composer support

Existing numeric input cells in ordinary, unprotected XLSX worksheets. The core rejects shared/array/data-table formulas, VBA, signatures, external workbook links, data connections, and query tables until a compatible workflow is implemented. A mapped input cannot be a formula cell or part of a merged range. Other parts are retained byte-for-byte inside the ZIP package. This is a supported subset, not a claim of general XLSX support.

The test suite uses synthetic fixtures to verify safeguards, exact package-part preservation, input bounds, stale-cache removal, and refusal to overwrite. It does not substitute for opening real workbooks in Excel and validating financial logic.
