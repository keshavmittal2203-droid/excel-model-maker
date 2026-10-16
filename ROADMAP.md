# Delivery roadmap

## User priority: template fidelity and reusable methodology

1. Exact uploaded Excel template and uploaded PPT layouts are the default. Use docs/TEMPLATE_FIRST.md and template_plan.py to propose destinations before engine composition.
2. Safely support company/date text labels, complete stale-company-content review, hash-bound workbook-to-deck values, and rendered overflow checks.
3. Implement lineage-preserving reuse/reordering of uploaded slides; no external deck designs.
4. Review and cite actual lessons from the user's [modelling course](docs/COURSE_REFERENCE.md); derive schedule rules and independent worked acceptance examples.
5. Onboard explicitly versioned universal Excel model families derived from those references; keep unsupported business types blocked. A universal template is not implemented yet.

## Current slice

- Development core: inspect templates, validate reviewed input maps, fill mapped existing numeric cells, preserve package parts, remove stale caches, request recalculation, emit provenance manifest.
- Contracts: model-pack JSON Schema, input example, HTTP API draft, cross-agent rules, workbook release policy.
- Verification: focused Python unit tests. No real user template or course onboarded yet.

## Next: first real model pack

Provide one exact template, one completed example using that template, and the relevant course materials. Include one task the model maker should perform and the expected output. The original workbook should retain its formulas and native features. Explain which cells/assumptions can change and which layout/features must remain exact.

Onboarding will inventory sheets, formulas, dependencies, units, timelines, named ranges, charts, tables, external links, macros, and data tables. Course rules receive page/section citations. Ambiguities become explicit decisions; missing facts are not invented. The first pack is reviewed against the completed example before generation is enabled for that model family.

## Cloud workspace

1. Connect the requested Gmail cloud account and confirm the selected project's cost.
2. Provision versioned asset storage and the PostgreSQL job/task ledger.
3. Implement the HTTP contract and authentication, then an MCP adapter over the same service.
4. Add transaction/lease concurrency tests and multi-client integration tests.
5. Deploy the calculation worker after template compatibility is known.
6. Add independent validation, render comparison, and a generation/download interface.

## Subsequent scope

Multiple model families, reusable schedule/rule library, historical-data extraction, source refresh, scenarios, sensitivity runs, and provider-specific reasoning adapters. Each supported family gets its own reviewed acceptance suite. Broader formula generation follows the first validated template-based workflow.
