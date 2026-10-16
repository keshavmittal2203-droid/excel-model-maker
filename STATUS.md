# Current shared status

Updated: 2026-10-16, Asia/Calcutta.

## Codex parallel template-planning change

Branch `agent/codex-template-first`, based on main `eef33cf`; independent of Gemini/Argon PR #3. Added a dependency-free preflight planner for hash-pinned uploaded workbook/deck destinations, fictional request examples, template-first acceptance rules, and the user-supplied YouTube course reference. The planner does not create documents or approve packs. Universal Excel derivatives, slide assembly, course transcript/rule extraction and rendered acceptance remain open. No tests or real builds were run for this change; existing test counts below describe earlier work. No model training, cloud deployment or paid services performed.

## Implemented

- Dependency-free Python XLSX inspector and conservative existing-numeric-input composer.
- Styled report builder (`cli.py` + `reportmaker/`): strict JSON spec validation, colored bold headers, frozen panes, banded rows, number/currency/percent/date formats, estimated widths, SUM totals as real Excel formulas, conditional formatting, embedded chart; refuses to overwrite without `--force`; openpyxl pinned in requirements.txt.
- Template/approval hashes, exact input mapping, numeric bounds/precision checks, source references, original/output overwrite protection.
- Formula-cache invalidation, calculation-chain cleanup, recalculation flags, semantic worksheet-change guard, and unchanged native package-part preservation.
- Explicit needs_recalculation output. No ready/error-free release transition.
- 30 focused tests passed locally, including CLI build/overwrite behavior and generated-workbook round-trip checks on synthetic fixtures.
- Model-pack schema, draft OpenAPI contract, workbook quality policy, agent collaboration rules, and AI connection/build smoke-check prompt.

## Next steps

1. Run the independent AI smoke checks in AI_START_HERE.md. Use separate branches/PRs.
2. Onboard the six supplied files in SOURCES.md: 17-sheet workbook and five decks (182 slides). Originals are hashed; special workbook formulas block composition pending a compatibility adapter and approved input map.
3. Connect Supabase under keshavmittal2203@gmail.com and check project cost before provisioning. The currently connected Supabase organization belongs to the SRCC identity; it has no projects.
4. Implement/deploy shared file storage, task/job database, HTTP/MCP service, and compatible calculation worker.
5. Validate the first real model pack with actual recalculation, independent financial checks, perturbation/restoration, and rendered-sheet comparison.

## Current limits

No deployed API/MCP service, shared runtime database, calculation worker, or approved real model pack yet. The GitHub repo is the shared development source and initial original-file store. Other AI coding tools require authenticated access to this private repo. Concurrent development uses branches/PRs; cloud runtime concurrency contracts are documented but not implemented.
