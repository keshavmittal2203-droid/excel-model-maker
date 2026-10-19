# Delivery roadmap

## Current slice

- The active supported example is the user's 16-sheet Marksans master adapted for M&M in V17. The template remains byte-verified and the workbook is rebuilt through native Excel COM on Windows.
- The current M&M candidate has 0 formula errors, 0 circular references, and 11 historical balance ties. Forecast and fair value are blocked where the approved source set lacks required facts.
- Current acceptance work is limited to a single-company Excel workbook. Other company deliverables are archived; PowerPoint is retired.
- Keep source mappings, workbook formulas, visual previews, and regression checks synchronized before publishing another candidate.

## Next: finish M&M acceptance

Complete the current visual and source-provenance cleanup, regenerate the M&M handoff workbook and archive, run the full repository verification, and publish only after the report and artifacts match the actual workbook. Keep `ready: false` and human approval pending.

## Cloud workspace

1. Connect the requested Gmail cloud account and confirm the selected project's cost.
2. Provision versioned asset storage and the PostgreSQL job/task ledger.
3. Implement the HTTP contract and authentication, then an MCP adapter over the same service.
4. Add transaction/lease concurrency tests and multi-client integration tests.
5. Deploy the calculation worker after template compatibility is known.
6. Add independent validation, render comparison, and a generation/download interface.

## Subsequent scope

Multiple model families, reusable schedule/rule library, historical-data extraction, source refresh, scenarios, sensitivity runs, and provider-specific reasoning adapters. Each supported family gets its own reviewed acceptance suite. Broader formula generation follows the first validated template-based workflow.
