# Shared cloud architecture

## Ownership and access

GitHub account: keshavmittal2203-droid. Runtime cloud account requested by the user: keshavmittal2203@gmail.com. Initial platform choice is Supabase for PostgreSQL and object storage, with a separate container/Excel-compatible worker. The currently connected Supabase organization belongs to the SRCC account, so provisioning must use the requested Gmail identity. No runtime resources have been provisioned.

All AI providers use the same records and artifacts. Changing the model does not change the source of truth. A client needs either GitHub access for development or the HTTP/MCP tool capability and workspace credentials for product use.

## Flow

    Excel models + course documents -> immutable cloud asset versions
        -> read-only extraction + cited course rules
        -> reviewed model pack + exact template hash
        -> supplied facts and explicit assumptions
        -> deterministic template composer
        -> calculation worker
        -> independent validation + render review
        -> versioned workbook and evidence report

The LLM proposes mappings, identifies assumptions, and explains model decisions. The composer enforces the approved map. The calculation engine computes formulas. Independent checks decide whether the output can be released.

## Shared records

- Asset: workspace ID, immutable object key, SHA-256, MIME type, size, source, version, upload time.
- Model pack: revision, template asset hash, allowed inputs, units, periods, signs, business rules, checks, source citations, reviewer approval.
- Build run: pinned pack/asset versions, input digest, idempotency key, state, output/evidence hashes, lease epoch, timestamps.
- Task: owner/agent, branch/PR, claimed files, status, expected revision, acceptance criteria.
- Event: actor, mutation, object version, timestamp, correlation ID; append-only.

## Simultaneous work

Use transactionally allocated revisions, unique idempotency keys, and atomic job claims. Workers claim queued rows with database locking; each claim increments a fencing epoch and sets an expiry. Heartbeats and completion updates must match that epoch. A reclaimed job rejects the old worker even if it later finishes. Write outputs to immutable keys, then atomically publish their hashes. Retry uploads can leave unreferenced objects; a separate cleanup process may remove these after retention.

Agents can create parallel candidate pack revisions. Only a successful compare-and-swap against the published revision changes the shared current pointer. Workbook bytes are immutable artifacts, so independent generation runs cannot overwrite one another.

## API and MCP

The planned API exposes asset upload/finalization, model-pack revisions, generation runs, status, evidence, and task claims. openapi.json is the initial contract and does not imply a running service. MCP tools will call the same application service and authorization rules: list_templates, inspect_asset, get_model_pack, create_build, get_build, get_validation_report, claim_task, and complete_task.

Provider credentials stay in server-side adapters. DeepSeek, Gemini, OpenAI, and other model providers can be added independently of workbook storage and calculation. Do not assume all chat interfaces support MCP or arbitrary HTTP calls; clients that lack those capabilities need a compatible host or coding environment.

## Calculation compatibility

Select the calculation worker after inspecting the real templates. A native Excel-compatible worker must prove support for every used function, data table, link, macro, chart, and other native feature. A library that writes formulas does not calculate them. The current core emits candidates and cannot mark them ready. Licensed Excel automation and headless alternatives need separate deployment and compatibility decisions.
