# Connect and build without overwriting shared work

Repository: https://github.com/keshavmittal2203-droid/excel-model-maker

The repository is private. Connect each coding assistant's GitHub integration to an account with access to this repo, or use its authenticated coding environment. A plain chat model cannot gain write access from a pasted link. Do not paste tokens into chats. One assistant may be Gemini and another OpenAI; they use the same Git repo and separate branches.

## Copy this into each coding assistant

    Work in https://github.com/keshavmittal2203-droid/excel-model-maker.
    First prove that you can access it: report the default-branch commit SHA
    and list the root files. Read AGENTS.md, README.md, QUALITY.md, and STATUS.md.

    Perform a connection/build smoke check using a unique branch named
    agent/<your-client-name>-smoke-<unique-suffix> from the current main.
    Run python -m unittest -v using Python 3.11 or later.
    Report the exact test count, pass/fail result, and Python version.
    Tests include a synthetic workbook build; no real course/model is needed.

    Add exactly one new file named AGENT_CHECK_<YOUR_UNIQUE_NAME>.md containing
    the starting commit SHA, commands run, actual results, and any blockers.
    Do not edit or delete existing project files during this smoke check.
    Commit that new file only on your branch and open a pull request to main.
    Do not merge the PR, force-push, change repository settings, provision paid
    resources, or modify the default branch. Give me the PR URL.

    If GitHub access, a Python runtime, or PR creation is unavailable, say which
    capability is missing. Do not claim a command ran or a PR exists without
    tool evidence. After the smoke check, wait for a separately assigned task.

Use a different client name/suffix in each assistant. Both checks may run at the same time because each creates its own file on its own branch. Review both PRs before merging. Git history preserves the starting version for recovery. Repository rulesets should enforce the PR/test requirement when supported by the account plan; instructions alone do not enforce permissions.

## After connectivity succeeds

Assign independent issues with different file scopes. An agent adds acceptance tests for its change and submits a PR. The shared product API, object store, job leases, and automatic workbook release service remain planned; they are not yet deployed. See STATUS.md before assigning runtime integration work.

## End-to-End Operation & Clean Checkout Guide

For comprehensive instructions on building packages, preparing canonical requests, selecting template layouts, understanding gate semantics, and running regression checks, see [docs/AI_CLEAN_CHECKOUT_GUIDE.md](docs/AI_CLEAN_CHECKOUT_GUIDE.md).

## Template V16 Clean-Checkout Model Reproduction

To reproduce all four company models (MM, LTM, CLEAN, EDELWEISS) from committed source inputs in a fresh or clean checkout:

1. **Verify Master Template Integrity**:
   ```powershell
   python verify_sources.py
   ```
   Both Excel master files must report `VERIFIED` byte-for-byte (`originals/Marksans Keshav.xlsx` SHA-256: `4dc79842855e5d1e689a2f5ca1de22c34be72fb09c0c0f9dd3a95818e10737d0`). Never modify files in `originals/`.

2. **Execute Single-Command Reproduction**:
   ```powershell
   python work/build_template_v16_models.py --output-dir staging/v16_reproducibility
   # OR:
   python build_package.py --template-v16 --output-dir staging/v16_reproducibility
   ```
   In headless or non-Windows environments without Microsoft Excel COM, add `--skip-com`.

3. **Verify Results**:
   - Exactly **0 formula errors** across all 64 sheets.
   - Dynamic `BS!C40:N40 = ={col}38-{col}16` in all 12 periods.
   - Authentic balance check results exposed without plugs:
     - MM: `+₹7,677.49 Cr` (`BS!N40`), forecast **BLOCKED**, valuation **Unavailable**.
     - LTM: `+₹1.00 Cr` (`BS!N40`), **UNRECONCILED**, valuation supported at **₹2,620.24**.
     - CLEAN: `−₹1.13 Cr` (`BS!N40`), **UNRECONCILED**, valuation supported at **₹393.03**.
     - EDELWEISS: `+₹1.00 Cr` (`BS!N40`), **UNRECONCILED**, valuation **Unavailable**.
   - 15 landscape PDF previews generated per target (60 total) in `staging/v16_reproducibility/<TARGET>/previews/`.
   - Governance: `ready: false`, `human_pack_approval: pending`. PR #3 remains open and unmerged.
