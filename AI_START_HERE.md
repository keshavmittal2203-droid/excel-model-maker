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
