# Critical package safety review — 2026-10-16

Scope: user requested live progress review and critical fixes alongside Gemini. Reviewed pushed Gemini commit `8cad0cfc8da0d1461516fc90ee5076f29942e4f0`. Task claim: [issue #6](https://github.com/keshavmittal2203-droid/excel-model-maker/issues/6). Patch branch: `agent/codex-critical-fixes`, separate from the active Gemini branch.

## Defects and fixes

1. **Output path escape / source deletion.** Request `output_name` was joined to the run directory without checking for absolute paths or traversal; the workbook destination was unlinked if it existed. A request could select an original or unrelated file. Validate both names before writes or Office activation: new plain filenames only, expected extensions, no separators/drive syntax/reserved device names/trailing spaces or dots, no existing files/symlinks. Remove unlink and use exclusive PPT creation.
2. **User Office session termination.** Capability probing used `Dispatch` followed by `Quit`, which can attach to an existing user application. Probe registration only, label that limitation explicitly, and leave version unknown rather than claiming an activated engine. Excel preview export uses `DispatchEx`, disables link updates/macros and quits only that application. PowerPoint may share a process even with `DispatchEx`: close only the presentation opened by the exporter, never call application `Quit`. This can leave an empty PowerPoint process; preserving user presentations takes precedence.
3. **Stale manifests on failed forced runs.** Nonempty run reuse could retain the previous successful-looking manifest/previews after a later stage failed. Reject reuse even with `--force`; use a new run directory. This is a deliberate behavior change, not an atomic overwrite implementation. Partial artifacts from fresh failed runs can still exist, but no previous package manifest is carried forward.
4. **Cross-run financial image contamination.** Orchestrator generated images both under the run and in shared `work/assets`. Remove shared writes; existing run-local replacement paths remain. Mocked orchestration verifies the generator is called once for that run.

## Actual checks

- Eight new tests in `test_package_safety.py`: traversal/absolute/Windows device/alternate stream names, existing output preservation, stale-manifest prevention, nonactivating capability probes, owned PPT closure, COM activation cleanup, private Excel export/link settings and run-local images. Only temporary fixtures and mocked COM used.
- Selected regression run: **65 tests passed**, covering the new eight, modelmaker, pptmaker, reportmaker and six filesystem/dimension/hash build-package tests.
- First run: 64 passed, one failed because a fresh checkout had only three reconstructed deck files. Ran `python verify_sources.py`; all six hashes matched and the two chunked decks were reconstructed. The same 65-test selection then passed.
- No actual Office activation/recalculation/rendering by this review. Deliberately excluded native build-package end-to-end tests while Gemini is active. This is not a claim the full native integration suite passed.
- `git diff --check`: no errors.

## Remaining review items

This is a bounded safety patch. It does not complete request-schema enforcement, replace every hardcoded Marksans mapping, independently verify financial assumptions, establish human approval, or prove native Office behavior on this machine. Gemini should integrate/review the patch, run native end-to-end builds in fresh run directories, and preserve pending release gates. COM registration is only a preflight hint; actual activation can still fail and must remain a build error.
