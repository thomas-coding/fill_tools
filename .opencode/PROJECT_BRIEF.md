# Project Brief

## Project Name
- Shengdan Tool V0.2 (Windows desktop automation for Excel reporting)

## Long-Term Goal
- Deliver a stable Windows desktop automation tool ("Shengdan Tool V0.2") for Excel reporting, with reliable `F5/F12/F10` workflow execution for non-developer users.
- Keep write-back safe and controlled: only `F10` writes to Excel, while workbook structure and embedded images stay intact.

## In Scope
- Patrol parsing and data prep for both fill profiles (`patrol1`, `patrol2`).
- Stable `F5/F12/F10` hotkey workflow and profile-specific field mapping.
- Multi-path image extraction compatibility (anchor image, WPS `DISPIMG/cellimages`, zip drawing fallback).
- Safe progress/time write-back with profile-isolated status columns.
- Release quality gates (`release_preflight.py`) and manual verification checklist (`RELEASE_CHECKLIST.md`).

## Out of Scope
- Mini-program account/login ownership and credential troubleshooting.
- Any workbook schema mutation beyond designated status/time write-back columns.
- Cloud/backend workflow changes; tool remains local desktop automation.

## Constraints
- Tech and style constraints:
  - Windows-first desktop app with GUI flow: select file, run/stop control, progress, and logs.
  - Keep AHK hotkey flow stable; `F5/F12/F10` behavior must be deterministic in repeated runs.
  - Template mapping follows `1.xlsx` structure (label row + content row).
  - Input formats must include `xlsx/xlsm/xltx/xltm/xls/xlsb/.excel`.
  - Prefer COM-based write-back in inspection mode to reduce embedded-image loss risk.
  - Keep patrol semantics stable:
    - `patrol1`: `address <- road`, `section <- location`, `description <- issue`.
    - `patrol2`: `section <- road`, `description <- rectify (fallback issue)`.
- Release quality constraints:
  - Before any customer release, `python release_preflight.py` MUST pass.
  - `build_exe.cmd` must keep preflight gate enabled; do not bypass release checks.
  - Manual verification evidence is mandatory per `RELEASE_CHECKLIST.md`.
- Must not do:
  - Never write back during `F5` or `F12`; write-back is allowed only in `F10`.
  - Do not alter non-target workbook content, break embedded images, or corrupt file structure.
  - Do not leave file-handle locks after close or require developer-only setup on target machines.

## Important Paths
- `app_engine.py`: parsing, session building, profile mapping, AHK runtime script generation.
- `tests/test_app_engine.py`: parser/write-back regression tests.
- `offline_smoke_check.py`: offline field and image smoke checks.
- `release_preflight.py`: release gate runner (unit tests + smoke check).
- `build_exe.cmd`: packaging entrypoint with mandatory preflight gate.
- `RELEASE_CHECKLIST.md`: required manual verification and release evidence template.
- `.opencode/PROJECT_BRIEF.md` and `.opencode/SESSION_STATE.md`: persistent context for new sessions.

## Milestone and Done Definition
- Current milestone:
  - Complete non-development-machine end-to-end acceptance (`select file -> F5/F12/F10 -> close -> reopen verification`).
- Done definition:
  - End-to-end flow passes on at least one non-development machine without manual workaround.
  - Data is written correctly after `F10`, and no write-back occurs during `F5/F12`.
  - Embedded images remain intact and workbook reopens normally after write-back.
  - File-dialog compatibility works across common window/control variants.
  - Packaged executable smoke test passes with no persistent file-lock issue.

## Collaboration Notes
- Keep responses concise.
- Before implementation or release work in a new session, read `.opencode/PROJECT_BRIEF.md`, `.opencode/SESSION_STATE.md`, and `RELEASE_CHECKLIST.md`.
