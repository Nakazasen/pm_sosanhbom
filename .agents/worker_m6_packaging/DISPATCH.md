# Task Assignment: Milestone M6 Worker (E2E Tier 5 Adversarial Hardening & Standalone Packaging)

## Identity
- Archetype: teamwork_preview_worker
- Role: Milestone M6 Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_m6_packaging
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Scope & File Ownership
You exclusively own and must implement:
- Resolve the defect noted by Test Writer: in `src/automation/sap/models.py`, ensure `R3ComponentRow.rev_r3` coerces inputs to `str` cleanly (e.g. converting int `1` -> `"01"` or `str(v)`).
- `packaging/build_exe.py` and `pm_sosanhbom.spec`:
  - F28: Standalone executable (.exe) packaging script using PyInstaller.
  - Bundles all necessary modules (`src/core`, `src/automation`, `src/gui`, `src/reporting`), dependencies, and assets.
  - Entry point: `src/gui/app.py`.
  - Configures hidden imports (openpyxl, win32com, PyQt6, selenium, etc.).
- `tests/tier5_adversarial/`:
  - F27: Tier 5 adversarial coverage hardening test cases.
  - Malformed Excel files, corrupted BOM headers, simultaneous timeout and network loss, extreme numerical quantities, memory limits, and invalid Unicode.
- Run `pytest tests/` and verify 100% passing across all 5 tiers.
- Run PyInstaller build check (or dry-run/validation) to verify packaging readiness.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md`.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read `TEST_READY.md` and `TEST_INFRA.md`.

## Deliverables
- Implement packaging scripts and Tier 5 tests.
- Run tests and packaging verification.
- Write completion report to `D:\Sandbox\pm_sosanhbom\.agents\worker_m6_packaging\handoff.md`.


## 2026-09-17T03:59:54Z
You are Worker M6 (Milestone M6: PyInstaller Packaging & Tier 5 Adversarial Hardening).
Your working directory is: D:\Sandbox\pm_sosanhbom\.agents\worker_m6_packaging
Original User Request is at: D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md
Project Index & Feature Inventory is at: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md
Your task dispatch file is at: D:\Sandbox\pm_sosanhbom\.agents\worker_m6_packaging\DISPATCH.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

You exclusively own and must implement:
1. Defect resolution: In src/automation/sap/models.py, ensure R3ComponentRow.rev_r3 coerces integer or float revision numbers (e.g. 1, 2) cleanly to string ("01", "02", or str(v)).
2. packaging/build_exe.py and pm_sosanhbom.spec:
   - F28: Standalone executable (.exe) packaging script using PyInstaller.
   - Bundles all modules: src.core, src.automation, src.gui, src.reporting, assets, openpyxl, win32com, PyQt6.
   - Entry point: src/gui/app.py.
   - Includes verification mode: test pyinstaller build or spec validation.
3. tests/tier5_adversarial/:
   - F27: Tier 5 adversarial coverage hardening test cases.
   - Malformed workbooks, corrupted BOM headers, sudden network drops, circular BOM structures, extreme numbers, invalid dates, and invalid Unicode.
4. Run full pytest suite across all tests: pytest tests/ -v.
5. Write your completion report to D:\Sandbox\pm_sosanhbom\.agents\worker_m6_packaging\handoff.md with test commands and results.
6. Send message to parent orchestrator when complete.
