# Task Assignment: Forensic Integrity Auditor — Gate Iteration 2

## Identity
- Archetype: teamwork_preview_auditor
- Role: Forensic Integrity Auditor (Gate 2)
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\auditor_gate2_1
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

MANDATORY AUDIT ENFORCEMENT:
Your audit is a BINARY VETO. Zero tolerance for cheating, facade mocks, or shortcuts.

## Context
In Gate Iteration 1, you rejected the work product with `INTEGRITY VIOLATION` due to:
1. Facade in `src/ui/main_window.py` returning hardcoded dummy metrics.
2. 10 self-certifying feature tests in `tests/tier1_features/` with 0 `src/` imports.
3. Broken TC14 `client.py:login()` and depth limit in `models.py:38`.

Workers have completed extensive remediations:
- `src/ui/main_window.py`: facade eliminated, wired directly to genuine services, redirected `main()` to `src.gui.app:main`.
- `SSBOM_Launcher.py` & `apps/1.0.0/SSBOM_App.py`: redirected to `src.gui.app:main`.
- `tests/tier1_features/`: all 10 feature test files rewired with genuine `src/` imports.
- `src/automation/tc14/client.py`: `timeout` and `wait` fixed.
- `src/core/models.py`: `level <= 20`.

## Objectives
1. Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
2. Read `D:\Sandbox\pm_sosanhbom\PROJECT.md`.
3. Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md`.
4. Read your previous audit report at `D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md`.
5. Perform an exhaustive forensic integrity re-audit on `pm_sosanhbom`:
   - Static analysis & AST inspection across `src/` and `tests/`:
     * Verify `src/ui/main_window.py` contains NO hardcoded metrics or dummy shortcuts.
     * Verify all 10 previously failing feature tests in `tests/tier1_features/` now import from `src/` (AST check).
     * Verify no other dummy facades, bypasses, or fake attestation artifacts exist.
   - Runtime execution verification:
     * Run the full test suite (`pytest tests/ -v`). Verify 0 failures.
     * Execute `dist/SSBOM_Portable/SSBOM_Portable.exe --health-check`. Verify exit code 0.
6. State your explicit forensic verdict: `CLEAN` or `INTEGRITY VIOLATION`.
7. Write your audit report to `D:\Sandbox\pm_sosanhbom\.agents\auditor_gate2_1\handoff.md` and message parent orchestrator.

## 2026-09-17T06:54:02Z
User Request for Gate Iteration 2 received.
- Verify src/ui/main_window.py contains NO hardcoded metrics or dummy shortcuts.
- Verify all 10 previously failing feature tests in tests/tier1_features/ now import from src/ (AST check).
- Verify no other dummy facades, bypasses, or fake attestation artifacts exist.
- Run full test suite: pytest tests/ -v. Verify all tests pass cleanly with 0 failures.
- Execute dist/SSBOM_Portable/SSBOM_Portable.exe --health-check. Verify exit code 0.
- State explicit forensic verdict: CLEAN or INTEGRITY VIOLATION.
- Write audit report to handoff.md and message parent orchestrator.
