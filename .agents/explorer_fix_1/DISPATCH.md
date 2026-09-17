# Task Assignment: Explorer Fix 1 (Integrity & GUI/Packaging Remediation)

## Identity
- Archetype: teamwork_preview_explorer
- Role: Integrity & GUI/Packaging Remediation Investigator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Context
Gate Iteration 1 FAILED due to Forensic Auditor INTEGRITY VIOLATION (hard veto).
You MUST read:
- `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` (authoritative user requirements)
- `D:\Sandbox\pm_sosanhbom\PROJECT.md`
- `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md`
- `D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md` (FULL FORENSIC AUDIT EVIDENCE REPORT)
- `D:\Sandbox\pm_sosanhbom\.agents\reviewer_1\handoff.md`
- `D:\Sandbox\pm_sosanhbom\.agents\reviewer_2\handoff.md`

## Objectives & Scope
Investigate and design exact remediation code changes for:
1. **Audit Finding 1.1 — Facade in `src/ui/main_window.py`**:
   - Inspect `src/ui/main_window.py` (lines 70-110).
   - Design removal of dummy dictionary and simulated progress. Wire `ReconciliationWorker` directly to `ReconciliationEngine`, `TC14AutomationClient`, `CS12Service`, and `ExcelReportGenerator`.
   - Ensure `SSBOM_Launcher.py`, `apps/1.0.0/SSBOM_App.py`, and `scripts/package_app.py` are properly directed to `src.gui.app:main` (the genuine PyQt6 application) or unified with `src/ui/`.
2. **Audit Finding 1.2 — Ten Self-Certifying Feature Tests in `tests/tier1_features/`**:
   - Inspect all 10 test files with 0 `src/` imports: `test_f07_missing_parts.py`, `test_f08_cross_station.py`, `test_f09_annotation_migration.py`, `test_f10_msi_decision.py`, `test_f22_leader_workspace.py`, `test_f23_member_workspace.py`, `test_f24_consolidated_report.py`, `test_f25_outlook_notification.py`, `test_f26_e2e_regression.py`, `test_f28_standalone_packaging.py`.
   - For each test file, map every test case to import and genuinely test the production class/function from `src.core`, `src.gui`, or `src.reporting`.
3. **Reviewer Finding F-02 & MAJ-01 — Launcher & Packaging**:
   - Inspect `SSBOM_Launcher.py` and `installer/SSBOM_Manager.iss`. Ensure real hash checking and clean packaging.

Write your comprehensive remediation report to `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1\handoff.md` and message parent orchestrator.

## 2026-09-17T06:30:20Z
You are Explorer Fix 1 (Integrity & GUI/Packaging Remediation Investigator).
Your working directory is: D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1
Original User Request is at: D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md
Project Index & Feature Inventory is at: D:\Sandbox\pm_sosanhbom\PROJECT.md
Dead Ends Log is at: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md
Full Forensic Audit Report is at: D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md
Reviewer Reports are at: D:\Sandbox\pm_sosanhbom\.agents\reviewer_1\handoff.md and D:\Sandbox\pm_sosanhbom\.agents\reviewer_2\handoff.md
Your task dispatch file is at: D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1\DISPATCH.md

MANDATORY: Read ORIGINAL_REQUEST.md first, then DEAD_ENDS.md and auditor_1/handoff.md.
Investigate and design complete, concrete remediation steps for:
1. Eliminating facade in src/ui/main_window.py: wire or replace with genuine implementation in src/gui/app.py and services (ReconciliationEngine, TC14AutomationClient, CS12Service). Ensure SSBOM_Launcher.py and packaging scripts launch genuine src.gui.app:main.
2. Rewiring all 10 self-certifying feature tests in tests/tier1_features/ (test_f07, f08, f09, f10, f22, f23, f24, f25, f26, f28) so they import from src/ and genuinely exercise the codebase.
3. Proper packaging and launcher checks.

Write your report to D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1\handoff.md and message parent orchestrator.

