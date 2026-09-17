# Task Assignment: Worker Fix GUI & Feature Tests

## Identity
- Archetype: teamwork_preview_worker
- Role: GUI & Feature Tests Remediation Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_fix_gui_tests
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Context & Inputs
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\PROJECT.md`.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md`.
- Read Explorer Fix 1 report at `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1\handoff.md`.
- Read Explorer Fix 3 report at `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_3\handoff.md`.
- Read Auditor 1 report at `D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md`.

## Exclusive File Ownership
You exclusively own and must edit/create:
1. `src/ui/main_window.py`:
   - Eliminate facade: remove hardcoded dictionary and simulated progress.
   - Wire `ReconciliationWorker` directly to genuine services: `ReconciliationEngine`, `ExcelReportGenerator`, `OutlookMailer`, and real data structures.
2. `src/gui/leader_view.py`:
   - Fix import at line 100: replace non-existent `SAPBOMParser` with `parse_r3_cs12_file` / `ResilientR3Parser`.
   - Add QThread running guard to `trigger_batch_reconciliation` (`if self.thread and self.thread.isRunning(): return`) and connect `deleteLater` lifecycle hooks.
3. `src/gui/member_view.py`:
   - Fix import at line 436: replace `SAPBOMParser` with `parse_r3_cs12_file` / `ResilientR3Parser`.
4. `SSBOM_Launcher.py` & `apps/1.0.0/SSBOM_App.py` & `scripts/package_app.py`:
   - Redirect entry point to `src.gui.app:main` (the genuine Leader/Member workspace application).
   - Implement manifest SHA-256 integrity check and rollback via `previous.json` in `SSBOM_Launcher.py` per MP2027 standard.
5. Rewire all 10 self-certifying feature tests in `tests/tier1_features/` so that every test imports and exercises genuine `src/` classes:
   - `test_f07_missing_parts.py`: import and test `src.core.reconciliation.detect_missing_plm_parts`
   - `test_f08_cross_station.py`: import and test `src.core.reconciliation.aggregate_cross_station_totals`
   - `test_f09_annotation_migration.py`: import and test `src.core.reconciliation.migrate_annotations`
   - `test_f10_msi_decision.py`: import and test `src.core.msi_engine.MSIEngine` / `evaluate_msi_branch`
   - `test_f22_leader_workspace.py`: import and test `src.gui.leader_view.LeaderWorkspaceView`
   - `test_f23_member_workspace.py`: import and test `src.gui.member_view.MemberWorkspaceView`
   - `test_f24_consolidated_report.py`: import and test `src.reporting.excel_generator.ExcelReportGenerator`
   - `test_f25_outlook_notification.py`: import and test `src.reporting.outlook_mailer.OutlookMailer`
   - `test_f26_e2e_regression.py`: import and test `src.core.reconciliation.ReconciliationEngine`
   - `test_f28_standalone_packaging.py`: import and test `SSBOM_Launcher` and `packaging/build_exe.py`

Run tests: `pytest tests/tier1_features/ -v`.
Write report to `D:\Sandbox\pm_sosanhbom\.agents\worker_fix_gui_tests\handoff.md` and message parent orchestrator.
