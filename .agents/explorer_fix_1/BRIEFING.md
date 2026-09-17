# BRIEFING — 2026-09-17T06:30:20Z

## Mission
Investigate and design complete, concrete remediation steps for:
1. Eliminating facade in `src/ui/main_window.py`, wiring/replacing with genuine implementation in `src/gui/app.py` and services (`ReconciliationEngine`, `TC14AutomationClient`, `CS12Service`), and ensuring `SSBOM_Launcher.py` and packaging scripts launch genuine `src.gui.app:main`.
2. Rewiring all 10 self-certifying feature tests in `tests/tier1_features/` (`test_f07`, `f08`, `f09`, `f10`, `f22`, `f23`, `f24`, `f25`, `f26`, `f28`) to import from `src/` and genuinely exercise the codebase.
3. Proper packaging and launcher checks (MP2027 standard, HASH_ONLY_LAN, Inno Setup).

## 🔒 My Identity
- Archetype: teamwork_preview_explorer
- Roles: Integrity & GUI/Packaging Remediation Investigator
- Working directory: D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_1
- Original parent: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43
- Milestone: Remediation Planning & Architecture Design

## 🔒 Key Constraints
- Read-only investigation — do NOT modify production code directly in `src/`, `tests/`, etc. (only write to own `.agents/explorer_fix_1` folder).
- Design concrete, precise, machine-applicable code changes / patches.
- Binary veto standard: zero tolerance for cheating, facade mocks, or shortcuts.

## Current Parent
- Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43
- Updated: 2026-09-17T06:36:00Z

## Investigation State
- **Explored paths**:
  - `src/ui/main_window.py` (lines 1-574)
  - `src/gui/app.py`, `src/gui/leader_view.py`, `src/gui/member_view.py`, `src/gui/settings_dialog.py`
  - `src/core/reconciliation.py`, `src/core/msi_engine.py`, `src/reporting/excel_generator.py`, `src/reporting/outlook_mailer.py`
  - `SSBOM_Launcher.py`, `current.json`, `apps/1.0.0/SSBOM_App.py`, `scripts/package_app.py`, `packaging/build_exe.py`, `installer/SSBOM_Manager.iss`, `pm_sosanhbom.spec`
  - All 10 self-certifying tests in `tests/tier1_features/`: `test_f07`, `f08`, `f09`, `f10`, `f22`, `f23`, `f24`, `f25`, `f26`, `f28`.
- **Key findings**:
  1. **Facade in `src/ui/main_window.py`**:
     - `ReconciliationWorker` emits progress strings without calling `TC14AutomationClient` or `CS12Service` if files are missing, and previously returned hardcoded summary stats.
     - `src/gui/app.py` has authentic PyQt6 `SSBOMMainWindow` hosting `LeaderWorkspaceView` and `MemberWorkspaceView` with `--health-check` CLI.
     - `SSBOM_Launcher.py`, `apps/1.0.0/SSBOM_App.py`, and `scripts/package_app.py` mistakenly pointed to `src.ui.main_window:main`.
     - `src/gui/leader_view.py:100` has an invalid import: `from src.automation.sap.parser import SAPBOMParser` (class is `ResilientR3Parser` or function `parse_r3_cs12_file`).
  2. **10 Self-Certifying Test Files**:
     - AST audit confirmed all 10 files have 0 `src` imports.
     - Every file defines local reference functions or classes and asserts against itself.
     - All 10 files have exact, 1-to-1 counterparts in `src.core.reconciliation`, `src.core.msi_engine`, `src.gui.leader_view`, `src.gui.member_view`, `src.reporting.excel_generator`, `src.reporting.outlook_mailer`, and `scripts.package_app` / `SSBOM_Launcher`.
  3. **Launcher & Packaging**:
     - `SSBOM_Launcher.py` lacks manifest SHA256 integrity check and rollback to `previous.json`.
     - `installer/SSBOM_Manager.iss` references `SSBOM_Launcher.exe` at root, which is missing from root.
     - `apps/1.0.0/SSBOM_App.py` and `scripts/package_app.py` write `from src.ui.main_window import main`.

## Key Decisions Made
- Design turnkey replacement patches for all 10 test files.
- Redirect all launchers and entrypoints to `src.gui.app:main`.
- Cleanly rewire `src/ui/main_window.py` to genuine services and make its `main()` delegate to `src.gui.app:main`.
- Add `qapp` session fixture recommendation to `tests/conftest.py`.

## Artifact Index
- `BRIEFING.md` — persistent working memory
- `progress.md` — liveness heartbeat
- `handoff.md` — complete 5-component handoff report

