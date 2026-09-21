# BRIEFING — 2026-09-21T08:47:00+07:00

## Mission
Survey the entire PyQt6 GUI codebase of pm_sosanhbom at src/gui/ (app.py, leader_view.py, member_view.py, plm_download_dialog.py, settings_dialog.py, update_dialog.py) and analyze architectural deltas to implement the Data-Dense Enterprise Dashboard specification (SPEC_UI_UX_ENTERPRISE_DASHBOARD.md) without regressions.

## 🔒 My Identity
- Archetype: teamwork_preview_explorer
- Roles: Legacy VBA & Algorithm Investigator
- Working directory: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1
- Original parent: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43
- Milestone: legacy_vba_algorithm_investigation
- Current Milestone: gui_codebase_survey_and_data_dense_analysis
- Current Role: Explorer 1 (Survey Codebase GUI)
- Current Parent ID: 6014734f-cacb-4480-97ab-1fc3957409fb

## 🔒 Key Constraints
- Read-only investigation — do NOT implement production source code changes
- Write only to our own agent folder (.agents/teamwork_preview_explorer_survey_1)
- Communicate results via handoff.md and send_message to parent
- Read-only investigation of src/gui/ and UI specs — do NOT modify production files
- Survey widget structures, table metrics, business logic couplings, and emoji removal risks
- Ensure 100% backward compatibility with existing test suites (tests/tier1_features/, tests/unit/)

## Current Parent
- Conversation ID: 6014734f-cacb-4480-97ab-1fc3957409fb
- Updated: 2026-09-21T08:47:00+07:00

## Investigation State
- **Explored paths**:
  - `src/gui/app.py`: Main QMainWindow container, dual-role tabs, menus, toolbar, status bar indicators, logging dock.
  - `src/gui/leader_view.py`: 4-step sequential wizard (`WizardStepHeader`, `Step1ProjectSetupWidget`, `Step2DataSourcingWidget`, `Step3TrackingConsolidationWidget`, `Step4ComparisonReportingWidget`, `LeaderWorkspaceView`).
  - `src/gui/member_view.py`: 3-tab workbook editor (CTTT, MSI, Labels), assignment package auto-discovery, self-check engine, Q2="OK" submission seal.
  - `src/gui/plm_download_dialog.py`: Non-tech 3-step download dialog for TC24 and SAP R3, multiline part code input, progress bar, Consolas log box.
  - `src/gui/settings_dialog.py`: 3-tab system settings (TC24, SAP, Paths), persistence in config/settings.json.
  - `src/gui/update_dialog.py`: LAN HASH_ONLY_LAN update delivery dialog.
  - `specs/SPEC_UI_UX_ENTERPRISE_DASHBOARD.md`: Target specification for Data-Dense Enterprise Dashboard.
  - `tests/tier1_features/test_f22_leader_workspace.py` & `test_f23_member_workspace.py`: Existing feature contracts.
  - `tests/unit/test_gui_and_reporting.py`: 30 unit tests covering all GUI and reporting components.
- **Key findings**:
  - Exact catalog of 50+ emoji instances embedded in strings across all GUI modules.
  - Identification of critical business logic string-dependency: `leader_view.py:1349` (`if "✓" in plm_status:`) and `line 1371` (`if "✓" in r3_status:`). Removing `✓` without updating condition will break file readiness detection.
  - Missing UI components vs Spec: Leader View compact KPI Header cards, Member View 3-step Stepper and Comparison Diff View, Settings Dialog Theme selection tab.
  - Current tables lack 32px height, explicit `#CBD5E1` / `#2A374A` grid lines, and standardized padding.
  - Stylesheets are currently scattered as inline `setStyleSheet(...)` calls with non-WCAG-compliant and hardcoded colors.
  - Zero regression requirement: All 30 tests in `test_gui_and_reporting.py` and 10 tests in `test_f22`/`test_f23` pass; their API surface must be preserved via delegations/properties.
- **Unexplored areas**: None within the survey scope; complete inventory and delta roadmap established.

## Key Decisions Made
- Mapped all widget hierarchies, layout trees, table schemas, and signals across the 6 GUI files.
- Formulated the exact delta specifications and refactoring blueprint for implementer agents.

## Artifact Index
- `handoff.md` — Comprehensive 5-component survey report for Explorer 1
- `DISPATCH.md` — Recorded dispatch instructions with timestamps
- `scratch/survey_inspect.py` — Inspection script used during survey
