# Task Assignment: Milestone M5 Worker (PyQt6 Desktop GUI Application & Reporting)

## Identity
- Archetype: teamwork_preview_worker
- Role: Milestone M5 Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_m5_gui
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Scope & File Ownership
You exclusively own and must implement:
- `src/gui/__init__.py`
- `src/gui/leader_view.py`:
  - F22: Leader Management Workspace:
    - Create project folder structure by machine model.
    - Check member submission status across sub-units.
    - Trigger batch automated BOM comparison across CTTT, PLM, and R3.
    - Trigger consolidated report generation.
    - Trigger Outlook notification email preview.
- `src/gui/member_view.py`:
  - F23: Member Input Workspace:
    - Sub-unit selection and CTTT component entry / import.
    - MSI barcode and 3-char fixed code entry.
    - Label 7980/7990 input.
    - Preliminary self-check against PLM and R3 with instant OK/NG visual feedback.
    - Final confirmation and submission button.
- `src/gui/settings_dialog.py`:
  - Configuration dialog for Teamcenter credentials, SAP logon parameters, and paths.
- `src/gui/app.py`:
  - Main PyQt6 application window with tabbed or role-based navigation switching between Leader View and Member View, status bar, logging console, and error dialogs.
- `src/reporting/__init__.py`
- `src/reporting/excel_generator.py`:
  - F24: Consolidated report generation matching `form_ssbom.xlsm` workbook schema (Sheets `Tongket`, `List JIG`, `MSI_7980_7990`, `CTTT`, `PLM`, `R3`, `CTTT_Total`), applying exact conditional formatting and color vectors (Red 255 for NG, Green 6750054 for OK).
- `src/reporting/outlook_mailer.py`:
  - F25: Outlook COM email automation (`win32com.client.Dispatch("Outlook.Application")`) generating and displaying/sending HTML notification emails to team members with reconciliation status.
- `tests/unit/test_gui_and_reporting.py`:
  - Comprehensive unit and mock tests verifying Leader View, Member View, report generation, and mailer without requiring interactive display or live Outlook.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read Explorer 1's detailed report at `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md` (Sections 1.4, 1.5, 1.6).
- Use components from `src/core` (`ReconciliationEngine`, `BOMTree`, `MSIEngine`, etc.) and `src/automation` (`TC14Client`, `SAPCS12Client`).

## Deliverables
- Implement all modules in `src/gui/` and `src/reporting/`.
- Run pytest to verify all unit tests pass with >= 80% coverage.
- Write your completion report to `D:\Sandbox\pm_sosanhbom\.agents\worker_m5_gui\handoff.md`.
- Send message to parent orchestrator when complete.
