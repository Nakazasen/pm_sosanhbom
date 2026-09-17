# Task Assignment: Milestone M4 Worker (SAP R3 CS12 Automation Module)

## Identity
- Archetype: teamwork_preview_worker
- Role: Milestone M4 Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_m4_sap
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Scope & File Ownership
You exclusively own and must implement:
- `src/automation/sap/__init__.py`
- `src/automation/sap/models.py`: Data models (`SAPCredentials`, `CS12Params`, `ExportResult`).
- `src/automation/sap/connection.py`: `SAPConnectionManager` handling process discovery (`saplogon.exe 770`), ROT wait, `P1J(ERP60-AWS)-VN` connection acquisition, multi-logon resolution (`radMULTI_LOGON_OPT2`), and login execution.
- `src/automation/sap/cs12.py`: `CS12Service` handling `/nCS12` navigation, parameter binding (Plant 2200, BOM Usage pp01, Alternative 01, date `YYYY/MM/DD`), status bar fail-closed error detection (`wnd[0]/sbar.MessageType == 'E'`), `btn[45]` export trigger, `SAPLSPO5:0150` dialog selection, and directory routing.
- `src/automation/sap/parser.py`: Resilient R3 spreadsheet parser dynamically identifying `Component`, `Quantity`, and `RevLev` headers across layout variants without fragile column deletions.
- `tests/unit/test_sap_automation.py`: Unit and mock tests for connection manager, CS12 service, and resilient parser.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read Explorer 3's detailed specification report at `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\handoff.md`.

## Deliverables
- Implement all modules in `src/automation/sap/`.
- Run unit tests with `pytest` and verify pass rate.
- Write your completion report to `D:\Sandbox\pm_sosanhbom\.agents\worker_m4_sap\handoff.md`.
- Send message to parent orchestrator when complete.

## 2026-09-17T03:10:30Z
You are Worker M4 (SAP R3 CS12 Automation Module).
Your working directory is: D:\Sandbox\pm_sosanhbom\.agents\worker_m4_sap
Original User Request is at: D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md
Project Index & Feature Inventory is at: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md
Your task dispatch file is at: D:\Sandbox\pm_sosanhbom\.agents\worker_m4_sap\DISPATCH.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

You exclusively own and must implement:
- src/automation/sap/__init__.py
- src/automation/sap/models.py: Data models (SAPCredentials, CS12Params, ExportResult).
- src/automation/sap/connection.py: SAPConnectionManager handling process discovery (saplogon.exe 770), ROT wait, P1J(ERP60-AWS)-VN connection acquisition, multi-logon resolution (radMULTI_LOGON_OPT2), and login execution.
- src/automation/sap/cs12.py: CS12Service handling /nCS12 navigation, parameter binding (Plant 2200, BOM Usage pp01, Alternative 01, date YYYY/MM/DD), status bar fail-closed error detection (wnd[0]/sbar.MessageType == 'E'), btn[45] export trigger, SAPLSPO5:0150 dialog selection, and directory routing.
- src/automation/sap/parser.py: Resilient R3 spreadsheet parser dynamically identifying Component, Quantity, and RevLev headers across layout variants without fragile column deletions.
- tests/unit/test_sap_automation.py: Unit and mock tests for connection manager, CS12 service, and resilient parser.

Reference Explorer 3's detailed report at D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\handoff.md.
Run pytest to verify your tests pass.
Write your completion report to D:\Sandbox\pm_sosanhbom\.agents\worker_m4_sap\handoff.md with test commands and results.
Send message to parent orchestrator when complete.

