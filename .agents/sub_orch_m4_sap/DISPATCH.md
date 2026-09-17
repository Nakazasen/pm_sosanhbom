# Task Assignment: Milestone M4 Sub-Orchestrator (SAP R3 CS12 Automation)

## Identity
- Archetype: teamwork_preview_orchestrator
- Role: Milestone M4 Sub-Orchestrator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\sub_orch_m4_sap
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Scope & Objective
Implement and verify Milestone M4: SAP R3 Multilevel BOM (CS12) Automation via SAP GUI Scripting.
Features:
- F15: SAP GUI 770 COM automation (`saplogon.exe`) connecting to system `P1J(ERP60-AWS)-VN` via `win32com.client`
- F16: Multi-logon detection and handling via option 2 (`radMULTI_LOGON_OPT2`)
- F17: CS12 transaction execution (Plant `2200`, BOM Usage `pp01`, Alternative `01`, valid date `YYYY/MM/DD`)
- F18: Status bar fail-closed safety guard (`wnd[0]/sbar.MessageType == 'E'`) preventing crashes on missing BOMs
- F19: SAP spreadsheet export via `SAPLSPO5:0150` option `radSPOPLI-SELFLAG[1,0]` to target directory
- F20: Machine model code unification (unifying `ma1` `110*` and `maT` `T10*` with configurable pattern)
- F21: Dynamic R3 header resolution (robust parsing of Part Code, Quantity, Revision without brittle column deletion)

File ownership: `src/automation/sap/connection.py`, `src/automation/sap/cs12.py`, `src/automation/sap/parser.py`, `src/automation/sap/__init__.py`.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md`
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`
- Read Explorer 3 handoff report: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\handoff.md`

## Procedure
Apply Project Orchestrator procedure (Assess -> Iteration Loop: Explorer -> Worker -> Reviewer -> Challenger -> Auditor -> Gate).
Ensure tests pass, reviewers approve, challenger confirms, and auditor provides CLEAN verdict.
Write `SCOPE.md`, `GATE_STATUS.md`, `progress.md`, and report completion with `handoff.md` back to parent.
