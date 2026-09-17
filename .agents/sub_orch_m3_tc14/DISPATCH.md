# Task Assignment: Milestone M3 Sub-Orchestrator (Teamcenter TC14 Web Automation)

## Identity
- Archetype: teamwork_preview_orchestrator
- Role: Milestone M3 Sub-Orchestrator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\sub_orch_m3_tc14
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Scope & Objective
Implement and verify Milestone M3: Siemens Teamcenter Active Workspace (TC14) Web Automation at `http://tcmp3gwb:3000/`.
Features:
- F11: Headless Edge/Chrome WebDriver session management with CDP download behavior (`Page.setDownloadBehavior`)
- F12: Automated authentication with `vn_pe03 / vn_pe03`, cookie handling, session timeout detection, and auto-relogin
- F13: Direct hash-based search and item revision opening (`#/teamcenter.search.search?searchCriteria=...`)
- F14: Automated BOM export pipeline triggering `Awp0ExportToExcel` / Content tab export to download full 14-column PLM `.xlsx`

File ownership: `src/automation/tc14/client.py`, `src/automation/tc14/selectors.py`, `src/automation/tc14/session.py`, `src/automation/tc14/__init__.py`.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md`
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`
- Read Explorer 2 handoff report: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2\handoff.md`

## Procedure
Apply Project Orchestrator procedure (Assess -> Iteration Loop: Explorer -> Worker -> Reviewer -> Challenger -> Auditor -> Gate).
Ensure tests pass, reviewers approve, challenger confirms, and auditor provides CLEAN verdict.
Write `SCOPE.md`, `GATE_STATUS.md`, `progress.md`, and report completion with `handoff.md` back to parent.
