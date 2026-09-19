# DISPATCH: Codebase & Architecture Explorer (Survey Agent 3)

## Mission
Survey the existing Python/PyQt6 codebase, tests, scripts, and packaging utilities in `D:\Sandbox\pm_sosanhbom` to evaluate current architectural status, feature coverage against R1..R6, and identify all gaps.

## Inputs
- `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` (MUST READ FIRST)
- Workspace directory: `D:\Sandbox\pm_sosanhbom`

## Working Directory
- `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3`

## Specific Areas to Investigate
1. Repository architecture: directory structure, modules (`gui/`, `core/`, `services/`, `models/`, `utils/`, etc.).
2. Leader Workspace: current UI layout vs required 4-step wizard. What exists and what is missing?
3. Member Workspace: current UI vs required formnguoidung spec (auto-loading, 3 tables, self-check, Q2=OK stamp).
4. Core engines: BOM filter engine (recursive Level 1..6, BolocBom), comparison engine, explanation inheritance (`ham_match_index_mix`), MSI deep cross-check (`FIX_SERIAL_DLTOOL_VER010.xls`), JIG master manager & 4M assessment, Outlook email service.
5. Existing test suite: inspect tests in `tests/`, test fixtures, test runner, run existing tests if needed to report current pass/fail status.
6. Packaging & Deployment: `package_app.py`, batch scripts (`.bat`), executable build readiness.
7. Detailed gap matrix: for each requirement R1..R6, what is fully done, what is partially done, what is missing or broken.

## Output Requirement
Write comprehensive codebase assessment to `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\codebase_report.md` and deliver `handoff.md`. Send completion message to parent.

## 2026-09-19T09:47:07Z
You are the Codebase & Architecture Explorer for the project 'Chương trình so sánh BOM tự động'.
Your working directory is D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3.
Read D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md first!
Then read your dispatch file at D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\DISPATCH.md.
Explore the existing Python/PyQt6 codebase, tests, and packaging tools in D:\Sandbox\pm_sosanhbom.
Inspect the Leader Workspace, Member Workspace, Core Engines (BOM filter, comparison, inheritance, MSI checker, JIG manager, email service), test suite, and package_app.py.
Evaluate the exact gaps against R1..R6.
Write your findings to D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\codebase_report.md and deliver a complete handoff.md in your working directory. Send a message to your parent when complete.

