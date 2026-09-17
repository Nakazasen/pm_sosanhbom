# Task Assignment: Milestone M3 Worker (Teamcenter TC14 Web Automation Module)

## Identity
- Archetype: teamwork_preview_worker
- Role: Milestone M3 Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_m3_tc14
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Scope & File Ownership
You exclusively own and must implement:
- `src/automation/tc14/__init__.py`
- `src/automation/tc14/selectors.py`: Verified DOM and CSS/XPath selectors for login, search, tree expansion, and Excel export.
- `src/automation/tc14/session.py`: Headless browser driver factory (Edge and Chrome with `--headless=new`, CDP download behavior setup, window size 1920x1080), session lifecycle management, cookie persistence, and session timeout detection.
- `src/automation/tc14/client.py`: High-level `TC14AutomationClient` supporting automated authentication (`vn_pe03 / vn_pe03`), direct hash search navigation (`#/teamcenter.search.search?searchCriteria=...`), item opening, native Excel export (`Awp0ExportToExcel`), and fallback interactive tree scraper.
- `tests/unit/test_tc14_automation.py`: Unit and mock tests for TC14 client, selectors, and download routing.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read Explorer 2's detailed specification report at `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2\handoff.md`.
- You can inspect working probe scripts in `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2/` (`test_tc14_login.py`, `test_tc14_content_and_export.py`).

## Deliverables
- Implement all modules in `src/automation/tc14/`.
- Run unit/integration tests with `pytest` and verify pass rate.
- Write your completion report to `D:\Sandbox\pm_sosanhbom\.agents\worker_m3_tc14\handoff.md`.
- Send message to parent orchestrator when complete.

## 2026-09-17T03:10:29Z
You are Worker M3 (Teamcenter TC14 Web Automation Module).
Your working directory is: D:\Sandbox\pm_sosanhbom\.agents\worker_m3_tc14
Original User Request is at: D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md
Project Index & Feature Inventory is at: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md
Your task dispatch file is at: D:\Sandbox\pm_sosanhbom\.agents\worker_m3_tc14\DISPATCH.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

You exclusively own and must implement:
- src/automation/tc14/__init__.py
- src/automation/tc14/selectors.py: Verified DOM and CSS/XPath selectors for login, search, tree expansion, and Excel export.
- src/automation/tc14/session.py: Headless browser driver factory (Edge and Chrome with --headless=new, CDP download behavior setup, window size 1920x1080), session lifecycle management, cookie persistence, and session timeout detection.
- src/automation/tc14/client.py: High-level TC14AutomationClient supporting automated authentication (vn_pe03 / vn_pe03), direct hash search navigation (#/teamcenter.search.search?searchCriteria=...), item opening, native Excel export (Awp0ExportToExcel), and fallback interactive tree scraper.
- tests/unit/test_tc14_automation.py: Unit and mock tests for TC14 client, selectors, and download routing.

Reference Explorer 2's detailed report at D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2\handoff.md.
Run pytest to verify your tests pass.
Write your completion report to D:\Sandbox\pm_sosanhbom\.agents\worker_m3_tc14\handoff.md with test commands and results.
Send message to parent orchestrator when complete.

