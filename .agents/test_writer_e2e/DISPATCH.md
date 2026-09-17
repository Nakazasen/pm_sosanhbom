# Task Assignment: E2E Test Suite Writer (Dual Track)

## Identity
- Archetype: teamwork_preview_test_writer
- Role: E2E Test Writer
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\test_writer_e2e
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All tests must be authentic, independently derived from requirements, and genuinely exercise specifications. Dummy tests or hardcoded passing mocks will be rejected by the Forensic Auditor.

## Scope & File Ownership
You exclusively own and must implement:
- `TEST_INFRA.md` at project root: Documentation of testing methodology, tier breakdown, and test commands.
- `tests/conftest.py`: Shared pytest fixtures, sample BOM generators, temporary workspace setup.
- `tests/tier1_features/`: At least 5 test cases per feature covering F1..F28 in isolation (happy-path).
- `tests/tier2_boundaries/`: Boundary conditions (empty BOM, single-row BOM, max 6 levels, date transitions, leap years, non-ASCII part names, special characters).
- `tests/tier3_combinations/`: Combinatorial feature tests (tree filter + unit resolver + reconciliation, date filter with model pruning).
- `tests/tier4_real_world/`: Real-world tests using ground truth data from legacy workbooks in `D:\Sandbox\pm_sosanhbom` (`form_ssbom.xlsm`, `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`).
- `TEST_READY.md` at project root: Publish coverage checklist and test runner command.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read Explorer handoff reports:
  - `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md`
  - `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2\handoff.md`
  - `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\handoff.md`

## Deliverables
- Create test suite in `tests/` with clear directory separation.
- Create `TEST_INFRA.md` and `TEST_READY.md`.
- Verify tests can be collected and run with `pytest`.
- Write your completion report to `D:\Sandbox\pm_sosanhbom\.agents\test_writer_e2e\handoff.md`.


## 2026-09-17T03:10:29Z
You are the E2E Test Suite Writer for the BOM Comparison Automation Modernization project.
Your working directory is: D:\Sandbox\pm_sosanhbom\.agents\test_writer_e2e
Original User Request is at: D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md
Project Index & Feature Inventory is at: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md
Your task dispatch file is at: D:\Sandbox\pm_sosanhbom\.agents\test_writer_e2e\DISPATCH.md

## 2026-09-17T03:43:13Z
From: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43 (parent)
Context: Status Inquiry on E2E Test Suite
Content: Milestones M1, M2, M3, and M4 are now fully implemented in `src/core/` and `src/automation/` with passing unit tests. We noticed your state is waiting/paused.
Action: Please proceed to implement the Tier 4 real-world legacy tests in `tests/tier4_real_world/` (verifying against `form_ssbom.xlsm` and `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`), verify all tests via `pytest tests/`, publish `TEST_READY.md` at project root, and deliver your completion report in `handoff.md`.
