# Task Assignment: E2E Testing Track Orchestrator

## Identity
- Archetype: teamwork_preview_orchestrator
- Role: E2E Testing Orchestrator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\sub_orch_e2e_testing
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Scope & Objective
Lead the E2E Testing Track for the BOM Comparison Automation Modernization project.
Derive and implement a comprehensive opaque-box test suite based directly on requirements in `ORIGINAL_REQUEST.md` and feature inventory in `PROJECT.md`:
1. Create `TEST_INFRA.md` covering test architecture, runners, and format.
2. Build 4-tier test suite:
   - Tier 1: Feature Coverage (>=5 tests per feature for F1..F28).
   - Tier 2: Boundary & Corner Cases (empty BOMs, max 6 levels, leap years, missing dates, unicode names, layout variants).
   - Tier 3: Cross-Feature Combinations (PLM tree filter + Unit resolver + 3-way reconciliation).
   - Tier 4: Real-World Workload Testing (bit-accurate verification against legacy `form_ssbom.xlsm` sample data, `formnguoidung.xlsm`, and `tonghop_*.xlsm`).
3. Create automated test runner command: `pytest tests/ -v`.
4. When complete, publish `TEST_READY.md` at project root with coverage checklist.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md`
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`
- Read Explorer handoff reports:
  - `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md`
  - `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2\handoff.md`
  - `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\handoff.md`

## Procedure
Apply Project Orchestrator procedure (Assess -> Decompose into test creation milestones or run Explorer -> Test Writer / Worker -> Reviewer -> Gate loop).
Enforce forensic integrity and pass/fail gate checks.
Report progress and completion back to parent.
