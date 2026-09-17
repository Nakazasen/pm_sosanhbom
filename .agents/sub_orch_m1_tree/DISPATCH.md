# Task Assignment: Milestone M1 Sub-Orchestrator (Core BOM Tree & Unit Resolver)

## Identity
- Archetype: teamwork_preview_orchestrator
- Role: Milestone M1 Sub-Orchestrator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\sub_orch_m1_tree
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Scope & Objective
Implement and verify Milestone M1: Core BOM Data Structures, Tree Filter & Rapid Unit Resolver Engine.
Features:
- F1: 14-Column PLM Excel Parser (`src/core/tree_parser.py`)
- F2: 6-Level Hierarchical In-Memory BOM Tree (`src/core/models.py`)
- F3: Dual-Pass Date Validity Filtering ("to", "UP", elapsed year/month, empty effectivity) (`src/core/date_filter.py`)
- F4: Model Decomposition & Sub-assembly Pruning across 6 machine types based on `BolocBom` (`src/core/model_pruner.py`)
- F5: Rapid $O(N)$ Unit Resolver replacing 5,619-row lookup sheet (`src/core/unit_resolver.py`)

File ownership: `src/core/models.py`, `src/core/tree_parser.py`, `src/core/date_filter.py`, `src/core/model_pruner.py`, `src/core/unit_resolver.py`.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md`
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`
- Read Explorer 1 handoff report: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md`

## Procedure
Apply Project Orchestrator procedure (Assess -> Iteration Loop: Explorer -> Worker -> Reviewer -> Challenger -> Auditor -> Gate).
Ensure tests pass, reviewers approve, challenger confirms, and auditor provides CLEAN verdict.
Write `SCOPE.md`, `GATE_STATUS.md`, `progress.md`, and report completion with `handoff.md` back to parent.
