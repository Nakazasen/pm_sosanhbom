# Task Assignment: Milestone M1 Worker (Core BOM Tree, Tree Filter & Rapid Unit Resolver)

## Identity
- Archetype: teamwork_preview_worker
- Role: Milestone M1 Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_m1_tree
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Scope & File Ownership
You exclusively own and must implement:
- `src/core/__init__.py`
- `src/core/models.py`: Pydantic/dataclass models for `BOMNode`, `BOMTree`, `FilterCriteria`, `ModelRule`.
- `src/core/tree_parser.py`: Robust 14-column PLM Excel parser extracting `Level`, `Item Id`, `Has Children`, `Quantity`, `Occurrence Effectivities`, `Item Name`, `Revision`.
- `src/core/date_filter.py`: Dual-pass effectivity date filtering ("to <date>" regex extraction, "UP" retention, elapsed year/month comparison with target reference date, empty effectivity pruning for subtrees vs leaf nodes).
- `src/core/model_pruner.py`: Model decomposition pruner supporting the 4 action rules on matched nodes across the 6 machine models (Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris) derived from legacy `BolocBom`.
- `src/core/unit_resolver.py`: Fast single-pass $O(N)$ depth-first stack traversal unit resolver replacing the 5,619-row lookup sheet `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`.
- `tests/unit/test_core_tree.py`: Comprehensive unit tests verifying tree parsing, date filtering, model pruning, and unit resolving.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read Explorer 1's detailed specification report at `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md`.

## Deliverables
- Implement all modules in `src/core/`.
- Run unit tests with `pytest` and verify 100% pass rate.
- Write your completion report to `D:\Sandbox\pm_sosanhbom\.agents\worker_m1_tree\handoff.md` with build/test command results.
- Send message to parent orchestrator when complete.

## 2026-09-17T03:10:29Z
You are Worker M1 (Core BOM Tree & Unit Resolver Engine).
Your working directory is: D:\Sandbox\pm_sosanhbom\.agents\worker_m1_tree
Original User Request is at: D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md
Project Index & Feature Inventory is at: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md
Your task dispatch file is at: D:\Sandbox\pm_sosanhbom\.agents\worker_m1_tree\DISPATCH.md

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

You exclusively own and must implement:
- src/core/__init__.py
- src/core/models.py: Pydantic/dataclass models for BOMNode, BOMTree, FilterCriteria, ModelRule.
- src/core/tree_parser.py: Robust 14-column PLM Excel parser extracting Level, Item Id, Has Children, Quantity, Occurrence Effectivities, Item Name, Revision.
- src/core/date_filter.py: Dual-pass effectivity date filtering ("to <date>" regex extraction, "UP" retention, elapsed year/month comparison with target reference date, empty effectivity pruning for subtrees vs leaf nodes).
- src/core/model_pruner.py: Model decomposition pruner supporting the 4 action rules on matched nodes across the 6 machine models (Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris) derived from legacy BolocBom.
- src/core/unit_resolver.py: Fast single-pass O(N) depth-first stack traversal unit resolver replacing the 5,619-row lookup sheet Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx.
- tests/unit/test_core_tree.py: Comprehensive unit tests verifying tree parsing, date filtering, model pruning, and unit resolving.

Reference Explorer 1's detailed reverse-engineering report at D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md.
Run pytest to verify all your tests pass.
Write your completion report to D:\Sandbox\pm_sosanhbom\.agents\worker_m1_tree\handoff.md with test commands and results.
Send message to parent orchestrator when complete.
