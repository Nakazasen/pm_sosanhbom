# Task Assignment: Milestone M2 Worker (Cross-Reconciliation & MSI Engine)

## Identity
- Archetype: teamwork_preview_worker
- Role: Milestone M2 Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_m2_reconciliation
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Mandatory Integrity Warning
> DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Scope & File Ownership
You exclusively own and must implement:
- `src/core/reconciliation.py`:
  - F6: Line-by-line three-way reconciliation (CTTT vs PLM vs R3) evaluating part quantities, revision matching, and row status (`OK` / `NG`).
  - F7: Missing parts detection (reverse lookup on PLM identifying engineered parts omitted in CTTT).
  - F8: Cross-station total aggregation (`CTTT_Total`) summing shared parts across sub-units and reconciling against R3 machine totals.
  - F9: Annotation migration engine preserving and migrating member explanations from old PLM sheets across revisions (`ham_match_index_mix`).
- `src/core/msi_engine.py`:
  - F10: MSI 9-branch decision table evaluating barcode parts, 3-character fixed codes, and service comments against `FIX_SERIAL_DLTOOL_VER010.xls` (or cached master dictionary), with 9 distinct evaluation branches and status output (`OK` / `NG`).
- Update `src/core/__init__.py` to export reconciliation and MSI classes.
- `tests/unit/test_reconciliation_msi.py`: Comprehensive unit tests covering 3-way reconciliation, omission detection, total aggregation, annotation migration, and all 9 branches of the MSI decision table.

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\PROJECT.md`.
- Read Explorer 1's detailed specification report at `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md` (Sections 1.4 and 1.5).
- Inspect existing modules in `src/core/` implemented by Worker M1 (`models.py`, `tree_parser.py`, etc.).

## Deliverables
- Implement `src/core/reconciliation.py` and `src/core/msi_engine.py`.
- Run pytest to verify all unit tests pass with >= 80% coverage.
- Write your completion report to `D:\Sandbox\pm_sosanhbom\.agents\worker_m2_reconciliation\handoff.md`.
- Send message to parent orchestrator when complete.

## 2026-09-17T03:28:45Z
Received task assignment to implement F6, F7, F8, F9 (reconciliation.py), F10 (msi_engine.py), update src/core/__init__.py, and write tests/unit/test_reconciliation_msi.py.

