# Task Assignment: Explorer Fix 3 (Core Tree, Edge Cases & Stress Remediation)

## Identity
- Archetype: teamwork_preview_explorer
- Role: Core Tree, Edge Cases & Stress Remediation Investigator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_3
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Context
Gate Iteration 1 FAILED due to Forensic Auditor INTEGRITY VIOLATION, Reviewer REQUEST_CHANGES, and Challenger DEFECTS_DETECTED.
You MUST read:
- `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` (authoritative user requirements, especially Section R3)
- `D:\Sandbox\pm_sosanhbom\PROJECT.md`
- `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md`
- `D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md` (Finding 1.3 Group B & C)
- `D:\Sandbox\pm_sosanhbom\.agents\challenger_1\handoff.md` (Deep hierarchy & benchmark findings)
- `D:\Sandbox\pm_sosanhbom\.agents\challenger_2\handoff.md` (Cyclic graphs, revision normalization, comma parsing)

## Objectives & Scope
Investigate and design exact remediation code changes for:
1. **Audit Finding 1.3 Group B & Challenger 1 Observation 1 — Hierarchy Depth Limit**:
   - Inspect `src/core/models.py:38`: `level: int = Field(ge=0, le=10)`.
   - Design change to relax validation to `le=20` (or `le=30`) so that industrial BOMs with 11 or 12 levels validate cleanly.
2. **Challenger 2 Finding — Infinite Recursion on Cyclic BOM Graphs**:
   - Inspect `DateFilter._filter_node_recursive()` in `src/core/date_filter.py` and `ModelPruner._apply_rule_to_nodes()` in `src/core/model_pruner.py`.
   - Design cycle detection using `visited: set[int] = None` tracking object IDs `id(node)` or `(node.item_id, node.level)` to break cycles gracefully without `RecursionError`.
3. **Challenger 2 Finding — Revision Normalization Elasticity**:
   - Inspect `_normalize_rev()` in `src/core/reconciliation.py`.
   - Design elastic normalization handling: leading zeros (`"01"` vs `"1"`), case insensitivity (`"a"` vs `"A"`), float formatting (`"1.0"` vs `"01"`), and prefix stripping (`"Rev.01"` vs `"01"`).
4. **Challenger 2 Finding — Comma Decimal Parsing**:
   - Inspect quantity parsing in `src/core/reconciliation.py` and `src/core/tree_parser.py`.
   - Ensure strings with decimal commas (e.g. `"1,5"`) are normalized to decimal point (`"1.5"`) before converting to float.
5. **Challenger 1 Finding — QThread Lifecycle Guard in Leader View**:
   - Inspect `src/gui/leader_view.py:568-575`.
   - Add guard against re-entrancy if `self.thread and self.thread.isRunning()`.

Write your comprehensive remediation report to `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_3\handoff.md` and message parent orchestrator.
