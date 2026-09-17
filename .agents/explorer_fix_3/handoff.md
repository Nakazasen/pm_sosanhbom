# Handoff Report — Explorer Fix 3 (Core Tree, Edge Cases & Stress Remediation)

**Working Directory**: `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_3`  
**Role**: Core Tree, Edge Cases & Stress Remediation Investigator  
**Archetype**: teamwork_preview_explorer  
**Parent Conversation ID**: `8a26cf43-3f4f-42ea-ac18-3875de8c9a43`  
**Date**: 2026-09-17  

---

## 1. Observation

Direct empirical observations, line numbers, verbatim code snippets, and automated test command outputs across all 5 investigated areas:

### 1.1 Observation 1 — Hierarchy Depth Limit (`src/core/models.py:38`)
- **File**: `D:\Sandbox\pm_sosanhbom\src\core\models.py`, lines 33-38:
  ```python
  class BOMNode(BaseModel):
      """A node in the hierarchical 6-level BOM tree."""

      model_config = ConfigDict(arbitrary_types_allowed=True)

      level: int = Field(ge=0, le=10, description="Hierarchy level (1..6 typically; 0 for root machine body)")
  ```
- **Execution**:
  ```powershell
  python -m pytest tests/tier5_adversarial/test_adversarial_stress_perf.py::TestDeepNestedHierarchy -v
  ```
- **Verbatim Error**:
  ```text
  FAILED tests/tier5_adversarial/test_adversarial_stress_perf.py::TestDeepNestedHierarchy::test_bom_node_level_12_validation
  pydantic_core._pydantic_core.ValidationError: 1 validation error for BOMNode
  level
    Input should be less than or equal to 10 [type=less_than_equal, input_value=12, input_type=int]

  FAILED tests/tier5_adversarial/test_adversarial_stress_perf.py::TestDeepNestedHierarchy::test_bom_tree_deep_chain_12_levels
  pydantic_core._pydantic_core.ValidationError: 1 validation error for BOMNode
  level
    Input should be less than or equal to 10 [type=less_than_equal, input_value=11, input_type=int]

  FAILED tests/tier5_adversarial/test_adversarial_stress_perf.py::TestDeepNestedHierarchy::test_plm_parser_with_12_level_dataset
  src\core\tree_parser.py:341: ValidationError: 1 validation error for BOMNode
  level
    Input should be less than or equal to 10 [type=less_than_equal, input_value=11, input_type=int]
  ```
- **Context**: In complex multi-tier industrial assemblies (printers/MFPs), mechanical nesting frequently extends to levels 11 and 12. Pydantic's hard upper limit of `le=10` prematurely aborts tree instantiation.

---

### 1.2 Observation 2 — Infinite Recursion on Cyclic BOM Graphs (`DateFilter` & `ModelPruner`)
- **Files**:
  - `D:\Sandbox\pm_sosanhbom\src\core\date_filter.py`, lines 170-192 (`_filter_node_recursive`)
  - `D:\Sandbox\pm_sosanhbom\src\core\model_pruner.py`, lines 121-151 (`_apply_rule_to_nodes`)
- **Execution**:
  ```powershell
  python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestCyclicBOMGraphsAndRecursion -v
  ```
- **Verbatim Error**:
  ```text
  src\core\date_filter.py:186: in _filter_node_recursive
      filtered_child = self._filter_node_recursive(child, reference_date)
  E   RecursionError: maximum recursion depth exceeded
  Failed: Defect: DateFilter.filter_tree() entered infinite recursion (RecursionError) on cyclic BOM graph!

  src\core\model_pruner.py:148: in _apply_rule_to_nodes
      node.children = self._apply_rule_to_nodes(node.children, rule)
  E   RecursionError: maximum recursion depth exceeded
  Failed: Defect: ModelPruner.prune_tree() entered infinite recursion (RecursionError) on cyclic BOM graph!
  ```
- **Context**: Unlike `BOMNode.flatten()` (`models.py:66`), `BOMNode.clone()` (`models.py:97`), and `UnitResolver._propagate_unit()` (`unit_resolver.py:53`) which all maintain a `visited: set[int]` of object IDs `id(node)`, `DateFilter` and `ModelPruner` lack cycle guards. When corrupt, circular, or self-referential BOM trees occur in customer data, the Python call stack is exhausted.

---

### 1.3 Observation 3 — Revision Normalization Brittleness (`src/core/reconciliation.py:57-63`)
- **File**: `D:\Sandbox\pm_sosanhbom\src\core\reconciliation.py`, lines 57-63:
  ```python
  def _normalize_rev(val: Any) -> str:
      """Normalize revision string, stripping whitespace."""
      if val is None or pd.isna(val):
          return ""
      s = str(val).strip()
      return "" if s.lower() == "nan" else s
  ```
- **Execution**:
  ```powershell
  python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestRevisionVariationsAndNormalization -v
  ```
- **Verbatim Error**:
  ```text
  FAILED test_leading_zero_revision_matching:
  AssertionError: Defect: Leading-zero revision mismatch: PLM '01' vs R3 '1' produced comp_rev='NG'
  assert 'NG' == 'OK'

  FAILED test_case_insensitivity_in_revision_matching:
  AssertionError: Defect: Case-sensitive revision mismatch: 'a' vs 'A' produced comp_rev='NG'
  assert 'NG' == 'OK'

  FAILED test_float_representation_of_revision_from_excel:
  AssertionError: Defect: Float-imported revision 1.0 vs '01' produced comp_rev='NG'
  assert 'NG' == 'OK'

  FAILED test_slash_and_prefix_revision_notations:
  AssertionError: Defect: Revision notation 'Rev.01' vs '01' failed to match: 'NG'
  assert 'NG' == 'OK'
  ```
- **Context**: Teamcenter Active Workspace exports leading-zero numeric revisions (`"01"`), while SAP R3 exports or Excel sheet formulas frequently yield single digits (`"1"`), float strings (`1.0`), or prefix labels (`"Rev.01"`). Literal string comparison causes false `NG` reconciliation flags.

---

### 1.4 Observation 4 — Silent Decimal Truncation on Comma Notation (`src/core/reconciliation.py:47-55`)
- **File**: `D:\Sandbox\pm_sosanhbom\src\core\reconciliation.py`, lines 47-55:
  ```python
  def _to_float(val: Any) -> float:
      """Safely convert value to float, defaulting to 0.0 if invalid or NaN."""
      if val is None or pd.isna(val):
          return 0.0
      try:
          return float(val)
      except (ValueError, TypeError):
          return 0.0
  ```
- **Execution**:
  ```powershell
  python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestQuantityPrecisionAndRounding::test_european_comma_decimal_handling_in_to_float -v
  ```
- **Verbatim Error**:
  ```text
  AssertionError: Defect: '1,5' was converted to 0.0 instead of 1.5
  assert 0.0 == 1.5
  ```
- **Context**: On Vietnamese/European operating system locales or when operators type comma decimals (`1,5`), `float("1,5")` raises `ValueError`, coercing valid quantities to `0.0`. In contrast, `src/automation/sap/parser.py:567` properly sanitizes commas with `s.replace(",", ".")`.

---

### 1.5 Observation 5 — Unchecked QThread Re-entrancy (`src/gui/leader_view.py:568-583`)
- **File**: `D:\Sandbox\pm_sosanhbom\src\gui\leader_view.py`, lines 563-583:
  ```python
          # Launch QThread worker
          self.btn_batch_reconcile.setEnabled(False)
          self.progress_bar.setValue(5)
          self.lbl_progress_status.setText("Khởi động tiến trình đối soát hàng loạt...")

          self.thread = QThread()
          self.worker = BatchReconciliationWorker(
              collected_cttt=collected,
              plm_path=plm_path,
              r3_path=r3_path,
              model_name=model,
          )
          self.worker.moveToThread(self.thread)
  ```
- **Context**: `trigger_batch_reconciliation` instantiates `self.thread = QThread()` without first checking `if self.thread and self.thread.isRunning(): return`. Rapid double-clicks or programmatic re-invocations overwrite `self.thread`, orphaning the background OS thread, which can cause `QThread: Destroyed while thread is still running` and native application crashes. Furthermore, `worker.deleteLater` and `thread.deleteLater` lifecycle signals are not attached.

---

## 2. Logic Chain

1. **Hierarchy Depth Constraint (Observation 1.1)**:
   - *Requirement*: Engineering BOM trees must accommodate nested structures up to 12 levels (Section R3 of `ORIGINAL_REQUEST.md`).
   - *Cause*: `BOMNode.level` in `models.py:38` specifies `Field(ge=0, le=10)`. When level reaches 11 or 12, Pydantic V2 raises `ValidationError`.
   - *Deduction*: Relaxing `le=10` to `le=20` eliminates this artificial bottleneck while maintaining input validation sanity against absurd values (e.g. negative numbers or runaway level indices > 20).

2. **Graph Traversal & Cycles (Observation 1.2)**:
   - *Mechanism*: BOM trees in Teamcenter Active Workspace can contain circular references if occurrence structures or shared sub-assemblies are improperly linked.
   - *Cause*: `DateFilter._filter_node_recursive` and `ModelPruner._apply_rule_to_nodes` traverse `node.children` recursively without tracking visited nodes.
   - *Deduction*: Adopting the exact cycle pattern from `UnitResolver._propagate_unit` (`visited: set[int] | None = None` storing `id(node)`) bounds the recursion, preventing `RecursionError` and terminating the traversal gracefully.

3. **Revision Normalization (Observation 1.3)**:
   - *Requirement*: Multi-source reconciliation requires semantic equivalence matching across heterogeneous systems (Teamcenter PLM vs SAP R3 vs Excel CTTT).
   - *Cause*: Teamcenter uses 2-digit zero-padded revisions (`"01"`), SAP R3 CS12 uses integer strings (`"1"`), Excel imports numeric cells as floats (`1.0`), and CAD/PLM notations include prefixes (`"Rev.01"` or `"/A;"`).
   - *Deduction*: `_normalize_rev` must sanitize whitespace, strip common prefixes (`"Rev."`, `"Rev "`), strip punctuation (`/`, `;`), collapse integer-equivalent floats (`1.0` -> `1`), and normalize numeric digits via `str(int(s))` so `"01"` and `"1"` evaluate to identical string representations. Non-numeric revisions must be uppercase-folded (`"a"` -> `"A"`).

4. **Comma Decimal Conversion (Observation 1.4)**:
   - *Cause*: Standard Python `float("1,5")` raises `ValueError`. `_to_float()` swallows the exception and returns `0.0`.
   - *Impact*: In CTTT, PLM, or R3 data with comma decimals, quantities turn to 0.0, evaluating row checks to `NG`.
   - *Deduction*: Applying regex sanitation to convert decimal commas (`"," -> "."`) and strip thousand separators prior to `float()` conversion (matching `ResilientR3Parser` in `src/automation/sap/parser.py`) guarantees that `"1,5"` parses accurately to `1.5`.

5. **QThread Concurrency Safety (Observation 1.5)**:
   - *Mechanism*: In PyQt6, instantiating a new `QThread` and assigning it to `self.thread` while the previous thread is active de-references the C++ wrapper of the running thread.
   - *Deduction*: Adding an early guard `if self.thread and self.thread.isRunning(): return` prevents re-entrancy. Adding `worker.deleteLater` and `thread.deleteLater` connections ensures clean memory reclamation upon task completion.

---

## 3. Caveats

- **Depth Limit Upper Bound**: We set `le=20`. In Kyocera copier/printer manufacturing, BOM depth rarely exceeds 8 to 12 levels. An upper bound of 20 allows extreme sub-assembly hierarchies while preventing runaway loops from corrupt inputs.
- **DAG vs Cycle Semantics**: A BOM node instance appearing in multiple sub-assemblies (DAG) will have its children processed during the first visit. Subsequent visits by reference ID will be skipped, which is correct and avoids duplicate traversal.
- **Revision Prefix Scope**: The prefix stripper matches `"Rev."`, `"Rev "`, `"Revision"`, and delimiters like `/` and `;`. Unorthodox custom prefixes outside these standard CAD/PLM patterns will be treated as raw alphanumeric strings.

---

## 4. Conclusion & Proposed Remediation Steps

### Concrete Remediation Inventory

#### Change 1: `src/core/models.py` (Line 38)
Relax level validation from `le=10` to `le=20`.

```python
# Before (Line 38):
level: int = Field(ge=0, le=10, description="Hierarchy level (1..6 typically; 0 for root machine body)")

# After:
level: int = Field(ge=0, le=20, description="Hierarchy level (1..6 typically; up to 20 for deep nested hierarchies; 0 for root machine body)")
```

---

#### Change 2A: `src/core/date_filter.py` (Lines 151-192)
Add `visited: set[int]` to `filter_tree` and `_filter_node_recursive`.

```python
# Before (Lines 152-192):
    def filter_tree(
        self,
        tree: BOMTree,
        reference_date: datetime.date | None = None,
        in_place: bool = False,
    ) -> BOMTree:
        """Filter the entire BOMTree in-memory, pruning expired nodes and subtrees."""
        ref_date = reference_date or self.criteria.reference_date or _get_current_date()
        target_tree = tree if in_place else tree.clone()

        filtered_roots: list[BOMNode] = []
        for root in target_tree.roots:
            filtered_node = self._filter_node_recursive(root, ref_date)
            if filtered_node is not None:
                filtered_roots.append(filtered_node)

        target_tree.roots = filtered_roots
        return target_tree

    def _filter_node_recursive(
        self,
        node: BOMNode,
        reference_date: datetime.date,
    ) -> BOMNode | None:
        if not self.is_valid(node.effectivity, reference_date):
            return None

        # Filter children recursively
        surviving_children: list[BOMNode] = []
        for child in node.children:
            filtered_child = self._filter_node_recursive(child, reference_date)
            if filtered_child is not None:
                surviving_children.append(filtered_child)

        node.children = surviving_children
        return node

# After:
    def filter_tree(
        self,
        tree: BOMTree,
        reference_date: datetime.date | None = None,
        in_place: bool = False,
    ) -> BOMTree:
        """Filter the entire BOMTree in-memory, pruning expired nodes and subtrees."""
        ref_date = reference_date or self.criteria.reference_date or _get_current_date()
        target_tree = tree if in_place else tree.clone()

        visited: set[int] = set()
        filtered_roots: list[BOMNode] = []
        for root in target_tree.roots:
            filtered_node = self._filter_node_recursive(root, ref_date, visited=visited)
            if filtered_node is not None:
                filtered_roots.append(filtered_node)

        target_tree.roots = filtered_roots
        return target_tree

    def _filter_node_recursive(
        self,
        node: BOMNode,
        reference_date: datetime.date,
        visited: set[int] | None = None,
    ) -> BOMNode | None:
        """Recursively filter a node and its children with cyclic graph protection.
        
        If node itself is expired, returns None (pruning node and all descendants).
        Otherwise, filters children recursively and returns node.
        """
        if visited is None:
            visited = set()
        if id(node) in visited:
            return node
        visited.add(id(node))

        if not self.is_valid(node.effectivity, reference_date):
            return None

        # Filter children recursively
        surviving_children: list[BOMNode] = []
        for child in node.children:
            filtered_child = self._filter_node_recursive(child, reference_date, visited=visited)
            if filtered_child is not None:
                surviving_children.append(filtered_child)

        node.children = surviving_children
        return node
```

---

#### Change 2B: `src/core/model_pruner.py` (Lines 121-152)
Add `visited: set[int]` to `_apply_rule_to_nodes`.

```python
# Before (Lines 121-152):
    def _apply_rule_to_nodes(
        self,
        nodes: list[BOMNode],
        rule: ModelRule,
    ) -> list[BOMNode]:
        """Apply a single ModelRule to a list of sibling nodes and their subtrees."""
        surviving_nodes: list[BOMNode] = []

        for node in nodes:
            if rule.matches(node):
                action = rule.determine_action(node)

                if action == PruneAction.DELETE_NODE:
                    # Rule 1 or Rule 4: Delete the node itself
                    logger.debug("Pruning node '%s' (%s) under action %s", node.item_name, node.item_id, action)
                    continue
                elif action == PruneAction.DELETE_CHILDREN:
                    # Rule 2: Keep the unit assembly node, but prune all its child sub-components
                    logger.debug("Pruning children of unit '%s' (%s) under Rule 2", node.item_name, node.item_id)
                    node.children = []
                    surviving_nodes.append(node)
                elif action == PruneAction.KEEP:
                    # Rule 3: Keep node
                    surviving_nodes.append(node)
            else:
                # Node didn't match rule; apply rule recursively to its children
                if node.children:
                    node.children = self._apply_rule_to_nodes(node.children, rule)
                surviving_nodes.append(node)

        return surviving_nodes

# After:
    def _apply_rule_to_nodes(
        self,
        nodes: list[BOMNode],
        rule: ModelRule,
        visited: set[int] | None = None,
    ) -> list[BOMNode]:
        """Apply a single ModelRule to a list of sibling nodes and their subtrees with cycle protection."""
        if visited is None:
            visited = set()

        surviving_nodes: list[BOMNode] = []

        for node in nodes:
            if id(node) in visited:
                surviving_nodes.append(node)
                continue
            visited.add(id(node))

            if rule.matches(node):
                action = rule.determine_action(node)

                if action == PruneAction.DELETE_NODE:
                    # Rule 1 or Rule 4: Delete the node itself
                    logger.debug("Pruning node '%s' (%s) under action %s", node.item_name, node.item_id, action)
                    continue
                elif action == PruneAction.DELETE_CHILDREN:
                    # Rule 2: Keep the unit assembly node, but prune all its child sub-components
                    logger.debug("Pruning children of unit '%s' (%s) under Rule 2", node.item_name, node.item_id)
                    node.children = []
                    surviving_nodes.append(node)
                elif action == PruneAction.KEEP:
                    # Rule 3: Keep node
                    surviving_nodes.append(node)
            else:
                # Node didn't match rule; apply rule recursively to its children
                if node.children:
                    node.children = self._apply_rule_to_nodes(node.children, rule, visited=visited)
                surviving_nodes.append(node)

        return surviving_nodes
```

---

#### Change 3 & 4: `src/core/reconciliation.py` (Lines 10-63)
Add `import re` and enhance `_to_float` and `_normalize_rev`.

```python
# Before (Lines 10-63):
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any

import pandas as pd

from src.core.models import BOMTree
...
def _to_float(val: Any) -> float:
    """Safely convert value to float, defaulting to 0.0 if invalid or NaN."""
    if val is None or pd.isna(val):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _normalize_rev(val: Any) -> str:
    """Normalize revision string, stripping whitespace."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

# After:
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Any

import pandas as pd

from src.core.models import BOMTree
...
def _to_float(val: Any) -> float:
    """Safely convert value to float, handling commas and defaulting to 0.0 if invalid or NaN.
    
    Supports:
        - Native int and float types
        - European/Vietnamese comma decimal strings ('1,5' -> 1.5)
        - European thousands dot + comma decimal ('1.250,50' -> 1250.5)
        - Standard thousands comma + decimal dot ('1,250.50' -> 1250.5)
        - Fallback to 0.0 on None, NaN, empty strings, or unparseable text
    """
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return 0.0

    # Clean European thousand dots with comma decimal (e.g. 1.250,50)
    if re.search(r"^\d{1,3}(\.\d{3})*,\d+$", s):
        s = s.replace(".", "").replace(",", ".")
    # Clean standard thousand commas with dot decimal (e.g. 1,250.50)
    elif re.search(r"^\d{1,3}(,\d{3})*\.\d+$", s):
        s = s.replace(",", "")
    # Single comma decimal (e.g. '1,5' or '0,5')
    elif "," in s and "." not in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _normalize_rev(val: Any) -> str:
    """Normalize revision string handling leading zeros, casing, floats, and prefixes.
    
    Examples:
        None, 'nan', '' -> ''
        '  B \t\n' -> 'B'
        'a' -> 'A'
        '01' -> '1'
        1.0 -> '1'
        '1.0' -> '1'
        'Rev.01' -> '1'
        'REV A' -> 'A'
        '/A;' -> 'A'
        '-', '--' -> ''
    """
    if val is None or pd.isna(val):
        return ""

    if isinstance(val, (int, float)):
        if isinstance(val, float) and val.is_integer():
            s = str(int(val))
        else:
            s = str(val)
    else:
        s = str(val)

    s = s.strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""

    if s in ("-", "--"):
        return ""

    # Strip enclosing punctuation e.g. "/A;", "(01)"
    s = s.strip("/;:,()[]{}")

    # Strip revision prefixes (e.g. "Rev.01", "REV A", "Revision 02")
    s_upper = s.upper()
    prefixes = ["REVISION.", "REVISION ", "REVISION", "REV.", "REV ", "REV_", "REV-", "REV"]
    for prefix in prefixes:
        if s_upper.startswith(prefix):
            s = s[len(prefix):].strip(" .-_/;:()")
            s_upper = s.upper()
            break

    # Strip remaining leading/trailing punctuation
    s = s.strip("/;:,()[]{} ._")

    # Handle float string representation from Excel imports (e.g. "1.0", "01.0")
    if re.match(r"^\d+\.0+$", s):
        s = s.split(".")[0]

    # If purely numeric, strip leading zeros for consistent integer comparison (e.g. "01" -> "1")
    if s.isdigit():
        s = str(int(s))
    else:
        # Uppercase non-numeric revisions (e.g. "a" -> "A")
        s = s.upper()

    return s
```

---

#### Change 5: `src/gui/leader_view.py` (Lines 518-525 & 568-615)
Add QThread concurrency guard and lifecycle cleanup in `LeaderWorkspaceView`.

```python
# Before (Lines 518-525 and Lines 568-583):
    def trigger_batch_reconciliation(self) -> None:
        """Start asynchronous batch automated BOM comparison across CTTT, PLM, and R3."""
        model = self.model_combo.currentText().strip()
        ...
        self.thread = QThread()
        self.worker = BatchReconciliationWorker(
            collected_cttt=collected,
            plm_path=plm_path,
            r3_path=r3_path,
            model_name=model,
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.error.connect(self._on_worker_error)

        self.thread.start()

# After:
    def trigger_batch_reconciliation(self) -> None:
        """Start asynchronous batch automated BOM comparison across CTTT, PLM, and R3."""
        # Concurrency guard: avoid orphaning running thread
        if self.thread is not None and self.thread.isRunning():
            logger.warning("Batch reconciliation thread is already running; ignoring duplicate trigger.")
            return

        model = self.model_combo.currentText().strip()
        ...
        self.thread = QThread()
        self.worker = BatchReconciliationWorker(
            collected_cttt=collected,
            plm_path=plm_path,
            r3_path=r3_path,
            model_name=model,
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.error.connect(self._on_worker_error)

        # Connect cleanup lifecycle hooks to prevent leaking resources
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def _on_worker_finished(self, result: ReconciliationResult) -> None:
        self.current_result = result
        self.btn_batch_reconcile.setEnabled(True)
        self.progress_bar.setValue(100)
        self.lbl_progress_status.setText(f"Đã đối soát xong! Phán định: {result.overall_status}")
        self.batch_finished.emit(result)

        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()
        self.thread = None
        self.worker = None

        QMessageBox.information(...)

    def _on_worker_error(self, err_msg: str) -> None:
        self.btn_batch_reconcile.setEnabled(True)
        self.lbl_progress_status.setText(f"Lỗi: {err_msg}")
        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()
        self.thread = None
        self.worker = None

        QMessageBox.critical(...)
```

---

## 5. Verification Method

Once Worker Fix 3 applies these changes, run the following commands to independently verify that all 10 previously failing probes pass:

### Step 1: Run Deep Hierarchy Benchmark
```powershell
python -m pytest tests/tier5_adversarial/test_adversarial_stress_perf.py::TestDeepNestedHierarchy -v
```
*Expected Result*: 3 PASSED (`test_bom_node_level_12_validation`, `test_bom_tree_deep_chain_12_levels`, `test_plm_parser_with_12_level_dataset`).

### Step 2: Run Cyclic BOM Recursion Suite
```powershell
python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestCyclicBOMGraphsAndRecursion -v
```
*Expected Result*: 3 PASSED (`test_bom_node_flatten_and_clone_with_self_cycle`, `test_date_filter_recursive_infinite_loop_on_cyclic_graph`, `test_model_pruner_recursive_infinite_loop_on_cyclic_graph`).

### Step 3: Run Revision Normalization Suite
```powershell
python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestRevisionVariationsAndNormalization -v
```
*Expected Result*: 5 PASSED (`test_leading_zero_revision_matching`, `test_case_insensitivity_in_revision_matching`, `test_float_representation_of_revision_from_excel`, `test_whitespace_and_newline_in_revisions`, `test_slash_and_prefix_revision_notations`).

### Step 4: Run Comma Decimal Precision Suite
```powershell
python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestQuantityPrecisionAndRounding::test_european_comma_decimal_handling_in_to_float -v
```
*Expected Result*: 1 PASSED (`test_european_comma_decimal_handling_in_to_float`).

### Step 5: Run Combined Target Suite
```powershell
python -m pytest tests/tier5_adversarial/test_adversarial_stress_perf.py::TestDeepNestedHierarchy tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestCyclicBOMGraphsAndRecursion tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestRevisionVariationsAndNormalization tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestQuantityPrecisionAndRounding::test_european_comma_decimal_handling_in_to_float tests/tier5_adversarial/test_adversarial_stress_perf.py::TestGUIBatchWorkerConcurrency -v
```
*Expected Result*: 14 PASSED, 0 FAILED.
