# Challenger 2 Empirical Challenge Report: Reconciliation, Edge Cases & Fault Injection

**Verdict**: `DEFECTS_DETECTED`  
**Overall Risk Assessment**: `CRITICAL`  
**Date**: 2026-09-17  
**Working Directory**: `D:\Sandbox\pm_sosanhbom\.agents\challenger_2`  
**Parent Conversation ID**: `8a26cf43-3f4f-42ea-ac18-3875de8c9a43`  
**Target Test Suite**: `tests/tier2_boundaries/test_challenger2_empirical_probes.py` (23 tests: 11 passed, 12 failed)

---

## 1. Observation

Direct empirical evidence obtained by inspecting implementation code and running automated verification suites:

### Observation 1.1: Baseline Test Suite Failure (`NameError` in `TC14AutomationClient.login`)
- **Command**: `python -m pytest tests/tier1_features tests/tier2_boundaries tests/tier3_combinations tests/tier4_real_world -q`
- **Result**: `2 failed, 182 passed, 4 warnings in 16.86s`
- **Verbatim Error**:
  ```
  FAILED tests/tier1_features/test_f12_tc14_authentication.py::TestF12TC14Authentication::test_f12_login_form_interaction
  FAILED tests/tier1_features/test_f12_tc14_authentication.py::TestF12TC14Authentication::test_f12_login_timeout_raises_authentication_error
  ...
  src\automation\tc14\client.py:157: in login
      op_timeout = timeout if timeout is not None else self.timeout
  E   NameError: name 'timeout' is not defined
  ```
- **Code Reference** (`src/automation/tc14/client.py`, lines 134-165):
  ```python
  def login(
      self,
      username: str = "vn_pe03",
      password: str = "vn_pe03",
      force: bool = False,
  ) -> bool:
      ...
      driver = self.driver
      op_timeout = timeout if timeout is not None else self.timeout # Line 157: 'timeout' not in parameters!
      logger.info("Navigating to TC14 base URL: %s", self.base_url)
      driver.get(self.base_url)
      try:
          by_user, sel_user = TC14Selectors.USERNAME_INPUT
          username_field = wait.until(EC.presence_of_element_located((by_user, sel_user))) # Line 164: 'wait' uninitialized!
  ```

### Observation 1.2: Silent Decimal Truncation on European / Vietnamese Comma Notation
- **Command**: `python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k test_european_comma_decimal_handling_in_to_float`
- **Result**: `FAILED`
- **Verbatim Error**:
  ```
  AssertionError: Defect: '1,5' was converted to 0.0 instead of 1.5
  assert 0.0 == 1.5
  ```
- **Code Reference** (`src/core/reconciliation.py`, lines 47-55):
  ```python
  def _to_float(val: Any) -> float:
      if val is None or pd.isna(val):
          return 0.0
      try:
          return float(val)
      except (ValueError, TypeError):
          return 0.0
  ```
  `float("1,5")` raises `ValueError`, resulting in silent coercion to `0.0`. In shop-floor Excel files where users input `"1,5"`, the quantity becomes 0.0 and fails reconciliation against PLM/R3 (evaluating to "NG"). In contrast, `src/automation/sap/parser.py:567` properly sanitizes comma decimals (`s = s.replace(",", ".")`).

### Observation 1.3: Revision Normalization Brittleness (Leading Zeros, Casing, Floats, Prefixes)
- **Command**: `python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k TestRevisionVariationsAndNormalization`
- **Result**: `4 failed, 1 passed in 0.44s`
- **Verbatim Error**:
  - `test_leading_zero_revision_matching`: `AssertionError: Defect: Leading-zero revision mismatch: PLM '01' vs R3 '1' produced comp_rev='NG'`
  - `test_case_insensitivity_in_revision_matching`: `AssertionError: Defect: Case-sensitive revision mismatch: 'a' vs 'A' produced comp_rev='NG'`
  - `test_float_representation_of_revision_from_excel`: `AssertionError: Defect: Float-imported revision 1.0 vs '01' produced comp_rev='NG'`
  - `test_slash_and_prefix_revision_notations`: `AssertionError: Defect: Revision notation 'Rev.01' vs '01' failed to match: 'NG'`
- **Code Reference** (`src/core/reconciliation.py`, lines 57-63 and 100-102):
  ```python
  def _normalize_rev(val: Any) -> str:
      if val is None or pd.isna(val):
          return ""
      s = str(val).strip()
      return "" if s.lower() == "nan" else s
  ...
  norm_plm_rev = _normalize_rev(plm_rev)
  norm_r3_rev = _normalize_rev(r3_rev)
  comp_rev = "OK" if norm_plm_rev == norm_r3_rev else "NG"
  ```
  `_normalize_rev` only calls `.strip()`. It does not compare numeric equivalence for numeric revisions (`"01"` vs `"1"`), nor does it case-fold (`"a"` vs `"A"`), nor handle float representations (`1.0` vs `"01"`), nor strip prefix tags (`"Rev.01"` vs `"01"`).

### Observation 1.4: SAP COM Disconnection Blindness & Stale Session Reuse
- **Command**: `python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k TestAdapterFaultInjectionSAPCOM`
- **Result**: `2 failed, 1 passed in 0.45s`
- **Verbatim Error**:
  - `test_sap_com_sudden_disconnection_state_detection`: `AssertionError: Defect: SAPConnectionManager.is_connected() returned True despite disconnected COM session! assert True is False`
  - `test_sap_com_stale_session_recovery_on_get_session`: `AssertionError: Defect: SAPConnectionManager.get_session() returned dead session instead of reconnecting!`
- **Code Reference** (`src/automation/sap/connection.py`, lines 352-358):
  ```python
  def is_connected(self) -> bool:
      return bool(self._connected and (self.session is not None or self.connection is not None))

  def get_session(self) -> Any:
      return self.session or self.get_or_create_session()
  ```
  When the COM server disconnects (e.g. process terminates, RPC error `-2147417848`), `self.session` remains a non-None dead pointer. `is_connected()` continues reporting `True`, and `get_session()` blindly returns the dead session pointer without health validation or reconnection.

### Observation 1.5: TC14 Headless Browser Crash Blindness
- **Command**: `python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k test_tc14_driver_crash_detection_in_is_session_alive`
- **Result**: `FAILED`
- **Verbatim Error**: `AssertionError: Defect: TC14SessionManager.is_session_alive() returned True when browser had crashed! assert True is False`
- **Code Reference** (`src/automation/tc14/session.py`, lines 262-301):
  ```python
  def is_login_page_present(self) -> bool:
      ...
      try:
          current_url = self._driver.current_url.lower() # Raises WebDriverException when browser crashed!
          ...
      except Exception as exc:
          return False # Swallows exception and returns False!

  def is_session_alive(self) -> bool:
      if self._driver is None:
          return False
      if self.is_session_timed_out():
          return False
      if self.is_login_page_present(): # Returns False on crash!
          self.mark_unauthenticated()
          return False
      return self._authenticated # Returns True!
  ```
  If the Chromium browser process dies or is killed, `is_session_alive()` returns `True`, preventing session auto-recovery in downstream workflows.

### Observation 1.6: Cyclic BOM Graphs Cause Infinite Recursion and Process Crash
- **Command**: `python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k TestCyclicBOMGraphsAndRecursion`
- **Result**: `2 failed, 1 passed in 0.52s`
- **Verbatim Error**:
  - `test_date_filter_recursive_infinite_loop_on_cyclic_graph`:
    ```
    src\core\date_filter.py:186: in _filter_node_recursive
        filtered_child = self._filter_node_recursive(child, reference_date)
    E   RecursionError: maximum recursion depth exceeded
    Failed: Defect: DateFilter.filter_tree() entered infinite recursion (RecursionError) on cyclic BOM graph!
    ```
  - `test_model_pruner_recursive_infinite_loop_on_cyclic_graph`:
    ```
    src\core\model_pruner.py:148: in _apply_rule_to_nodes
        node.children = self._apply_rule_to_nodes(node.children, rule)
    E   RecursionError: maximum recursion depth exceeded
    Failed: Defect: ModelPruner.prune_tree() entered infinite recursion (RecursionError) on cyclic BOM graph!
    ```
- **Code Reference**:
  - `src/core/date_filter.py:170-192`: `_filter_node_recursive` lacks a `visited: set[int]` parameter.
  - `src/core/model_pruner.py:121-151`: `_apply_rule_to_nodes` lacks a `visited: set[int]` parameter.
  While `BOMNode.flatten()`, `BOMNode.clone()`, and `UnitResolver._propagate_unit()` implement cycle guards, `DateFilter` and `ModelPruner` omit them, leading to unhandled `RecursionError` and process termination on cyclic structures.

---

## 2. Logic Chain

1. **Premise**: In production, automated comparison engines must be robust against real-world data imperfections (e.g. comma decimals, revision tag differences) and transient external adapter failures (SAP COM disconnects, browser process termination).
2. **Analysis of Observation 1.1 & 1.5**:
   - `TC14AutomationClient.login()` contains fatal syntax/scope errors: `timeout` and `wait` are referenced without definition. Any call to `login()` in live environments crashes with `NameError`.
   - `TC14SessionManager.is_session_alive()` masks browser crashes by swallowing `WebDriverException` in `is_login_page_present()` and returning `True`, defeating session re-establishment.
3. **Analysis of Observation 1.2**:
   - `_to_float("1,5")` converts decimal comma numbers to `0.0`. When production members use comma decimals, valid parts are coerced to 0.0 and erroneously flagged "NG".
4. **Analysis of Observation 1.3**:
   - Teamcenter exports leading zero revisions (`"01"`), while SAP R3 or Excel HTML conversion yields integer-like strings (`"1"`). `_normalize_rev` performs literal string comparison, failing with `"01" != "1" -> comp_rev = "NG"`.
   - Case differences (`"a"` vs `"A"`) and floating point tags (`1.0` vs `"01"`) similarly cause false "NG" reconciliation errors.
5. **Analysis of Observation 1.4**:
   - `SAPConnectionManager` does not check COM connection validity in `is_connected()` or `get_session()`. When SAP Logon crashes or disconnects, the system continues using stale COM pointers and fails transactions without attempting re-acquisition.
6. **Analysis of Observation 1.6**:
   - `DateFilter` and `ModelPruner` recurse without cycle detection. When circular references occur (e.g. recursive assemblies or circular occurrence data), the system enters unbounded recursion and crashes with `RecursionError`.
7. **Deduction**: The core comparison and automation adapter layers contain critical defects in error recovery, data normalization, and graph traversal safety.
8. **Conclusion**: The empirical verdict must be `DEFECTS_DETECTED`.

---

## 3. Caveats

- **No live SAP GUI 770 / AWS connection**: Testing was executed on Windows using mock COM objects and simulated HRESULT disconnection exceptions. Live SAP server behavior was not probed against real SAP Logon AWS instances due to lack of network credentials.
- **No live TC14 Active Workspace endpoint**: Tested via simulated WebDriver sessions and CDP download handlers. Live network latency and server-side session eviction at `http://tcmp3gwb:3000/` were not tested on physical hardware.
- **Review-only enforcement**: Per instructions, no implementation files in `src/` were modified to fix these defects. All findings represent unpatched production code.

---

## 4. Conclusion

**Verdict**: `DEFECTS_DETECTED`

### Summary of Discovered Defects & Recommended Mitigations:

| # | Component | Severity | Defect Description | Recommended Fix |
|---|-----------|----------|--------------------|-----------------|
| **D1** | `src/automation/tc14/client.py:157,164` | **CRITICAL** | `login()` crashes with `NameError: name 'timeout' is not defined` and `wait` is uninitialized. | Add `timeout: Optional[int] = None` to `login()` signature and initialize `wait = WebDriverWait(driver, op_timeout)` before line 164. |
| **D2** | `src/core/date_filter.py:170` | **CRITICAL** | `DateFilter._filter_node_recursive` lacks cycle detection; crashes with `RecursionError` on cyclic BOMs. | Pass `visited: set[int] | None = None` and guard with `if id(node) in visited: return node`. |
| **D3** | `src/core/model_pruner.py:121` | **CRITICAL** | `ModelPruner._apply_rule_to_nodes` lacks cycle detection; crashes with `RecursionError` on cyclic BOMs. | Pass `visited: set[int] | None = None` and guard against cyclic child expansion. |
| **D4** | `src/automation/tc14/session.py:286` | **HIGH** | `is_session_alive()` returns `True` when WebDriver has crashed or closed. | In `is_session_alive()`, verify driver liveness by catching `WebDriverException` and returning `False`. |
| **D5** | `src/automation/sap/connection.py:354,358` | **HIGH** | `is_connected()` returns `True` and `get_session()` returns stale pointer after COM disconnection. | Verify session health with `is_logged_in()` in `is_connected()` and auto-reconnect in `get_session()`. |
| **D6** | `src/core/reconciliation.py:57` | **HIGH** | `_normalize_rev` is string-literal and case-sensitive; fails on `"01"` vs `"1"`, `"a"` vs `"A"`, `1.0` vs `"01"`. | Normalize revisions: uppercase, strip `"Rev."`, and if numeric (e.g. `s.isdigit()`), strip leading zeros. |
| **D7** | `src/core/reconciliation.py:47` | **MEDIUM** | `_to_float("1,5")` coerces comma decimals to `0.0`, triggering false reconciliation "NG". | Replace comma with dot (`val.replace(",", ".")`) before `float()` conversion, matching `ResilientR3Parser`. |
| **D8** | `src/core/tree_parser.py:239` | **MEDIUM** | `PLMTreeParser.parse_file` leaks raw `zipfile.BadZipFile` on 0-byte or corrupted workbooks. | Wrap `openpyxl.load_workbook` in `try..except (BadZipFile, InvalidFileException)` and raise domain error. |

---

## 5. Verification Method

To independently verify and reproduce all 12 defects:

1. **Run the empirical challenge probe suite**:
   ```powershell
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -v
   ```
   *Expected result*: 11 PASSED, 12 FAILED.

2. **Run individual defect verifications**:
   ```powershell
   # D1: TC14 login NameError
   python -m pytest tests/tier1_features/test_f12_tc14_authentication.py::TestF12TC14Authentication::test_f12_login_form_interaction -v

   # D2: DateFilter infinite loop
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestCyclicBOMGraphsAndRecursion::test_date_filter_recursive_infinite_loop_on_cyclic_graph -v

   # D3: ModelPruner infinite loop
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestCyclicBOMGraphsAndRecursion::test_model_pruner_recursive_infinite_loop_on_cyclic_graph -v

   # D4: TC14 driver crash blindness
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestAdapterFaultInjectionTC14BrowserCrash::test_tc14_driver_crash_detection_in_is_session_alive -v

   # D5: SAP COM disconnection blindness
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestAdapterFaultInjectionSAPCOM::test_sap_com_sudden_disconnection_state_detection -v

   # D6: Revision variation failures ("01" vs "1", "a" vs "A")
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestRevisionVariationsAndNormalization::test_leading_zero_revision_matching -v
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestRevisionVariationsAndNormalization::test_case_insensitivity_in_revision_matching -v

   # D7: Comma decimal conversion to 0.0
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestQuantityPrecisionAndRounding::test_european_comma_decimal_handling_in_to_float -v
   ```

3. **Invalidation conditions**:
   The verdict of `DEFECTS_DETECTED` is invalidated only when all 12 probe tests in `tests/tier2_boundaries/test_challenger2_empirical_probes.py` and the 2 tests in `tests/tier1_features/test_f12_tc14_authentication.py` pass at 100% without assertion errors or unhandled exceptions.
