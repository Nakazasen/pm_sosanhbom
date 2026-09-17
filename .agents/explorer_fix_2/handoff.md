# Handoff Report: Automation & Provider Adapters Remediation Investigation

**Agent**: Explorer Fix 2 (Automation & Provider Adapters Remediation Investigator)  
**Roles**: explorer, investigator  
**Working Directory**: `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_2`  
**Parent Conversation ID**: `8a26cf43-3f4f-42ea-ac18-3875de8c9a43`  
**Date**: 2026-09-17  
**Type**: Hard Handoff (Investigation & Architecture Design Completed)  

---

## 1. Observation

Direct empirical evidence obtained from codebase inspection, git state verification, and automated test execution across `src/`, `apps/1.0.0/`, and `tests/`:

### 1.1 TC14 Client Scoping & Parameter Discrepancy Between `src/` and `apps/1.0.0/`
- **Location**: `src/automation/tc14/client.py` vs `apps/1.0.0/src/automation/tc14/client.py`
- In `src/automation/tc14/client.py` (lines 133-164), `login()` was updated to include `timeout: Optional[int] = None` and `wait = WebDriverWait(driver, op_timeout)`.
- However, in `apps/1.0.0/src/automation/tc14/client.py` (lines 133-165), the code remains **unpatched and broken**:
  ```python
  # File: apps/1.0.0/src/automation/tc14/client.py (lines 133-164)
      def login(
          self,
          username: str = "vn_pe03",
          password: str = "vn_pe03",
          force: bool = False,
      ) -> bool:
  ...
          driver = self.driver
          op_timeout = timeout if timeout is not None else self.timeout
          logger.info("Navigating to TC14 base URL: %s", self.base_url)
          driver.get(self.base_url)

          try:
              # Locate username input
              by_user, sel_user = TC14Selectors.USERNAME_INPUT
              username_field = wait.until(EC.presence_of_element_located((by_user, sel_user)))
  ```
- **Consequence**: When packaged into the MP2027 portable release bundle, `apps/1.0.0` will crash with `NameError: name 'timeout' is not defined` whenever TC14 login is invoked.

---

### 1.2 Total Omission of Requirement R5: Flexible Provider Adapter Pattern
- **Project Index**: `ORIGINAL_REQUEST.md`, Section R5 ("Kiến Trúc Mở Rộng Linh Hoạt (Adapter Pattern)").
- Explicit requirement:
  - Core BOM comparison and tree traversal algorithms must be isolated from data collection mechanisms.
  - `PLMProvider` ABC with interchangeable implementations: `TeamcenterSeleniumAdapter` and `ExcelPLMAdapter`.
  - `ERPProvider` ABC with interchangeable implementations: `SAPR3COMAdapter` and `ExcelR3Adapter`.
  - Acceptance Criteria: *"Kiểm thử tính độc lập của Core Engine khi thay đổi Provider đầu vào (Adapter pattern)."*
- **Current Codebase Inspection**:
  - A project-wide ripgrep search (`grep_search`) for `PLMProvider`, `ERPProvider`, `TeamcenterSeleniumAdapter`, `ExcelPLMAdapter`, `SAPR3COMAdapter`, and `ExcelR3Adapter` returned **0 matches**.
  - `src/core/` contains no `adapters.py` module.
  - `tests/unit/` contains no tests validating adapter interchangeability.

---

### 1.3 Empirical Reproduction of Adapter Fault Blindness
- **Command Executed**:
  ```powershell
  python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k "test_sap_com_sudden_disconnection_state_detection or test_sap_com_stale_session_recovery_on_get_session or test_tc14_driver_crash_detection_in_is_session_alive" -v
  ```
- **Execution Result**: `3 failed, 20 deselected in 3.43s`
- **Failure 1 (SAP COM Disconnection Blindness)**:
  ```text
  FAILED tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestAdapterFaultInjectionSAPCOM::test_sap_com_sudden_disconnection_state_detection
  AssertionError: Defect: SAPConnectionManager.is_connected() returned True despite disconnected COM session!
  assert True is False
  ```
  *Root Cause* (`src/automation/sap/connection.py:352-354`):
  ```python
  def is_connected(self) -> bool:
      """Check if an active connection or session exists."""
      return bool(self._connected and (self.session is not None or self.connection is not None))
  ```
  When COM severance occurs (HRESULT `-2147417848`), `self.session` remains a dead pointer; `is_connected()` returns `True`.

- **Failure 2 (Stale Session Reuse Without Health-Check)**:
  ```text
  FAILED tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestAdapterFaultInjectionSAPCOM::test_sap_com_stale_session_recovery_on_get_session
  AssertionError: Defect: SAPConnectionManager.get_session() returned dead session instead of reconnecting!
  assert <MagicMock id='1811679275264'> == <MagicMock name='get_or_create_session()' id='1811679277952'>
  ```
  *Root Cause* (`src/automation/sap/connection.py:356-358`):
  ```python
  def get_session(self) -> Any:
      """Return existing session or acquire new session."""
      return self.session or self.get_or_create_session()
  ```
  `get_session()` never probes session health and returns dead pointers without reconnecting.

- **Failure 3 (TC14 Browser Crash Blindness)**:
  ```text
  FAILED tests/tier2_boundaries/test_challenger2_empirical_probes.py::TestAdapterFaultInjectionTC14BrowserCrash::test_tc14_driver_crash_detection_in_is_session_alive
  AssertionError: Defect: TC14SessionManager.is_session_alive() returned True when browser had crashed!
  assert True is False
  ```
  *Root Cause* (`src/automation/tc14/session.py:270-301`):
  ```python
  def is_login_page_present(self) -> bool:
      try:
          current_url = self._driver.current_url.lower()
          ...
      except Exception as exc:
          return False  # Catches WebDriverException when browser crashed!

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

---

### 1.4 Non-Existent `SAPBOMParser` Import in `LeaderWorkspaceView` and `MemberWorkspaceView`
- **Location 1**: `src/gui/leader_view.py:98-106` (and `apps/1.0.0/src/gui/leader_view.py:98-106`)
  ```python
  99:  from src.automation.sap.parser import SAPBOMParser
  100: df_r3 = SAPBOMParser.parse_sap_export(self.r3_path)
  ```
- **Location 2**: `src/gui/member_view.py:434-445` (and `apps/1.0.0/src/gui/member_view.py:434-445`)
  ```python
  436: from src.automation.sap.parser import SAPBOMParser
  437: self.r3_data = SAPBOMParser.parse_sap_export(Path(r3_path))
  ```
- **Codebase Reality**: In `src/automation/sap/parser.py`, the parser class is named `ResilientR3Parser` (with module-level convenience function `parse_r3_cs12_file(filepath)`). `SAPBOMParser` does not exist. Both files consistently throw `ImportError`, silently caught by bare `except Exception:`, falling back to `pd.read_excel()` which fails to parse SAP ALV HTML/TSV formats.

---

## 2. Logic Chain

1. **Premise 1 (Spec Compliance - Requirement R5)**:
   The core engine must be completely insulated from I/O mechanisms. It must be possible to run BOM reconciliation against live enterprise backends (Active Workspace + SAP GUI Scripting) or entirely offline from local files (Excel PLM + Excel R3) by swapping pluggable provider adapters without modifying a single line of `ReconciliationEngine`, `DateFilter`, `ModelPruner`, `UnitResolver`, or `MSIEngine`.
2. **Premise 2 (Production Resilience & Fault Tolerance)**:
   In shop-floor production, network hiccups, socket timeouts, browser crashes (e.g. Chrome OOM), and SAP GUI COM severance (HRESULT `-2147417848`) are expected events. Adapters and session managers must detect severance immediately (`fail-closed`), discard dead pointers, and re-establish sessions cleanly without throwing unhandled exceptions to the user.
3. **Premise 3 (Bundle Parity)**:
   Any bug fix applied to `src/` must also be synchronized to `apps/1.0.0/src/` to prevent launcher execution of unpatched code.
4. **Deduction**:
   - `src/automation/tc14/client.py` and `apps/1.0.0/src/automation/tc14/client.py` must share the exact same clean signature `login(..., timeout: Optional[int] = None)` and initialize `wait = WebDriverWait(driver, op_timeout)`.
   - A dedicated `src/core/adapters.py` module must be created with full type annotations, docstrings, and concrete classes `PLMProvider`, `TeamcenterSeleniumAdapter`, `ExcelPLMAdapter`, `ERPProvider`, `SAPR3COMAdapter`, and `ExcelR3Adapter`.
   - `TC14SessionManager.is_session_alive()` must probe driver responsiveness and return `False` on `WebDriverException`. `SAPConnectionManager.is_connected()` must verify `is_logged_in()` and `get_session()` must auto-reconnect if stale.
   - `leader_view.py` and `member_view.py` must import `parse_r3_cs12_file` directly, while `src/automation/sap/parser.py` must export a backward-compatible `SAPBOMParser` facade.

---

## 3. Caveats

- **External Network Access**: In isolated/offline environments without access to `http://tcmp3gwb:3000/` or live SAP GUI 770 AWS instances, `TeamcenterSeleniumAdapter` and `SAPR3COMAdapter` should be tested using mock objects and simulated driver/COM fixtures. `ExcelPLMAdapter` and `ExcelR3Adapter` provide 100% offline functionality.
- **Lazy Imports**: `TeamcenterSeleniumAdapter` and `SAPR3COMAdapter` must use deferred/lazy imports of Selenium and win32com to ensure `src/core/adapters.py` can be imported in environments where Selenium or Windows COM bindings are absent or restricted.
- **App Version Mirroring**: When implementing, the implementer must ensure both `src/` and `apps/1.0.0/src/` are synchronized, or re-run `python scripts/package_app.py` after patching `src/`.

---

## 4. Conclusion & Complete Remediation Code

### Summary Table of Required Changes

| # | File Path | Defect Addressed | Action Required |
|---|---|---|---|
| **1** | `src/automation/tc14/client.py` & `apps/1.0.0/src/automation/tc14/client.py` | NameError `timeout` and uninitialized `wait` | Add `timeout: Optional[int] = None` and `wait = WebDriverWait(driver, op_timeout)` before line 164 |
| **2** | `src/core/adapters.py` (NEW) & `apps/1.0.0/src/core/adapters.py` | Missing R5 Provider Adapter pattern | Create `PLMProvider`, `TeamcenterSeleniumAdapter`, `ExcelPLMAdapter`, `ERPProvider`, `SAPR3COMAdapter`, `ExcelR3Adapter` |
| **3** | `src/core/__init__.py` & `apps/1.0.0/src/core/__init__.py` | Adapters not exported from core | Export all 6 provider classes in `__all__` |
| **4** | `tests/unit/test_adapters.py` (NEW) | Missing adapter test suite | Comprehensive unit tests validating provider swapping and core engine independence |
| **5** | `src/automation/tc14/session.py` & `apps/1.0.0/src/automation/tc14/session.py` | Browser crash blindness in `is_session_alive()` | Probe `current_url`, catch `WebDriverException`, mark unauthenticated, add `ensure_driver_alive()` |
| **6** | `src/automation/sap/connection.py` & `apps/1.0.0/src/automation/sap/connection.py` | COM disconnection blindness in `is_connected()` & `get_session()` | Check `is_logged_in()` in `is_connected()`; auto-reconnect in `get_session()` |
| **7** | `src/automation/sap/parser.py` & `apps/1.0.0/src/automation/sap/parser.py` | Missing `SAPBOMParser` class | Add `SAPBOMParser` compatibility class delegating to `ResilientR3Parser` |
| **8** | `src/gui/leader_view.py` & `apps/1.0.0/src/gui/leader_view.py` | Broken import at line 99 | Import `parse_r3_cs12_file` directly from `src.automation.sap.parser` |
| **9** | `src/gui/member_view.py` & `apps/1.0.0/src/gui/member_view.py` | Broken import at line 436 | Import `parse_r3_cs12_file` directly from `src.automation.sap.parser` |

---

### Concrete Implementation Specifications (Drop-In Code)

#### 4.1 Remediation Code for `src/automation/tc14/client.py` and `apps/1.0.0/src/automation/tc14/client.py`

Replace `login()` method (lines 133-181) with:

```python
    def login(
        self,
        username: str = "vn_pe03",
        password: str = "vn_pe03",
        force: bool = False,
        timeout: Optional[int] = None,
    ) -> bool:
        """Authenticate with Teamcenter Active Workspace.

        Args:
            username: Login user account (default 'vn_pe03').
            password: Login password (default 'vn_pe03').
            force: Force fresh login even if session appears active.
            timeout: Optional timeout override in seconds.

        Returns:
            True if login succeeded.

        Raises:
            TC14AuthenticationError: If login fails or credentials invalid.
        """
        if not force and self.session.is_session_alive():
            logger.info("Session already active and authenticated.")
            return True

        driver = self.driver
        op_timeout = timeout if timeout is not None else self.timeout
        wait = WebDriverWait(driver, op_timeout)
        logger.info("Navigating to TC14 base URL: %s", self.base_url)
        driver.get(self.base_url)

        try:
            # Locate username input
            by_user, sel_user = TC14Selectors.USERNAME_INPUT
            username_field = wait.until(EC.presence_of_element_located((by_user, sel_user)))
            username_field.clear()
            username_field.send_keys(username)

            # Locate password input
            by_pass, sel_pass = TC14Selectors.PASSWORD_INPUT
            password_field = driver.find_element(by_pass, sel_pass)
            password_field.clear()
            password_field.send_keys(password)

            # Allow React state to register input changes
            time.sleep(0.5)

            # Click login button
            by_btn, sel_btn = TC14Selectors.LOGIN_BUTTON
            login_btn = wait.until(EC.element_to_be_clickable((by_btn, sel_btn)))
            login_btn.click()
```

---

#### 4.2 Remediation Code for `src/core/adapters.py` (Complete New Module)

Create `src/core/adapters.py` (and copy to `apps/1.0.0/src/core/adapters.py`):

```python
"""R5 Flexible Provider Adapter Architecture.

Decouples the core comparison engine and tree algorithms from concrete data sources
(Selenium web automation vs local Excel files, SAP GUI COM automation vs exported TSV/HTML/XLS).

Enables hot-swapping between:
- PLMProvider: TeamcenterSeleniumAdapter <-> ExcelPLMAdapter
- ERPProvider: SAPR3COMAdapter <-> ExcelR3Adapter
"""

from __future__ import annotations

import abc
import datetime
import logging
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd

from src.core.models import BOMTree
from src.core.tree_parser import PLMTreeParser
from src.automation.sap.parser import ResilientR3Parser

logger = logging.getLogger(__name__)


# =============================================================================
# 1. PLM Provider Interface & Implementations
# =============================================================================

class PLMProvider(abc.ABC):
    """Abstract Base Class for PLM (Product Lifecycle Management) data providers."""

    @abc.abstractmethod
    def fetch_bom(self, item_id: str, rev: Optional[str] = None) -> BOMTree:
        """Fetch and parse the full BOM hierarchy for an item.

        Args:
            item_id: Part code / Item ID (e.g. '110C103NL0', 'Virgo', etc.).
            rev: Revision identifier (e.g. '01', 'A', etc.).

        Returns:
            Structured in-memory BOMTree.
        """
        pass

    @abc.abstractmethod
    def export_excel(
        self,
        item_id: str,
        rev: Optional[str] = None,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Path:
        """Export the BOM structure to an Excel file on disk.

        Args:
            item_id: Part code / Item ID.
            rev: Optional revision identifier.
            output_path: Target output path (file or directory).

        Returns:
            Path to the saved Excel file.
        """
        pass


class TeamcenterSeleniumAdapter(PLMProvider):
    """PLM Adapter wrapping TC14AutomationClient for live Active Workspace extraction."""

    def __init__(
        self,
        client: Optional[Any] = None,
        tree_parser: Optional[PLMTreeParser] = None,
    ) -> None:
        """Initialize adapter with TC14AutomationClient instance."""
        if client is None:
            # Lazy import to avoid selenium dependency unless adapter is used
            from src.automation.tc14.client import TC14AutomationClient
            self.client = TC14AutomationClient()
        else:
            self.client = client
        self.parser = tree_parser or PLMTreeParser()

    def export_excel(
        self,
        item_id: str,
        rev: Optional[str] = None,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Path:
        """Export BOM Excel via TC14AutomationClient."""
        target_dir = (
            Path(output_path).parent
            if output_path and Path(output_path).suffix
            else (Path(output_path) if output_path else Path("./downloads"))
        )
        target_dir.mkdir(parents=True, exist_ok=True)

        exported_file = self.client.download_bom_full(
            part_number=item_id,
            output_dir=target_dir,
            part_rev=rev,
        )
        if output_path and Path(output_path).suffix and Path(exported_file) != Path(output_path):
            shutil.copy2(exported_file, output_path)
            return Path(output_path)
        return Path(exported_file)

    def fetch_bom(self, item_id: str, rev: Optional[str] = None) -> BOMTree:
        """Export BOM from Active Workspace and parse into a BOMTree."""
        temp_excel = self.export_excel(item_id, rev=rev)
        return self.parser.parse_file(temp_excel)


class ExcelPLMAdapter(PLMProvider):
    """PLM Adapter for reading local PLM Excel files without network or browser dependencies."""

    def __init__(
        self,
        file_map: Optional[Dict[str, Union[str, Path]]] = None,
        default_file: Optional[Union[str, Path]] = None,
        search_dir: Optional[Union[str, Path]] = None,
        tree_parser: Optional[PLMTreeParser] = None,
    ) -> None:
        """Initialize ExcelPLMAdapter.

        Args:
            file_map: Dictionary mapping item_id (or item_id:rev) to specific file paths.
            default_file: Fallback file path if item_id not in file_map.
            search_dir: Directory to search for matching files by name.
            tree_parser: Custom or default PLMTreeParser instance.
        """
        self.file_map = {k: Path(v) for k, v in (file_map or {}).items()}
        self.default_file = Path(default_file) if default_file else None
        self.search_dir = Path(search_dir) if search_dir else None
        self.parser = tree_parser or PLMTreeParser()

    def _resolve_file(self, item_id: str, rev: Optional[str] = None) -> Path:
        """Resolve file path for given item_id and rev."""
        key_with_rev = f"{item_id}:{rev}" if rev else item_id
        if key_with_rev in self.file_map and self.file_map[key_with_rev].exists():
            return self.file_map[key_with_rev]
        if item_id in self.file_map and self.file_map[item_id].exists():
            return self.file_map[item_id]

        if self.search_dir and self.search_dir.exists():
            candidates = list(self.search_dir.glob(f"*{item_id}*.xls*"))
            if candidates:
                return candidates[0]

        if self.default_file and self.default_file.exists():
            return self.default_file

        raise FileNotFoundError(
            f"ExcelPLMAdapter: Could not locate PLM Excel file for item_id='{item_id}', rev='{rev}'."
        )

    def fetch_bom(self, item_id: str, rev: Optional[str] = None) -> BOMTree:
        """Read and parse local PLM Excel file into a BOMTree."""
        file_path = self._resolve_file(item_id, rev=rev)
        return self.parser.parse_file(file_path)

    def export_excel(
        self,
        item_id: str,
        rev: Optional[str] = None,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Path:
        """Return path to existing Excel file, copying to output_path if requested."""
        src_path = self._resolve_file(item_id, rev=rev)
        if output_path is None:
            return src_path

        dest = Path(output_path)
        if dest.is_dir():
            dest = dest / src_path.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.resolve() != src_path.resolve():
            shutil.copy2(src_path, dest)
        return dest


# =============================================================================
# 2. ERP Provider Interface & Implementations
# =============================================================================

class ERPProvider(abc.ABC):
    """Abstract Base Class for ERP (Enterprise Resource Planning) data providers."""

    @abc.abstractmethod
    def fetch_multilevel_bom(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> pd.DataFrame:
        """Fetch and return multilevel BOM data as a normalized pandas DataFrame.

        DataFrame guaranteed columns:
            ['part_code', 'quantity', 'rev_r3']
        """
        pass

    @abc.abstractmethod
    def export_multilevel_bom_file(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """Export CS12 multilevel BOM to a file on disk.

        Returns:
            Path to exported file.
        """
        pass


class SAPR3COMAdapter(ERPProvider):
    """ERP Adapter wrapping CS12Service via SAP GUI Scripting COM interface."""

    def __init__(
        self,
        cs12_service: Optional[Any] = None,
        parser: Optional[ResilientR3Parser] = None,
    ) -> None:
        """Initialize SAPR3COMAdapter."""
        if cs12_service is None:
            from src.automation.sap.cs12 import CS12Service
            self.cs12_service = CS12Service()
        else:
            self.cs12_service = cs12_service
        self.parser = parser or ResilientR3Parser()

    def export_multilevel_bom_file(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """Execute CS12 transaction via SAP COM and export result file."""
        dest = Path(destination_dir or "./downloads")
        dest.mkdir(parents=True, exist_ok=True)

        date_val = valid_date
        if isinstance(date_val, str):
            date_val = datetime.date.fromisoformat(date_val)
        elif date_val is None:
            date_val = datetime.date.today()

        result = self.cs12_service.execute_cs12(
            material=material,
            plant=plant,
            valid_date=date_val,
            bom_usage=bom_usage,
            alternative=alternative,
            destination_dir=dest,
        )
        if not result.success or not result.file_path:
            raise RuntimeError(f"SAP CS12 execution failed: {result.error_message}")

        return Path(result.file_path)

    def fetch_multilevel_bom(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> pd.DataFrame:
        """Export CS12 via COM and parse into a normalized DataFrame."""
        file_path = self.export_multilevel_bom_file(
            material=material,
            plant=plant,
            valid_date=valid_date,
            bom_usage=bom_usage,
            alternative=alternative,
        )
        return self.parser.parse(file_path)


class ExcelR3Adapter(ERPProvider):
    """ERP Adapter reading local SAP R3 CS12 export files without SAP GUI or COM."""

    def __init__(
        self,
        file_map: Optional[Dict[str, Union[str, Path]]] = None,
        default_file: Optional[Union[str, Path]] = None,
        search_dir: Optional[Union[str, Path]] = None,
        parser: Optional[ResilientR3Parser] = None,
    ) -> None:
        """Initialize ExcelR3Adapter."""
        self.file_map = {k: Path(v) for k, v in (file_map or {}).items()}
        self.default_file = Path(default_file) if default_file else None
        self.search_dir = Path(search_dir) if search_dir else None
        self.parser = parser or ResilientR3Parser()

    def _resolve_file(self, material: str) -> Path:
        """Resolve file path for given material."""
        if material in self.file_map and self.file_map[material].exists():
            return self.file_map[material]

        if self.search_dir and self.search_dir.exists():
            candidates = list(self.search_dir.glob(f"*{material}*.xls*"))
            if candidates:
                return candidates[0]

        if self.default_file and self.default_file.exists():
            return self.default_file

        raise FileNotFoundError(
            f"ExcelR3Adapter: Could not locate SAP R3 export file for material='{material}'."
        )

    def fetch_multilevel_bom(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> pd.DataFrame:
        """Read and parse local SAP R3 file into a normalized DataFrame."""
        file_path = self._resolve_file(material)
        return self.parser.parse(file_path)

    def export_multilevel_bom_file(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """Return path to existing R3 export file, copying if destination requested."""
        src_path = self._resolve_file(material)
        if destination_dir is None:
            return src_path

        dest = Path(destination_dir)
        if dest.is_dir():
            dest = dest / src_path.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.resolve() != src_path.resolve():
            shutil.copy2(src_path, dest)
        return dest
```

Update `src/core/__init__.py` to export the adapters:
```python
from src.core.adapters import (
    ERPProvider,
    ExcelPLMAdapter,
    ExcelR3Adapter,
    PLMProvider,
    SAPR3COMAdapter,
    TeamcenterSeleniumAdapter,
)

__all__ = [
    ...,
    "PLMProvider",
    "TeamcenterSeleniumAdapter",
    "ExcelPLMAdapter",
    "ERPProvider",
    "SAPR3COMAdapter",
    "ExcelR3Adapter",
]
```

---

#### 4.3 Unit Test Suite for Adapters (`tests/unit/test_adapters.py`)

Create `tests/unit/test_adapters.py`:

```python
"""Unit tests for R5 Provider Adapter pattern.

Verifies:
1. Abstract base class contract enforcement (PLMProvider, ERPProvider).
2. ExcelPLMAdapter file resolution, BOMTree parsing, and export.
3. TeamcenterSeleniumAdapter wrapping TC14AutomationClient.
4. ExcelR3Adapter file resolution, DataFrame parsing, and export.
5. SAPR3COMAdapter wrapping CS12Service.
6. Complete Core Engine Independence: swapping providers produces 100% bit-accurate
   reconciliation results without touching ReconciliationEngine.
"""

import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch
import pandas as pd
import pytest

from src.core.adapters import (
    ERPProvider,
    ExcelPLMAdapter,
    ExcelR3Adapter,
    PLMProvider,
    SAPR3COMAdapter,
    TeamcenterSeleniumAdapter,
)
from src.core.models import BOMNode, BOMTree
from src.core.reconciliation import ReconciliationEngine


class TestProviderContracts:
    """Validate ABC instantiation prevention and interface adherence."""

    def test_plm_provider_abc_cannot_be_instantiated(self):
        with pytest.raises(TypeError):
            PLMProvider()

    def test_erp_provider_abc_cannot_be_instantiated(self):
        with pytest.raises(TypeError):
            ERPProvider()


class TestExcelPLMAdapter:
    """Test offline ExcelPLMAdapter functionality."""

    def test_resolve_file_from_file_map(self, tmp_path: Path):
        sample_file = tmp_path / "sample_plm.xlsx"
        sample_file.write_text("dummy")

        adapter = ExcelPLMAdapter(file_map={"110C103NL0": sample_file})
        resolved = adapter._resolve_file("110C103NL0")
        assert resolved == sample_file

    def test_resolve_file_missing_raises_filenotfound(self):
        adapter = ExcelPLMAdapter()
        with pytest.raises(FileNotFoundError):
            adapter._resolve_file("NON_EXISTENT")

    def test_fetch_bom_delegates_to_parser(self, tmp_path: Path):
        sample_file = tmp_path / "sample_plm.xlsx"
        sample_file.write_text("dummy")

        mock_parser = MagicMock()
        expected_tree = BOMTree(roots=[BOMNode(level=1, item_id="PART_1", quantity=1.0)])
        mock_parser.parse_file.return_value = expected_tree

        adapter = ExcelPLMAdapter(default_file=sample_file, tree_parser=mock_parser)
        tree = adapter.fetch_bom("PART_1")
        assert tree == expected_tree
        mock_parser.parse_file.assert_called_once_with(sample_file)


class TestTeamcenterSeleniumAdapter:
    """Test TeamcenterSeleniumAdapter wrapping TC14AutomationClient."""

    def test_export_excel_calls_download_bom_full(self, tmp_path: Path):
        mock_client = MagicMock()
        mock_client.download_bom_full.return_value = tmp_path / "exported.xlsx"
        (tmp_path / "exported.xlsx").write_text("data")

        adapter = TeamcenterSeleniumAdapter(client=mock_client)
        out = adapter.export_excel("110C103NL0", rev="A", output_path=tmp_path / "target.xlsx")

        assert out.exists()
        mock_client.download_bom_full.assert_called_once()

    def test_fetch_bom_parses_exported_file(self, tmp_path: Path):
        mock_client = MagicMock()
        temp_file = tmp_path / "temp.xlsx"
        temp_file.write_text("data")
        mock_client.download_bom_full.return_value = temp_file

        mock_parser = MagicMock()
        expected_tree = BOMTree(roots=[BOMNode(level=1, item_id="TC_PART", quantity=2.0)])
        mock_parser.parse_file.return_value = expected_tree

        adapter = TeamcenterSeleniumAdapter(client=mock_client, tree_parser=mock_parser)
        tree = adapter.fetch_bom("TC_PART")
        assert tree == expected_tree


class TestExcelR3Adapter:
    """Test offline ExcelR3Adapter functionality."""

    def test_fetch_multilevel_bom_delegates_to_parser(self, tmp_path: Path):
        sample_r3 = tmp_path / "sample_r3.xls"
        sample_r3.write_text("dummy")

        mock_parser = MagicMock()
        expected_df = pd.DataFrame([{"part_code": "PART_R3", "quantity": 3.0, "rev_r3": "01"}])
        mock_parser.parse.return_value = expected_df

        adapter = ExcelR3Adapter(default_file=sample_r3, parser=mock_parser)
        df = adapter.fetch_multilevel_bom("110C103NL0")
        assert df.equals(expected_df)
        mock_parser.parse.assert_called_once_with(sample_r3)


class TestSAPR3COMAdapter:
    """Test SAPR3COMAdapter wrapping CS12Service."""

    def test_fetch_multilevel_bom_executes_cs12_and_parses(self, tmp_path: Path):
        mock_cs12 = MagicMock()
        exported_file = tmp_path / "cs12_export.xls"
        exported_file.write_text("export data")

        mock_result = MagicMock()
        mock_result.success = True
        mock_result.file_path = exported_file
        mock_cs12.execute_cs12.return_value = mock_result

        mock_parser = MagicMock()
        expected_df = pd.DataFrame([{"part_code": "MAT_01", "quantity": 10.0, "rev_r3": "00"}])
        mock_parser.parse.return_value = expected_df

        adapter = SAPR3COMAdapter(cs12_service=mock_cs12, parser=mock_parser)
        df = adapter.fetch_multilevel_bom("MAT_01")
        assert df.equals(expected_df)
        mock_cs12.execute_cs12.assert_called_once()
        mock_parser.parse.assert_called_once_with(exported_file)


class TestCoreEngineIndependenceViaAdapterSwapping:
    """Verify core ReconciliationEngine operates identically regardless of swapped provider."""

    def test_engine_independence_with_swapped_providers(self, tmp_path: Path):
        # 1. Setup sample datasets
        cttt_data = pd.DataFrame([
            {"part_code": "302FP02010", "quantity": 2.0, "usage": "Unit-1", "rev": "01"}
        ])
        plm_df = pd.DataFrame([
            {"part_code": "302FP02010", "quantity": 2.0, "level": 1, "rev_plm": "01"}
        ])
        r3_df = pd.DataFrame([
            {"part_code": "302FP02010", "quantity": 2.0, "rev_r3": "01"}
        ])

        # Provider Pair A: Offline Excel Adapters
        plm_file = tmp_path / "plm_test.xlsx"
        plm_file.write_text("dummy")
        r3_file = tmp_path / "r3_test.xls"
        r3_file.write_text("dummy")

        mock_tree = MagicMock()
        mock_tree.to_dataframe.return_value = plm_df

        mock_plm_parser = MagicMock()
        mock_plm_parser.parse_file.return_value = mock_tree

        mock_r3_parser = MagicMock()
        mock_r3_parser.parse.return_value = r3_df

        excel_plm = ExcelPLMAdapter(default_file=plm_file, tree_parser=mock_plm_parser)
        excel_r3 = ExcelR3Adapter(default_file=r3_file, parser=mock_r3_parser)

        # Provider Pair B: Automated Selenium / COM Adapters (Mocked)
        mock_tc14_client = MagicMock()
        mock_tc14_client.download_bom_full.return_value = plm_file
        selenium_plm = TeamcenterSeleniumAdapter(client=mock_tc14_client, tree_parser=mock_plm_parser)

        mock_cs12_service = MagicMock()
        mock_cs12_res = MagicMock()
        mock_cs12_res.success = True
        mock_cs12_res.file_path = r3_file
        mock_cs12_service.execute_cs12.return_value = mock_cs12_res
        com_r3 = SAPR3COMAdapter(cs12_service=mock_cs12_service, parser=mock_r3_parser)

        # Execute Engine with Provider Pair A
        engine = ReconciliationEngine()
        tree_a = excel_plm.fetch_bom("VIRGO")
        df_r3_a = excel_r3.fetch_multilevel_bom("VIRGO")
        result_a = engine.run_full_reconciliation(
            cttt_data=cttt_data,
            plm_data=tree_a.to_dataframe(),
            r3_data=df_r3_a,
        )

        # Execute Engine with Provider Pair B
        tree_b = selenium_plm.fetch_bom("VIRGO")
        df_r3_b = com_r3.fetch_multilevel_bom("VIRGO")
        result_b = engine.run_full_reconciliation(
            cttt_data=cttt_data,
            plm_data=tree_b.to_dataframe(),
            r3_data=df_r3_b,
        )

        # Verify 100% bit-accurate identity across providers
        assert result_a.overall_status == result_b.overall_status == "OK"
        assert result_a.summary_statistics == result_b.summary_statistics
        assert len(result_a.reconciled_rows) == len(result_b.reconciled_rows) == 1
        assert result_a.reconciled_rows[0].comp_qty == result_b.reconciled_rows[0].comp_qty == "OK"
        assert result_a.reconciled_rows[0].comp_rev == result_b.reconciled_rows[0].comp_rev == "OK"
```

---

#### 4.4 Remediation Code for Adapter Fault Blindness & Reconnection

##### File 1: `src/automation/tc14/session.py` (and `apps/1.0.0/src/automation/tc14/session.py`)

Replace `is_session_alive()` (lines 286-302) with:

```python
    def is_session_alive(self) -> bool:
        """Verify whether the browser session is currently active and authenticated.

        Returns False if the driver is dead, timed out, crashed, or showing the login screen.
        """
        if self._driver is None:
            return False

        if self.is_session_timed_out():
            return False

        # Direct liveness probe: catching browser crash, killed process, or closed socket
        try:
            _ = self._driver.current_url
        except WebDriverException as exc:
            logger.warning("WebDriver connection severed or browser crashed (%s). Marking unauthenticated.", exc)
            self.mark_unauthenticated()
            return False
        except Exception:
            self.mark_unauthenticated()
            return False

        if self.is_login_page_present():
            self.mark_unauthenticated()
            return False

        return self._authenticated

    def ensure_driver_alive(self) -> WebDriver:
        """Verify driver health; terminate dead process and reinitialize cleanly if crashed."""
        if self._driver is not None:
            try:
                _ = self._driver.current_url
                return self._driver
            except Exception:
                logger.warning("Dead driver detected in ensure_driver_alive. Re-initializing...")
                self.close()
        return self.start_session()
```

##### File 2: `src/automation/sap/connection.py` (and `apps/1.0.0/src/automation/sap/connection.py`)

Replace `is_connected()` and `get_session()` (lines 352-359) with:

```python
    def is_connected(self) -> bool:
        """Check if an active connection or session exists and is responsive."""
        if not self._connected:
            return False
        if self.session is None and self.connection is None:
            return False
        if self.session is not None:
            # Active liveness health check via is_logged_in
            try:
                return bool(self.is_logged_in(self.session))
            except Exception:
                return False
        if self.connection is not None:
            try:
                _ = getattr(self.connection, "Children", None)
                return True
            except Exception:
                return False
        return False

    def get_session(self) -> Any:
        """Return existing healthy session or automatically reconnect if severed/stale."""
        if self.session is not None:
            try:
                if self.is_logged_in(self.session):
                    return self.session
                logger.warning("Existing SAP session is not logged in or unresponsive; re-acquiring...")
            except Exception as ex:
                logger.warning("Existing SAP session threw error on health-check (%s); reconnecting...", ex)
            self.session = None

        return self.get_or_create_session()
```

---

#### 4.5 Remediation Code for `SAPBOMParser` Import Bug

##### File 1: `src/automation/sap/parser.py` (and `apps/1.0.0/src/automation/sap/parser.py`)

Append backward-compatibility alias at the end of `parser.py`:

```python
# Backward compatibility alias for legacy imports
class SAPBOMParser:
    """Compatibility wrapper for ResilientR3Parser."""

    @staticmethod
    def parse_sap_export(filepath: Union[str, Path]) -> pd.DataFrame:
        """Parse SAP export file using ResilientR3Parser."""
        return ResilientR3Parser().parse(filepath)
```

##### File 2: `src/gui/leader_view.py:98-106` (and `apps/1.0.0/src/gui/leader_view.py:98-106`)

Replace lines 98-106 with:

```python
            self.progress.emit(55, "Đang nạp cấu trúc BOM SAP R3 CS12...")
            df_r3 = pd.DataFrame()
            if self.r3_path and self.r3_path.exists():
                try:
                    from src.automation.sap.parser import parse_r3_cs12_file
                    df_r3 = parse_r3_cs12_file(self.r3_path)
                except Exception as ex:
                    logger.warning("Error parsing R3 file with ResilientR3Parser: %s; trying Excel fallback", ex)
                    try:
                        df_r3 = pd.read_excel(self.r3_path)
                    except Exception as ex2:
                        logger.warning("Error reading R3 file: %s", ex2)
```

##### File 3: `src/gui/member_view.py:434-445` (and `apps/1.0.0/src/gui/member_view.py:434-445`)

Replace lines 434-445 with:

```python
        if r3_path:
            try:
                # SAP exports might be HTML table format or Excel
                from src.automation.sap.parser import parse_r3_cs12_file
                self.r3_data = parse_r3_cs12_file(Path(r3_path))
                self.lbl_ref_status.setText(f"Đã nạp PLM & R3 ({Path(r3_path).name})")
            except Exception as ex:
                logger.warning("Error parsing R3 file: %s; trying Excel fallback", ex)
                try:
                    self.r3_data = pd.read_excel(r3_path)
                    self.lbl_ref_status.setText("Đã nạp PLM & R3")
                except Exception as exc:
                    QMessageBox.warning(self, "Lỗi nạp R3", f"Không thể đọc tệp R3: {exc}")
```

---

## 5. Verification Method

To independently verify the implementation and confirm full remediation:

1. **Verify TC14 Authentication & Parameter Scoping Fix**:
   ```powershell
   python -m pytest tests/tier1_features/test_f12_tc14_authentication.py -v
   ```
   *Expected Result*: 5 passed, 0 failed.

2. **Verify Adapter Fault Blindness & Disconnection Recovery**:
   ```powershell
   python -m pytest tests/tier2_boundaries/test_challenger2_empirical_probes.py -k "test_sap_com_sudden_disconnection_state_detection or test_sap_com_stale_session_recovery_on_get_session or test_tc14_driver_crash_detection_in_is_session_alive" -v
   ```
   *Expected Result*: 3 passed, 0 failed (all 3 defects resolved).

3. **Verify R5 Provider Adapter Pattern & Core Engine Independence**:
   ```powershell
   python -m pytest tests/unit/test_adapters.py -v
   ```
   *Expected Result*: 8 passed, 0 failed.

4. **Verify Clean Imports in GUI Leader & Member Workspaces**:
   ```powershell
   python -c "from src.gui.leader_view import LeaderWorkspaceView; from src.gui.member_view import MemberWorkspaceView; from src.automation.sap.parser import SAPBOMParser, parse_r3_cs12_file; print('Imports OK')"
   ```
   *Expected Result*: Prints `Imports OK` without raising `ImportError`.

5. **Invalidation Condition**:
   This remediation report is invalidated if any of the above 4 commands produce an unhandled exception, syntax error, or assertion failure.
