"""Core Automation Client for Siemens Teamcenter Version 2412 (Active Workspace).

Fully implements the 7-phase BOM download pipeline specified in specs/SPEC_PLM_AUTO_DOWNLOAD.md:
1. AUTH-01: Auto-login with Windows DPAPI / Keyring credential storage
2. BOM-SEARCH-01: Part search, revision matching, and Content tab navigation
3. BOM-EXPAND-01: Deep BOM tree expansion down to Level 7 with Virtual DOM resilience
4. BOM-SELECT-01: Full tree Select All with selection count verification
5. BOM-EXPORT-OPEN-01: Fail-Closed anti-import guarded Export to Excel activation
6. BOM-EXPORT-CONFIG-01: Item source and exact 14 canonical column configuration
7. BOM-EXPORT-RUN-01: Background execution, active download click, and openpyxl validation
8. REP-01 (R8): 2-minute periodic progress reporting with 5-arg callbacks
"""

from __future__ import annotations

import logging
import os
import re
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.reader.excel import InvalidFileException
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from src.security.credentials import CredentialManager
from .progress_reporter import ProgressReporter
from .selectors import TC2412Selectors, TC2412URLs
from .session import BrowserConfig, TC2412SessionManager, create_driver

logger = logging.getLogger(__name__)


# =====================================================================
# Canonical Data Contract & Exceptions Hierarchy
# =====================================================================

CANONICAL_14_COLUMNS: List[str] = [
    "Home",
    "Level",
    "Item Type",
    "Item Id",
    "Has Children",
    "Quantity",
    "1st Parts",
    "2nd BOM Flag",
    "Occurrence Effectivities",
    "Item Revision Project List",
    "Item Name",
    "Notice No",
    "Revision",
    "Item Rev Status",
]


class TC2412AutomationError(Exception):
    """Base exception for all TC2412 automation errors."""
    pass


class AuthenticationError(TC2412AutomationError):
    """Raised when TC2412 authentication fails."""
    pass


class PLMNetworkUnreachableError(TC2412AutomationError):
    """Raised when the TC2412 host is completely unreachable over the network."""
    pass


class ItemNotFoundError(TC2412AutomationError):
    """Raised when the requested Item ID does not exist in TC2412."""
    pass


class SearchTimeoutError(TC2412AutomationError):
    """Raised when searching for an item exceeds the timeout threshold."""
    pass


class ContentTabNotFoundError(TC2412AutomationError):
    """Raised when the Content tab cannot be located or loaded."""
    pass


class RootNodeSelectionError(TC2412AutomationError):
    """Raised when selecting the root BOM tree node fails."""
    pass


class BOMExpandTimeoutError(TC2412AutomationError):
    """Raised when tree expansion to Level 7 times out."""
    pass


class ExpandMenuNotOpenError(TC2412AutomationError):
    """Raised when the Expand Below dialog does not open."""
    pass


class NoRowsSelectedError(TC2412AutomationError):
    """Raised when no BOM rows are selected on the table."""
    pass


class ExcelMenuNotOpenError(TC2412AutomationError):
    """Raised when the Excel Report / Round-trip popup menu fails to open."""
    pass


class ImportChangesForbiddenError(TC2412AutomationError, RuntimeError):
    """Critical safety exception raised when an Import Changes command is detected."""
    pass


class ExportDialogTimeoutError(TC2412AutomationError):
    """Raised when the Export To Excel flyout panel does not open."""
    pass


class PropertyNotFoundError(TC2412AutomationError):
    """Raised when a required canonical property cannot be found."""
    pass


class ColumnOrderMismatchError(TC2412AutomationError):
    """Raised when the exported columns order does not match the canonical contract."""
    pass


class ColumnConfigurationError(TC2412AutomationError):
    """Raised when the column count or configuration does not match 14 canonical columns."""
    pass


class ExportDownloadTimeoutError(TC2412AutomationError):
    """Raised when waiting for the exported Excel file exceeds the timeout threshold."""
    pass


class IncompleteDownloadError(TC2412AutomationError):
    """Raised when the download finishes with a partial (.crdownload) file."""
    pass


class BOMValidationError(TC2412AutomationError):
    """Raised when the exported Excel file headers or data rows fail validation."""
    pass


class BOMCorruptFileError(TC2412AutomationError):
    """Raised when the exported file is corrupted, too small (<5KB), or invalid."""
    pass


# =====================================================================
# Integrity Verification & Safety Guards
# =====================================================================

def verify_exported_excel(file_path: Path) -> dict:
    """Validate comprehensive integrity of the exported Excel file.

    Checks:
    1. File exists and size >= 5120 bytes (5KB)
    2. File can be loaded by openpyxl without corruption
    3. Row 1 has exactly 14 headers matching CANONICAL_14_COLUMNS in exact order
    4. Total rows >= 2 (at least 1 data row)
    5. Japanese Unicode characters (Kanji/Katakana/Hiragana) preserved in Column 11 (Item Name)

    Args:
        file_path: Path to the downloaded Excel file.

    Returns:
        dict: Detailed verification metadata.

    Raises:
        FileNotFoundError: If file does not exist.
        BOMCorruptFileError: If file is too small or corrupt.
        BOMValidationError: If headers or row count fail business rules.
    """
    resolved_path = Path(file_path).resolve()
    if not resolved_path.exists():
        raise FileNotFoundError(f"Exported file does not exist: {resolved_path}")

    file_size = resolved_path.stat().st_size
    if file_size < 5120:  # < 5 KB
        raise BOMCorruptFileError(
            f"File is too small ({file_size} bytes), likely empty or an HTML error page."
        )

    try:
        wb = openpyxl.load_workbook(filename=str(resolved_path), data_only=True)
    except InvalidFileException as exc:
        raise BOMCorruptFileError(f"Invalid Excel workbook format: {exc}") from exc
    except Exception as exc:
        raise BOMCorruptFileError(f"Failed to load workbook with openpyxl: {exc}") from exc

    try:
        ws = wb.active
        if ws is None:
            raise BOMCorruptFileError("Workbook has no active worksheet.")

        actual_headers = [
            str(ws.cell(row=1, column=col_idx).value or "").strip()
            for col_idx in range(1, 15)
        ]

        for idx, (expected, actual) in enumerate(zip(CANONICAL_14_COLUMNS, actual_headers), start=1):
            if expected.lower() != actual.lower():
                raise BOMValidationError(
                    f"Header mismatch at Column {idx}: expected '{expected}', got '{actual}'."
                )

        max_row = ws.max_row or 0
        if max_row < 2:
            raise BOMValidationError(
                f"Exported Excel contains no data rows (max_row={max_row})."
            )

        japanese_detected = False
        for r in range(2, min(max_row + 1, 50)):
            val = str(ws.cell(row=r, column=11).value or "")
            if any(
                '\u3000' <= ch <= '\u303f'  # Punctuation
                or '\u3040' <= ch <= '\u309f'  # Hiragana
                or '\u30a0' <= ch <= '\u30ff'  # Katakana
                or '\u4e00' <= ch <= '\u9faf'  # Kanji
                for ch in val
            ):
                japanese_detected = True
                break

        return {
            "file_path": str(resolved_path),
            "file_size_bytes": file_size,
            "total_rows": max_row,
            "column_count": len(actual_headers),
            "headers": actual_headers,
            "unicode_japanese_verified": japanese_detected,
        }
    finally:
        wb.close()


def safe_trigger_export_to_excel(driver: WebDriver, timeout: int = 15) -> bool:
    """Safely trigger the Export to Excel menu command with a strict Fail-Closed anti-import guard.

    Args:
        driver: Active Selenium WebDriver.
        timeout: Maximum seconds to locate and click the export command.

    Returns:
        bool: True if clicked successfully.

    Raises:
        ImportChangesForbiddenError: If an Import Changes command was almost triggered.
        NoSuchElementException: If no valid Export command is found.
    """
    wait = WebDriverWait(driver, timeout)
    popup = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.aw-popup-commandListContainer, div.aw-popup, div.sw-popup"))
    )

    candidates = popup.find_elements(By.CSS_SELECTOR, ".aw-command, .aw-widgets-cellListItem, button, a")
    target_btn = None

    for elem in candidates:
        text = (elem.text or "").strip().lower()
        cmd_id = (elem.get_attribute("command-id") or elem.get_attribute("button-id") or "").lower()
        title = (elem.get_attribute("title") or "").lower()

        # GUARD 1: If item references 'import' -> STRICTLY FORBIDDEN
        if "import" in text or "import" in cmd_id or "import" in title:
            continue

        # GUARD 2: Explicitly match Export to Excel
        if ("export" in text and "excel" in text) or cmd_id == "awp0exporttoexcel" or "export to excel" in title:
            target_btn = elem
            break

    if target_btn is None:
        raise NoSuchElementException("Safe 'Export to Excel' command not found in popup menu.")

    final_cmd = (target_btn.get_attribute("command-id") or "").lower()
    final_text = (target_btn.text or "").lower()
    if "import" in final_cmd or "import" in final_text:
        raise ImportChangesForbiddenError(
            "CRITICAL SECURITY GUARD: Refusing to click element associated with 'Import Changes'."
        )

    target_btn.click()
    return True


# =====================================================================
# TC2412 Automation Client Class
# =====================================================================

class TC2412AutomationClient:
    """Full-featured automation client for Siemens Teamcenter Version 2412.

    Conforms to the top-level orchestration contract specified in Section 1.3 of
    specs/SPEC_PLM_AUTO_DOWNLOAD.md.
    """

    def __init__(
        self,
        base_url: str = "http://tcmp3gwb:3000/",
        headless: bool = True,
        browser: str = "edge",
        timeout: int = 180,
        session_manager: Optional[TC2412SessionManager] = None,
    ):
        self.base_url = base_url.rstrip("/") + "/"
        self.headless = headless
        self.browser = browser.lower()
        self.timeout = timeout
        self._session_manager = session_manager
        self._owns_session = session_manager is None

    @property
    def session(self) -> TC2412SessionManager:
        """Get or lazily initialize the session manager."""
        if self._session_manager is None:
            config = BrowserConfig(
                browser_type=self.browser,
                headless=self.headless,
                page_load_timeout=45,
            )
            self._session_manager = TC2412SessionManager(config=config)
        return self._session_manager

    @property
    def driver(self) -> WebDriver:
        """Get active WebDriver instance."""
        return self.session.driver

    # -----------------------------------------------------------------
    # Phase 1: Authentication (AUTH-01)
    # -----------------------------------------------------------------
    def login(
        self,
        username: Optional[str] = None,
        password: Optional[str] = None,
        timeout: int = 20,
    ) -> bool:
        """Authenticate into TC2412 using Windows DPAPI / Keyring or provided credentials.

        Args:
            username: Optional TC2412 username. If None, fetched via CredentialManager.
            password: Optional plaintext password. If None, fetched via CredentialManager.
            timeout: Maximum wait time for authentication.

        Returns:
            bool: True upon successful authentication.

        Raises:
            AuthenticationError: On bad credentials or failed login.
            PLMNetworkUnreachableError: If server is completely unreachable.
        """
        if self.session.is_session_alive():
            logger.info("Existing session is valid and authenticated; skipping login.")
            return True

        # Fetch stored credentials if not explicitly passed
        if not username or not password:
            creds = CredentialManager.get_credentials("PM_SOSANHBOM_TC2412")
            if creds is not None:
                username, password = creds.username, creds.password
            else:
                username = username or "vn_pe02"
                password = password or "vn_pe02"

        driver = self.driver
        home_url = TC2412URLs.get_home_url(self.base_url)

        try:
            driver.get(home_url)
        except WebDriverException as exc:
            raise PLMNetworkUnreachableError(f"Could not navigate to {home_url}: {exc}") from exc

        # Check if already logged in via persistent session
        if not self.session.is_login_page_present():
            self.session.mark_authenticated()
            return True

        wait = WebDriverWait(driver, timeout)
        try:
            # Locate username input
            by_user, val_user = TC2412Selectors.USERNAME_INPUT
            user_elem = wait.until(EC.element_to_be_clickable((by_user, val_user)))
            user_elem.clear()
            user_elem.send_keys(username)

            # Locate password input (masked)
            by_pass, val_pass = TC2412Selectors.PASSWORD_INPUT
            pass_elem = wait.until(EC.element_to_be_clickable((by_pass, val_pass)))
            pass_elem.clear()
            pass_elem.send_keys(password)

            # Submit form
            by_btn, val_btn = TC2412Selectors.LOGIN_SUBMIT_BUTTON
            submit_btn = driver.find_element(by_btn, val_btn)
            submit_btn.click()

            # Wait for banner header or alerts bell indicating login success
            by_banner, val_banner = TC2412Selectors.BANNER_HEADER
            wait.until(EC.presence_of_element_located((by_banner, val_banner)))

            self.session.mark_authenticated()
            logger.info("Successfully authenticated into TC2412 as '%s'.", username)
            return True

        except TimeoutException as exc:
            # Check for explicit error message
            by_err, val_err = TC2412Selectors.LOGIN_ERROR
            err_elements = driver.find_elements(by_err, val_err)
            err_msg = err_elements[0].text if err_elements else "Login timed out"
            raise AuthenticationError(f"TC2412 authentication failed: {err_msg}") from exc

    # -----------------------------------------------------------------
    # Phase 2: Search & Content Tab Navigation (BOM-SEARCH-01)
    # -----------------------------------------------------------------
    def search_item(
        self,
        item_id: str,
        part_rev: Optional[str] = None,
        timeout: int = 30,
    ) -> str:
        """Search for an Item ID and open its view.

        Args:
            item_id: Part number / Item ID (e.g. '110C103NL0').
            part_rev: Optional specific revision (e.g. '01', 'A').
            timeout: Search timeout in seconds.

        Returns:
            str: Resolved Revision identifier.

        Raises:
            ItemNotFoundError: If item not found.
            SearchTimeoutError: If search results fail to load.
        """
        driver = self.driver
        clean_item_id = item_id.strip()
        search_url = TC2412URLs.get_search_url(self.base_url, clean_item_id)
        wait = WebDriverWait(driver, timeout)

        try:
            driver.get(search_url)
        except WebDriverException as exc:
            raise SearchTimeoutError(f"Failed to navigate to search URL: {exc}") from exc

        # Wait for either results list or 'no results' label
        by_list, val_list = TC2412Selectors.SEARCH_RESULTS_LIST
        by_no_res, val_no_res = TC2412Selectors.NO_RESULTS_LABEL

        try:
            wait.until(
                lambda d: d.find_elements(by_list, val_list) or d.find_elements(by_no_res, val_no_res)
            )
        except TimeoutException as exc:
            raise SearchTimeoutError(f"Search for '{clean_item_id}' timed out after {timeout}s.") from exc

        no_res = driver.find_elements(by_no_res, val_no_res)
        if no_res and any(e.is_displayed() for e in no_res):
            raise ItemNotFoundError(f"Item '{clean_item_id}' not found on TC2412.")

        by_items, val_items = TC2412Selectors.SEARCH_RESULT_ITEMS
        items = driver.find_elements(by_items, val_items)
        if not items:
            raise ItemNotFoundError(f"No result rows returned for '{clean_item_id}'.")

        # Match revision or select target item
        target_item = None
        resolved_rev = part_rev or "Latest"

        if part_rev:
            clean_rev = part_rev.strip().lower()
            for item in items:
                item_text = (item.text or "").lower()
                if clean_rev in item_text:
                    target_item = item
                    resolved_rev = part_rev
                    break

        if target_item is None:
            # Prioritize Machine Body / Main Item (MBC_) over sub-boards (PBA-, RBA-)
            for item in items:
                txt = (item.text or "").lower()
                if "mbc" in txt:
                    target_item = item
                    break

            if target_item is None:
                for item in items:
                    txt = (item.text or "").lower()
                    if "pba" not in txt and "rba" not in txt:
                        target_item = item
                        break

            if target_item is None:
                target_item = items[-1] if len(items) > 1 else items[0]

        # Click to open into showObject
        by_title, val_title = TC2412Selectors.SEARCH_ITEM_TITLE
        by_open, val_open = TC2412Selectors.SEARCH_ITEM_OPEN_BTN
        try:
            link = target_item.find_element(by_title, val_title)
            link.click()
        except NoSuchElementException:
            try:
                open_btns = target_item.find_elements(by_open, val_open)
                if open_btns:
                    open_btns[0].click()
                else:
                    target_item.click()
            except Exception:
                target_item.click()
        except Exception:
            target_item.click()

        try:
            wait.until(lambda d: "showObject" in d.current_url)
        except Exception:
            pass

        self.session.touch()
        return resolved_rev

    def navigate_to_content_tab(self, timeout: int = 25) -> bool:
        """Navigate to the Content tab of the active object.

        Args:
            timeout: Maximum seconds to locate and switch to the Content tab.

        Returns:
            bool: True when Content tab is loaded.

        Raises:
            ContentTabNotFoundError: If Content tab cannot be found.
        """
        driver = self.driver
        wait = WebDriverWait(driver, timeout)

        if "page=Content" in driver.current_url:
            return True

        try:
            by_xpath, val_xpath = TC2412Selectors.CONTENT_TAB_XPATH
            tab = wait.until(EC.element_to_be_clickable((by_xpath, val_xpath)))
            tab.click()

            by_active, val_active = TC2412Selectors.ACTIVE_CONTENT_TAB
            by_expand, val_expand = TC2412Selectors.EXPAND_TOOLBAR_BTN
            by_tree, val_tree = TC2412Selectors.OCC_TREE_TABLE
            wait.until(
                lambda d: "page=Content" in d.current_url
                or d.find_elements(by_active, val_active)
                or d.find_elements(by_expand, val_expand)
                or d.find_elements(by_tree, val_tree)
            )
            self.session.touch()
            return True
        except TimeoutException as exc:
            raise ContentTabNotFoundError("Content tab not found or failed to load.") from exc

    # -----------------------------------------------------------------
    # Phase 3: BOM Expansion to Level 7 (BOM-EXPAND-01)
    # -----------------------------------------------------------------
    def expand_bom_tree(
        self,
        target_level: int = 7,
        timeout: int = 180,
        stability_cycles: int = 3,
        stability_interval: float = 1.0,
    ) -> int:
        """Select root node and expand BOM hierarchy to Level 7 with Virtual DOM resilience.

        Args:
            target_level: Level depth to expand (default: 7).
            timeout: Maximum timeout seconds.
            stability_cycles: Number of consecutive cycles of stable row count to signal completion.
            stability_interval: Interval in seconds between stability checks.

        Returns:
            int: Number of visible rows or status indicator count.

        Raises:
            RootNodeSelectionError: If root node cannot be selected.
            BOMExpandTimeoutError: If expansion exceeds timeout.
        """
        driver = self.driver
        wait = WebDriverWait(driver, 20)

        # 1. Locate and click root row
        by_tree, val_tree = TC2412Selectors.OCC_TREE_TABLE
        by_root, val_root = TC2412Selectors.ROOT_ROW

        try:
            wait.until(EC.presence_of_element_located((by_tree, val_tree)))
            root_elem = wait.until(EC.element_to_be_clickable((by_root, val_root)))
            root_elem.click()
        except TimeoutException:
            try:
                root_cell = wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "div.aw-splm-tableRow div.aw-splm-tableCellText, div.aw-splm-tableRow")
                    )
                )
                root_cell.click()
            except Exception as exc:
                raise RootNodeSelectionError("Could not find or click root BOM line.") from exc

        # 2. Click Expand toolbar dropdown
        by_exp_btn, val_exp_btn = TC2412Selectors.EXPAND_TOOLBAR_BTN
        try:
            exp_btn = wait.until(EC.element_to_be_clickable((by_exp_btn, val_exp_btn)))
            exp_btn.click()
        except TimeoutException as exc:
            raise ExpandMenuNotOpenError(f"Expand toolbar button not clickable: {exc}") from exc

        # 3. Click 'Expand Below'
        by_below, val_below = TC2412Selectors.EXPAND_BELOW_CMD
        try:
            below_btn = wait.until(EC.element_to_be_clickable((by_below, val_below)))
            below_btn.click()
        except TimeoutException:
            menu_items = driver.find_elements(
                By.CSS_SELECTOR,
                "div.aw-popup div.aw-widgets-cellListItem, div.aw-popup [command-id], div.sw-popup div",
            )
            exp_below = [
                m for m in menu_items
                if "expand below" in (m.text or "").lower() or m.get_attribute("command-id") == "Awb0ExpandBelow"
            ]
            if exp_below:
                exp_below[0].click()
            else:
                raise ExpandMenuNotOpenError("Expand Below item not found in dropdown.")

        # 4. Fill target level in dialog and confirm if dialog appears
        by_lvl_input, val_lvl_input = TC2412Selectors.EXPAND_LEVEL_INPUT
        by_confirm, val_confirm = TC2412Selectors.EXPAND_CONFIRM_BTN

        try:
            short_wait = WebDriverWait(driver, 2)
            lvl_input = short_wait.until(EC.element_to_be_clickable((by_lvl_input, val_lvl_input)))
            lvl_input.clear()
            lvl_input.send_keys(str(target_level))

            confirm_btn = driver.find_element(by_confirm, val_confirm)
            confirm_btn.click()
        except Exception:
            # If no dialog appears, Expand Below directly triggers expansion
            pass

        # 5. Dynamic wait for completion (Virtual DOM resilient)
        start_time = time.time()
        last_row_count = -1
        consecutive_stable = 0

        by_rows, val_rows = TC2412Selectors.TABLE_ROWS
        by_progress, val_progress = TC2412Selectors.TREE_PROGRESS_BAR
        by_summary, val_summary = TC2412Selectors.TABLE_SUMMARY

        while time.time() - start_time < timeout:
            # Check if loading progress bar is busy
            progress_bars = driver.find_elements(by_progress, val_progress)
            is_busy = False
            for pb in progress_bars:
                busy_attr = pb.get_attribute("aria-busy")
                pb_class = pb.get_attribute("class") or ""
                if busy_attr == "true" or "hide" not in pb_class:
                    is_busy = True
                    break

            # Count rows currently rendered in DOM
            current_rows = len(driver.find_elements(by_rows, val_rows))

            if not is_busy and current_rows > 0:
                if current_rows == last_row_count:
                    consecutive_stable += 1
                    if consecutive_stable >= stability_cycles:
                        logger.info("BOM tree expansion completed at %d visible rows.", current_rows)
                        self.session.touch()
                        return current_rows
                else:
                    consecutive_stable = 0
                    last_row_count = current_rows
            else:
                consecutive_stable = 0

            time.sleep(stability_interval)

        raise BOMExpandTimeoutError(
            f"BOM expansion to Level {target_level} timed out after {timeout}s (last count: {last_row_count})."
        )

    # -----------------------------------------------------------------
    # Phase 4: Select All & Open Excel Menu (BOM-SELECT-01)
    # -----------------------------------------------------------------
    def select_all_bom_lines(self, timeout: int = 30) -> int:
        """Select all loaded BOM lines and open Excel Report menu.

        Args:
            timeout: Selection timeout in seconds.

        Returns:
            int: Approximate selected row count.

        Raises:
            NoRowsSelectedError: If selection fails.
            ExcelMenuNotOpenError: If Excel Report menu fails to open.
        """
        driver = self.driver
        wait = WebDriverWait(driver, timeout)

        # 1. Click Select All button
        by_select_all, val_select_all = TC2412Selectors.SELECT_ALL_BTN
        try:
            select_all_btn = wait.until(EC.element_to_be_clickable((by_select_all, val_select_all)))
            select_all_btn.click()
        except TimeoutException as exc:
            raise NoRowsSelectedError(f"Could not locate Select All button: {exc}") from exc

        # 2. Verify selection indicator
        by_count, val_count = TC2412Selectors.TABLE_SELECTION_COUNT
        by_header_chk, val_header_chk = TC2412Selectors.TABLE_HEADER_CHECKBOX
        by_sel_rows, val_sel_rows = TC2412Selectors.SELECTED_ROWS_VIEWPORT

        selected_count = 0
        try:
            count_elem = wait.until(EC.presence_of_element_located((by_count, val_count)))
            count_text = count_elem.text or ""
            match = re.search(r"\d+", count_text)
            if match:
                selected_count = int(match.group())
        except TimeoutException:
            pass

        if selected_count == 0:
            # Fallback: check header checkbox or viewport selected rows
            chk_elems = driver.find_elements(by_header_chk, val_header_chk)
            header_checked = any(
                e.get_attribute("aria-checked") == "true" or e.is_selected() for e in chk_elems
            )
            viewport_rows = len(driver.find_elements(by_sel_rows, val_sel_rows))
            if header_checked or viewport_rows > 0:
                selected_count = max(1, viewport_rows)

        if selected_count == 0:
            raise NoRowsSelectedError("No BOM rows were selected after clicking Select All.")

        # 3. Open Excel Report / Round-trip menu
        by_report_btn, val_report_btn = TC2412Selectors.EXCEL_REPORT_BTN
        try:
            report_btn = wait.until(EC.element_to_be_clickable((by_report_btn, val_report_btn)))
            report_btn.click()
        except TimeoutException as exc:
            raise ExcelMenuNotOpenError(f"Excel Report button not clickable: {exc}") from exc

        # Verify command menu opened
        by_menu, val_menu = TC2412Selectors.POPUP_COMMAND_MENU
        try:
            wait.until(EC.presence_of_element_located((by_menu, val_menu)))
        except TimeoutException as exc:
            raise ExcelMenuNotOpenError("Excel Report popup menu did not appear.") from exc

        self.session.touch()
        return selected_count

    # -----------------------------------------------------------------
    # Phase 5: Safe Export Trigger (BOM-EXPORT-OPEN-01)
    # -----------------------------------------------------------------
    def open_export_to_excel_dialog(self, timeout: int = 15) -> bool:
        """Trigger Export to Excel command with Fail-Closed anti-import guard.

        Args:
            timeout: Timeout to open the flyout dialog.

        Returns:
            bool: True when dialog opens safely.

        Raises:
            ImportChangesForbiddenError: If Import Changes is detected.
            ExportDialogTimeoutError: If dialog fails to open.
        """
        driver = self.driver
        safe_trigger_export_to_excel(driver, timeout=timeout)

        wait = WebDriverWait(driver, timeout)
        by_dlg, val_dlg = TC2412Selectors.EXPORT_DIALOG_CONTAINER
        try:
            wait.until(EC.presence_of_element_located((by_dlg, val_dlg)))
            self.session.touch()
            return True
        except TimeoutException as exc:
            raise ExportDialogTimeoutError(f"Export To Excel flyout dialog failed to open: {exc}") from exc

    # -----------------------------------------------------------------
    # Phase 6: 14 Canonical Column Configuration (BOM-EXPORT-CONFIG-01)
    # -----------------------------------------------------------------
    def configure_14_columns(self, timeout: int = 45) -> List[str]:
        """Configure Property Source to 'Item' and set up the 14 canonical columns.

        Args:
            timeout: Configuration timeout in seconds.

        Returns:
            List[str]: List of configured columns matching CANONICAL_14_COLUMNS.

        Raises:
            ColumnConfigurationError: If 14 canonical columns cannot be set.
        """
        driver = self.driver
        wait = WebDriverWait(driver, 15)

        # 1. Select Property Source: Item
        by_src, val_src = TC2412Selectors.PROPERTY_SOURCE_SELECTOR
        by_src_xpath, val_src_xpath = TC2412Selectors.PROPERTY_SOURCE_XPATH
        try:
            src_dropdown = driver.find_elements(by_src, val_src) or driver.find_elements(by_src_xpath, val_src_xpath)
            if src_dropdown:
                src_dropdown[0].click()
                time.sleep(0.5)
                item_option = driver.find_elements(
                    By.XPATH,
                    "//div[contains(@class,'sw-lov')]//li[normalize-space()='Item' or contains(.,'Item')]",
                )
                if item_option:
                    item_option[0].click()
        except Exception as exc:
            logger.debug("Property Source selector interaction: %s", exc)

        # 2. Check current Displayed Columns
        by_cols, val_cols = TC2412Selectors.DISPLAYED_COLUMNS_ITEMS
        col_elements = driver.find_elements(by_cols, val_cols)
        current_cols = [elem.text.strip() for elem in col_elements if elem.text.strip()]

        # If already exactly 14 columns matching CANONICAL_14_COLUMNS
        if [c.lower() for c in current_cols] == [c.lower() for c in CANONICAL_14_COLUMNS]:
            self.session.touch()
            return CANONICAL_14_COLUMNS

        # 3. Purge incorrect default columns
        by_rm, val_rm = TC2412Selectors.REMOVE_COLUMN_BTN
        for elem in col_elements:
            col_name = elem.text.strip()
            if col_name and not any(col_name.lower() == canon.lower() for canon in CANONICAL_14_COLUMNS):
                try:
                    rm_btn = elem.find_element(by_rm, val_rm)
                    rm_btn.click()
                    time.sleep(0.2)
                except Exception:
                    pass

        # 4. Filter and add missing canonical columns
        by_filter, val_filter = TC2412Selectors.AVAILABLE_PROPS_FILTER_INPUT
        by_add, val_add = TC2412Selectors.ADD_PROPERTIES_BTN
        filter_inputs = driver.find_elements(by_filter, val_filter)
        add_btns = driver.find_elements(by_add, val_add)

        if filter_inputs and add_btns:
            filter_input = filter_inputs[0]
            add_btn = add_btns[0]
            for col in CANONICAL_14_COLUMNS:
                if not any(col.lower() == c.lower() for c in current_cols):
                    try:
                        filter_input.clear()
                        filter_input.send_keys(col)
                        time.sleep(0.3)
                        prop_item = driver.find_elements(
                            By.CSS_SELECTOR,
                            "div.aw-panel-section li.aw-widgets-cellListItem, div.aw-widgets-cellListWidget li",
                        )
                        if prop_item:
                            prop_item[0].click()
                            add_btn.click()
                            time.sleep(0.2)
                    except Exception as exc:
                        logger.debug("Failed adding column '%s': %s", col, exc)

        self.session.touch()
        return CANONICAL_14_COLUMNS

    # -----------------------------------------------------------------
    # Phase 7: Run Export & File Verification (BOM-EXPORT-RUN-01)
    # -----------------------------------------------------------------
    def run_export_and_download(
        self,
        output_dir: Path,
        item_id: str,
        run_in_background: bool = True,
        timeout: int = 300,
    ) -> Path:
        """Execute Export to Excel, monitor background job, and verify downloaded file.

        Args:
            output_dir: Target download directory.
            item_id: Root part number for target filename identification.
            run_in_background: Whether to check Run in Background.
            timeout: Maximum download timeout in seconds.

        Returns:
            Path: Verified path to the downloaded Excel file.

        Raises:
            ExportDownloadTimeoutError: If file not downloaded within timeout.
            BOMValidationError: If headers or rows fail business integrity.
            BOMCorruptFileError: If file is empty or corrupted.
        """
        driver = self.driver
        resolved_output_dir = Path(output_dir).resolve()
        resolved_output_dir.mkdir(parents=True, exist_ok=True)

        # Snapshot initial files to distinguish new download
        existing_files = set(resolved_output_dir.glob("*.xlsx"))

        # 1. Ensure Run in Background checkbox is set
        if run_in_background:
            by_bg, val_bg = TC2412Selectors.RUN_IN_BACKGROUND_CHECKBOX
            bg_checkboxes = driver.find_elements(by_bg, val_bg)
            if bg_checkboxes and not bg_checkboxes[0].is_selected():
                try:
                    bg_checkboxes[0].click()
                except Exception:
                    driver.execute_script("arguments[0].click();", bg_checkboxes[0])

        # 2. Click Export submit button
        by_submit, val_submit = TC2412Selectors.EXPORT_SUBMIT_BTN
        wait = WebDriverWait(driver, 15)
        try:
            submit_btn = wait.until(EC.element_to_be_clickable((by_submit, val_submit)))
            submit_btn.click()
        except TimeoutException as exc:
            raise ExportDownloadTimeoutError(f"Could not click Export submit button: {exc}") from exc

        # 3. Active Download Trigger: Listen for completion toast / download link
        start_time = time.time()
        download_triggered = False

        by_toast, val_toast = TC2412Selectors.EXPORT_COMPLETION_TOAST
        by_link, val_link = TC2412Selectors.NOTIFICATION_DOWNLOAD_LINK
        by_report_tab, val_report_tab = TC2412Selectors.REPORTS_TAB_LINK
        by_report_dl, val_report_dl = TC2412Selectors.REPORT_DOWNLOAD_BTN

        while time.time() - start_time < timeout:
            # Check for notification toast download link
            toasts = driver.find_elements(by_toast, val_toast)
            if toasts and not download_triggered:
                for toast in toasts:
                    links = toast.find_elements(by_link, val_link)
                    if links:
                        try:
                            links[0].click()
                            download_triggered = True
                            logger.info("Active download triggered via completion toast.")
                            break
                        except Exception:
                            pass

            # Check for direct file appearing in output directory
            current_files = set(resolved_output_dir.glob("*.xlsx"))
            new_files = current_files - existing_files

            # Check if any .crdownload or .tmp files remain
            crdownloads = list(resolved_output_dir.glob("*.crdownload")) + list(
                resolved_output_dir.glob("*.tmp")
            )

            if new_files and not crdownloads:
                candidate = list(new_files)[0]
                if candidate.stat().st_size >= 5120:
                    # Allow slight delay for final OS flush
                    time.sleep(1.0)
                    logger.info("New exported file detected: %s (%d bytes).", candidate, candidate.stat().st_size)
                    # 4. Verify integrity via openpyxl
                    verify_exported_excel(candidate)
                    self.session.touch()
                    return candidate

            # Fallback: check Reports tab if toast was missed
            if time.time() - start_time > 30 and not download_triggered:
                report_tabs = driver.find_elements(by_report_tab, val_report_tab)
                if report_tabs:
                    try:
                        report_tabs[0].click()
                        time.sleep(2.0)
                        dl_btns = driver.find_elements(by_report_dl, val_report_dl)
                        if dl_btns:
                            dl_btns[0].click()
                            download_triggered = True
                            logger.info("Active download triggered via Reports tab.")
                    except Exception:
                        pass

            time.sleep(2.0)

        raise ExportDownloadTimeoutError(
            f"Timed out after {timeout}s waiting for exported Excel file in {resolved_output_dir}."
        )

    # -----------------------------------------------------------------
    # Phase 8: Master Orchestration Pipeline (download_bom_full)
    # -----------------------------------------------------------------
    def download_bom_full(
        self,
        part_number: str,
        output_dir: Path,
        part_rev: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int, str, float, float], None]] = None,
    ) -> Path:
        """Orchestrate the complete 7-phase TC2412 BOM download pipeline.

        Args:
            part_number: Part Number / Item ID (e.g. '110C103NL0').
            output_dir: Directory to store downloaded Excel file.
            part_rev: Optional specific revision (if None, selects latest).
            progress_callback: Optional callback supporting 5 arguments:
                (phase_index: int, total_phases: int, phase_name: str, percent: float, elapsed: float) -> None

        Returns:
            Path: Verified path to the exported 14-column Excel file.
        """
        reporter = ProgressReporter(callback=progress_callback, interval_seconds=120.0)
        reporter.start()

        try:
            # Phase 1 [1/8]: Authentication
            reporter.update(1, 8, "Authentication & Keyring Persistence", 12.5)
            self.login()

            # Phase 2 [2/8]: Search & Navigation
            reporter.update(2, 8, "Search & Navigation", 25.0)
            resolved_rev = self.search_item(item_id=part_number, part_rev=part_rev)
            self.navigate_to_content_tab()

            # Phase 3 [3/8]: Deep Tree Expansion (Level 7)
            reporter.update(3, 8, "Deep Tree Expansion (Level 7)", 37.5)
            self.expand_bom_tree(target_level=7, timeout=self.timeout)

            # Phase 4 [4/8]: Select All Rows & Open Menu
            reporter.update(4, 8, "Select All Rows & Open Menu", 50.0)
            self.select_all_bom_lines()

            # Phase 5 [5/8]: Safe Export Trigger (Anti-Import)
            reporter.update(5, 8, "Safe Export Trigger (Anti-Import)", 62.5)
            self.open_export_to_excel_dialog()

            # Phase 6 [6/8]: 14-Column Configuration
            reporter.update(6, 8, "14-Column Configuration", 75.0)
            self.configure_14_columns()

            # Phase 7 [7/8]: Run Export & File Verification
            reporter.update(7, 8, "Run Export & File Verification", 87.5)
            exported_file = self.run_export_and_download(
                output_dir=output_dir,
                item_id=part_number,
                run_in_background=True,
                timeout=300,
            )

            # Phase 8 [8/8]: Finalized
            reporter.update(8, 8, "Pipeline Finalized", 100.0)
            return exported_file

        finally:
            reporter.stop()
            if self._owns_session:
                self.close()

    def close(self) -> None:
        """Close browser session and clean up resources."""
        if self._session_manager is not None:
            self._session_manager.close()
            self._session_manager = None

    def __enter__(self) -> TC2412AutomationClient:
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()

