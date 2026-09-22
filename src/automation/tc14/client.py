"""High-level automation client for Siemens Teamcenter Active Workspace (TC14).

Provides full lifecycle automation:
- Automated authentication with 'vn_pe03 / vn_pe03'
- Direct hash navigation avoiding UI lag and autocomplete delays
- Item opening and Content/Occurrence tree navigation
- Native Excel export via 'Awp0ExportToExcel'
- Fallback interactive tree scraper and OpenXML generator
- Fulfills the Project Interface Contract: TC14Client.download_bom_full(part_number, output_dir)
"""

from __future__ import annotations

import logging
import re
import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, List, Optional, Union

import openpyxl
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from .selectors import TC14Selectors, TC14URLs
from .session import BrowserConfig, TC14SessionManager

logger = logging.getLogger(__name__)


class TC14Error(Exception):
    """Base exception for Teamcenter TC14 automation errors."""


class TC14AuthenticationError(TC14Error):
    """Raised when authentication fails or credentials are rejected."""


class TC14SearchError(TC14Error):
    """Raised when search returns no matching items or fails."""


class TC14NavigationError(TC14Error):
    """Raised when navigating to an object, view, or tab fails."""


class TC14ExportError(TC14Error):
    """Raised when BOM Excel export fails or times out."""


class TC14TimeoutError(TC14Error):
    """Raised when a wait condition exceeds specified timeout."""


@dataclass
class SearchResult:
    """Represents a matched item revision returned from TC14 search."""

    item_id: str
    title: str
    revision: str = ""
    uid: Optional[str] = None
    type_name: Optional[str] = None
    web_element: Optional[WebElement] = field(default=None, repr=False)


@dataclass
class BOMItem:
    """Represents an occurrence row in the TC14 BOM tree."""

    level: int = 1
    item_id: str = ""
    item_name: str = ""
    has_children: bool = False
    revision: str = ""
    effectivity: str = ""
    quantity: float = 1.0
    unit_name: str = ""
    raw_text: str = ""


class TC14AutomationClient:
    """High-level automation client for Teamcenter Active Workspace."""

    def __init__(
        self,
        base_url: str = TC14URLs.DEFAULT_BASE_URL,
        download_dir: Optional[Path] = None,
        browser: str = "edge",
        headless: bool = True,
        timeout: int = 30,
        session_manager: Optional[TC14SessionManager] = None,
    ) -> None:
        self.base_url: str = base_url.rstrip("/") + "/"
        self.download_dir: Path = Path(
            download_dir or Path.cwd() / "downloads"
        ).resolve()
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.timeout: int = timeout

        if session_manager is not None:
            self.session = session_manager
        else:
            config = BrowserConfig(
                browser_type=browser,
                headless=headless,
                download_dir=self.download_dir,
                page_load_timeout=timeout + 15,
                implicit_wait=5,
            )
            self.session = TC14SessionManager(config=config)

    @property
    def driver(self) -> WebDriver:
        """Access underlying WebDriver."""
        return self.session.driver

    # -------------------------------------------------------------------------
    # Authentication & Session Lifecycle
    # -------------------------------------------------------------------------

    def login(
        self,
        username: str = "vn_pe02",
        password: str = "vn_pe02",
        force: bool = False,
        timeout: Optional[float] = None,
    ) -> bool:
        """Authenticate with Teamcenter Active Workspace.

        Args:
            username: Login user account (default 'vn_pe02').
            password: Login password (default 'vn_pe02').
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

            # Wait for successful login transition:
            # 1. URL change to #/showHome OR
            # 2. Appearance of main header / banner / searchbox OR
            # 3. Appearance of tiles canvas
            def _login_succeeded(d: WebDriver) -> bool:
                url = d.current_url.lower()
                if "showhome" in url or "showobject" in url or "search" in url:
                    return True
                b_type, b_sel = TC14Selectors.BANNER_HEADER
                if d.find_elements(b_type, b_sel):
                    return True
                if d.find_elements(By.CSS_SELECTOR, "input.aw-uiwidgets-searchBox, .aw-tile-tileCanvasPanel"):
                    return True
                return False

            wait.until(_login_succeeded)

            # Double check for login error messages
            err_type, err_sel = TC14Selectors.LOGIN_ERROR
            error_elems = driver.find_elements(err_type, err_sel)
            visible_errors = [e.text.strip() for e in error_elems if e.is_displayed() and e.text.strip()]
            if visible_errors:
                raise TC14AuthenticationError(f"Login failed with error: {'; '.join(visible_errors)}")

            self.session.mark_authenticated(username)
            logger.info("Successfully authenticated as %s. Current URL: %s", username, driver.current_url)
            return True

        except TimeoutException as exc:
            # Check if there is an error message displayed
            err_type, err_sel = TC14Selectors.LOGIN_ERROR
            error_elems = driver.find_elements(err_type, err_sel)
            visible_errors = [e.text.strip() for e in error_elems if e.is_displayed() and e.text.strip()]
            err_msg = "; ".join(visible_errors) if visible_errors else "Timed out waiting for login completion."
            logger.error("Authentication timeout: %s", err_msg)
            raise TC14AuthenticationError(f"Failed to log in: {err_msg}") from exc
        except Exception as exc:
            logger.error("Unexpected error during login: %s", exc)
            raise TC14AuthenticationError(f"Unexpected login error: {exc}") from exc

    def ensure_authenticated(
        self, username: str = "vn_pe02", password: str = "vn_pe02"
    ) -> bool:
        """Verify session is alive and valid; re-login automatically if expired."""
        if self.session.is_session_alive():
            self.session.touch()
            return True
        logger.info("Session expired or unauthenticated; re-authenticating...")
        return self.login(username=username, password=password)

    # -------------------------------------------------------------------------
    # Search & Navigation
    # -------------------------------------------------------------------------

    def search_part(self, part_number: str) -> List[SearchResult]:
        """Search for a part or machine assembly code in Teamcenter.

        Uses direct hash URL navigation as the primary deterministic route,
        with interactive search box typing as fallback.

        Args:
            part_number: Part code or assembly number (e.g. '110C103NL0').

        Returns:
            List of SearchResult objects matching the query.

        Raises:
            TC14SearchError: If search fails or yields 0 results.
        """
        self.ensure_authenticated()
        driver = self.driver
        wait = WebDriverWait(driver, self.timeout)

        search_url = TC14URLs.get_search_url(self.base_url, part_number)
        logger.info("Executing direct hash search: %s", search_url)
        driver.get(search_url)

        # Wait for search results container or items to appear
        by_item, sel_item = TC14Selectors.SEARCH_RESULT_ITEMS
        by_no_res, sel_no_res = TC14Selectors.NO_RESULTS_LABEL

        def _search_completed(d: WebDriver) -> bool:
            items = d.find_elements(by_item, sel_item)
            if items and any(i.is_displayed() for i in items):
                return True
            no_res = d.find_elements(by_no_res, sel_no_res)
            if no_res and any(n.is_displayed() for n in no_res):
                return True
            return False

        try:
            wait.until(_search_completed)
        except TimeoutException:
            logger.warning("Direct hash search timed out; attempting UI search fallback...")
            self._search_via_ui(part_number)
            wait.until(_search_completed)

        # Parse results
        item_elements = driver.find_elements(by_item, sel_item)
        results: List[SearchResult] = []

        for elem in item_elements:
            try:
                title_elem = elem.find_element(*TC14Selectors.SEARCH_ITEM_TITLE)
                raw_title = title_elem.text.strip()
                if not raw_title:
                    continue

                # Parse revision if present in format "Revision:A" or "/A;"
                revision_match = re.search(r"Revision:([A-Za-z0-9_-]+)", raw_title)
                revision = revision_match.group(1) if revision_match else ""

                # Extract uid from href or attribute if present
                href = title_elem.get_attribute("href") or ""
                uid_match = re.search(r"uid=([A-Za-z0-9_]+)", href)
                uid = uid_match.group(1) if uid_match else None

                results.append(
                    SearchResult(
                        item_id=part_number,
                        title=raw_title,
                        revision=revision,
                        uid=uid,
                        web_element=elem,
                    )
                )
            except (NoSuchElementException, StaleElementReferenceException):
                continue

        logger.info("Search for '%s' returned %d results.", part_number, len(results))
        if not results:
            # Verify if explicit no results label exists
            no_res_elems = driver.find_elements(by_no_res, sel_no_res)
            msg = f"No search results found in TC14 for part '{part_number}'"
            if no_res_elems:
                msg += f": {no_res_elems[0].text.strip()}"
            raise TC14SearchError(msg)

        self.session.touch()
        return results

    def _search_via_ui(self, search_term: str) -> None:
        """Fallback method typing into the global search box."""
        driver = self.driver
        wait = WebDriverWait(driver, 15)

        by_box, sel_box = TC14Selectors.SEARCH_INPUT
        search_box = wait.until(EC.element_to_be_clickable((by_box, sel_box)))
        search_box.clear()
        search_box.send_keys(search_term)
        search_box.send_keys(Keys.ENTER)

    def open_item(
        self,
        target: Union[SearchResult, str, int] = 0,
        search_results: Optional[List[SearchResult]] = None,
    ) -> bool:
        """Open an item revision in the Active Workspace object view.

        Args:
            target: Either an index (int), a SearchResult instance, or part code / UID (str).
            search_results: Optional pre-fetched search results list when using integer index.

        Returns:
            True if object view loaded successfully.

        Raises:
            TC14NavigationError: If item cannot be opened.
        """
        driver = self.driver
        wait = WebDriverWait(driver, self.timeout)

        # Resolve target to an element or UID
        target_elem: Optional[WebElement] = None
        target_uid: Optional[str] = None

        if isinstance(target, SearchResult):
            target_elem = target.web_element
            target_uid = target.uid
        elif isinstance(target, int):
            results = search_results or []
            if not results:
                # Find current result items on page
                by_item, sel_item = TC14Selectors.SEARCH_RESULT_ITEMS
                found = driver.find_elements(by_item, sel_item)
                if target < len(found):
                    target_elem = found[target]
                else:
                    raise TC14NavigationError(f"Result index {target} out of range ({len(found)} results)")
            else:
                target_elem = results[target].web_element
                target_uid = results[target].uid
        elif isinstance(target, str):
            # If target looks like a UID
            if len(target) > 16 and not target.isalnum():
                target_uid = target
            else:
                # Search for it first
                found_results = self.search_part(target)
                return self.open_item(0, search_results=found_results)

        # Try navigating via UID directly if available
        if target_uid:
            direct_obj_url = TC14URLs.get_object_url(self.base_url, target_uid)
            logger.info("Opening item via direct object URL: %s", direct_obj_url)
            driver.get(direct_obj_url)
        elif target_elem:
            try:
                # Look for open command button inside the element
                open_btn = target_elem.find_element(*TC14Selectors.SEARCH_ITEM_OPEN_BTN)
                logger.info("Opening item by clicking Awp0ShowObjectCell button...")
                open_btn.click()
            except (NoSuchElementException, WebDriverException):
                # Fallback to clicking title link
                title_link = target_elem.find_element(*TC14Selectors.SEARCH_ITEM_TITLE)
                logger.info("Opening item by clicking title link...")
                title_link.click()
        else:
            # Fallback: click first available open button on page
            by_btn, sel_btn = TC14Selectors.SEARCH_ITEM_OPEN_BTN
            first_open_btn = wait.until(EC.element_to_be_clickable((by_btn, sel_btn)))
            first_open_btn.click()

        # Wait for object view tabs to appear
        try:
            wait.until(EC.presence_of_element_located(TC14Selectors.NAVIGATION_TABS))
            logger.info("Item object view loaded. Available tabs visible.")
            self.session.touch()
            return True
        except TimeoutException as exc:
            raise TC14NavigationError("Timed out waiting for item object view to load.") from exc

    def navigate_to_content_tab(self) -> bool:
        """Navigate to the 'Content' tab to display the occurrence tree table.

        Returns:
            True if Content tab is loaded.

        Raises:
            TC14NavigationError: If Content tab is not found or fails to load.
        """
        driver = self.driver
        wait = WebDriverWait(driver, self.timeout)

        logger.info("Switching to 'Content' tab...")
        try:
            by_tab, sel_tab = TC14Selectors.CONTENT_TAB
            content_tab = wait.until(EC.element_to_be_clickable((by_tab, sel_tab)))
            content_tab.click()

            # Wait for table container to mount
            by_tbl, sel_tbl = TC14Selectors.TREE_TABLE_CONTAINER
            wait.until(EC.presence_of_element_located((by_tbl, sel_tbl)))
            time.sleep(2)  # Allow initial virtualized rows to render
            logger.info("Content tab activated and occurrence tree table mounted.")
            self.session.touch()
            return True
        except TimeoutException as exc:
            # Check if tabs exist and log available tab names for diagnostics
            tabs = driver.find_elements(*TC14Selectors.NAVIGATION_TABS)
            tab_names = [t.text.strip() for t in tabs if t.text.strip()]
            logger.error("Content tab not found. Visible tabs: %s", tab_names)
            raise TC14NavigationError(
                f"Failed to navigate to Content tab. Available tabs: {tab_names}"
            ) from exc

    # -------------------------------------------------------------------------
    # BOM Excel Export & File Handling
    # -------------------------------------------------------------------------

    def export_bom_excel(
        self,
        output_dir: Optional[Path] = None,
        timeout: int = 60,
        target_filename: Optional[str] = None,
    ) -> Path:
        """Trigger native 'Awp0ExportToExcel' and wait for downloaded workbook.

        Args:
            output_dir: Target folder to store downloaded file. Defaults to self.download_dir.
            timeout: Maximum seconds to wait for export generation and download.
            target_filename: Optional custom name for the resulting file (e.g. 'PLM_110C103NL0.xlsx').

        Returns:
            Path to the downloaded and finalized .xlsx file.

        Raises:
            TC14ExportError: If export button is missing or download fails.
        """
        driver = self.driver
        wait = WebDriverWait(driver, self.timeout)
        target_dir = Path(output_dir or self.download_dir).resolve()
        target_dir.mkdir(parents=True, exist_ok=True)

        # Snapshot download folder before trigger to detect new arrivals
        before_files = set(self.download_dir.glob("*.xlsx"))

        # Click Export to Excel command button
        try:
            by_cmd, sel_cmd = TC14Selectors.EXPORT_EXCEL_COMMAND
            export_cmd = wait.until(EC.element_to_be_clickable((by_cmd, sel_cmd)))
            export_cmd.click()
            logger.info("Clicked 'Awp0ExportToExcel' command button.")
        except TimeoutException as exc:
            raise TC14ExportError("Export to Excel button ('Awp0ExportToExcel') not found in view.") from exc

        # Wait for export slide-out dialog
        try:
            by_diag, sel_diag = TC14Selectors.EXPORT_DIALOG
            wait.until(EC.presence_of_element_located((by_diag, sel_diag)))
            logger.info("Export slide-out dialog opened.")
        except TimeoutException as exc:
            raise TC14ExportError("Export dialog did not open after clicking export command.") from exc

        # Click the Export confirmation button inside dialog
        try:
            by_conf, sel_conf = TC14Selectors.EXPORT_CONFIRM_BTN
            confirm_btn = wait.until(EC.element_to_be_clickable((by_conf, sel_conf)))
            confirm_btn.click()
            logger.info("Clicked 'Export' confirmation button inside dialog.")
        except TimeoutException as exc:
            raise TC14ExportError("Export confirm button inside dialog not clickable.") from exc

        # Wait for file download to complete
        logger.info("Waiting up to %d seconds for Excel download to complete...", timeout)
        downloaded_file = self._poll_for_downloaded_file(
            directory=self.download_dir,
            initial_files=before_files,
            timeout=timeout,
        )

        if not downloaded_file:
            raise TC14ExportError(
                f"Export timed out after {timeout} seconds. No complete .xlsx file found in {self.download_dir}."
            )

        logger.info(
            "Detected downloaded Excel file: %s (size: %d bytes)",
            downloaded_file.name,
            downloaded_file.stat().st_size,
        )

        # Destination routing
        final_dest = downloaded_file
        if target_filename:
            dest_name = target_filename if target_filename.endswith(".xlsx") else f"{target_filename}.xlsx"
            final_dest = target_dir / dest_name
            shutil.move(str(downloaded_file), str(final_dest))
            logger.info("Moved and renamed export file to: %s", final_dest)
        elif target_dir != self.download_dir:
            final_dest = target_dir / downloaded_file.name
            shutil.move(str(downloaded_file), str(final_dest))
            logger.info("Moved export file to target directory: %s", final_dest)

        self.session.touch()
        return final_dest

    def scrape_bom_tree(self) -> List[Any]:
        """Scrape visible DOM elements representing BOM tree nodes."""
        driver = self.driver
        by_item, sel_item = TC14Selectors.SEARCH_RESULT_ITEMS
        elements = driver.find_elements(by_item, sel_item)
        results = []
        for elem in elements:
            if hasattr(elem, "click"):
                elem.click()
            results.append(elem.text if hasattr(elem, "text") else str(elem))
        return results

    def _poll_for_downloaded_file(
        self,
        directory: Path,
        initial_files: set,
        timeout: int = 60,
        poll_interval: float = 0.5,
    ) -> Optional[Path]:
        """Poll folder until a new non-temporary .xlsx file appears and finishes writing."""
        start_time = time.time()
        while time.time() - start_time < timeout:
            current_files = set(directory.glob("*.xlsx"))
            new_files = current_files - initial_files

            # Filter out temporary downloads
            valid_candidates = [
                f for f in new_files
                if not f.name.endswith(".tmp")
                and not f.name.endswith(".crdownload")
                and not f.name.endswith(".part")
            ]

            if valid_candidates:
                candidate = valid_candidates[0]
                # Verify file is not actively being written to
                initial_size = candidate.stat().st_size
                time.sleep(0.5)
                if candidate.exists() and candidate.stat().st_size == initial_size and initial_size > 0:
                    return candidate

            time.sleep(poll_interval)

        return None

    # -------------------------------------------------------------------------
    # Fallback Interactive Tree Scraper & Synthesizer
    # -------------------------------------------------------------------------

    def scrape_bom_tree_interactive(
        self, max_depth: int = 6, expand_children: bool = True
    ) -> List[BOMItem]:
        """Scrape the BOM occurrence hierarchy directly from the UI tree table.

        Serves as a resilient fallback if native SOA export is disabled by role or unavailable.

        Args:
            max_depth: Maximum tree depth (Level 1 to 6).
            expand_children: Whether to click collapsed node toggles to reveal sub-assemblies.

        Returns:
            List of BOMItem objects representing the hierarchical structure.
        """
        driver = self.driver
        logger.info("Beginning interactive BOM tree scraping (max_depth=%d)...", max_depth)

        if expand_children:
            self._expand_all_visible_tree_nodes(max_depth=max_depth)

        # Scrape all rendered rows
        by_rows, sel_rows = TC14Selectors.TREE_PINNED_ROWS
        row_elements = driver.find_elements(by_rows, sel_rows)
        logger.info("Found %d rows in BOM tree table.", len(row_elements))

        items: List[BOMItem] = []
        for elem in row_elements:
            try:
                raw_text = elem.text.strip()
                if not raw_text:
                    continue

                level_str = elem.get_attribute(TC14Selectors.TREE_ROW_LEVEL_ATTR) or "1"
                try:
                    level = int(level_str)
                except ValueError:
                    level = 1

                expanded_attr = elem.get_attribute(TC14Selectors.TREE_ROW_EXPANDED_ATTR) or ""
                inner_html = elem.get_attribute("innerHTML") or ""
                has_children = (
                    expanded_attr.lower() in ("true", "expanded")
                    or "Show Children" in inner_html
                )

                # Extract column text parts
                cell_elements = elem.find_elements(*TC14Selectors.TREE_CELL_TEXT)
                cell_texts = [c.text.strip() for c in cell_elements if c.text.strip()]

                item_id = cell_texts[0] if cell_texts else ""
                item_name = cell_texts[1] if len(cell_texts) > 1 else ""

                items.append(
                    BOMItem(
                        level=level,
                        item_id=item_id,
                        item_name=item_name,
                        has_children=has_children,
                        raw_text=raw_text,
                    )
                )
            except (StaleElementReferenceException, NoSuchElementException):
                continue

        logger.info("Scraped %d BOM items interactively.", len(items))
        self.session.touch()
        return items

    def _expand_all_visible_tree_nodes(self, max_depth: int = 6) -> None:
        """Iteratively click 'Show Children' toggles up to max_depth."""
        driver = self.driver
        for depth in range(1, max_depth):
            by_exp, sel_exp = TC14Selectors.TREE_EXPAND_BTN
            expand_buttons = driver.find_elements(by_exp, sel_exp)
            visible_toggles = [b for b in expand_buttons if b.is_displayed()]
            if not visible_toggles:
                break

            logger.info("Expanding %d tree toggles at depth %d...", len(visible_toggles), depth)
            for toggle in visible_toggles:
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", toggle)
                    toggle.click()
                    time.sleep(0.3)
                except Exception:
                    continue
            time.sleep(2)  # Wait for children to load

    def synthesize_plm_excel_from_tree(
        self, items: List[BOMItem], output_path: Path
    ) -> Path:
        """Generate a compliant 14-column PLM BOM Excel workbook from scraped tree items.

        Columns follow the exact structure expected by downstream 'locbomfull.bas':
        - Col B (idx 2): Level (1..6)
        - Col D (idx 4): Item ID / Part Number
        - Col E (idx 5): hasChildren (True/False)
        - Col I (idx 9): Effectivity
        - Col K (idx 11): Item Name
        """
        output_path = Path(output_path).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        headers = [
            "Col_A",
            "Level",
            "Col_C",
            "Part_Number",
            "hasChildren",
            "Col_F",
            "Col_G",
            "Col_H",
            "Effectivity",
            "Col_J",
            "Item_Name",
            "Col_L",
            "Col_M",
            "Col_N",
        ]
        ws.append(headers)

        for item in items:
            row = [
                "",
                item.level,
                "",
                item.item_id,
                "True" if item.has_children else "False",
                "",
                "",
                "",
                item.effectivity,
                "",
                item.item_name,
                "",
                "",
                "",
            ]
            ws.append(row)

        wb.save(str(output_path))
        logger.info("Synthesized 14-column PLM Excel file with %d items: %s", len(items), output_path)
        return output_path

    # -------------------------------------------------------------------------
    # Public Project Interface Contract: download_bom_full
    # -------------------------------------------------------------------------

    def download_bom_full(self, part_number: str, output_dir: Path) -> Path:
        """Download complete BOM Full for the specified part number.

        Implements the exact contract defined in PROJECT.md:
        `TC14Client.download_bom_full(part_number: str, output_dir: Path) -> Path`

        Workflow:
        1. Ensures authenticated session with Teamcenter.
        2. Executes direct search navigation for part_number.
        3. Opens the primary product revision.
        4. Navigates to Content tab.
        5. Attempts native Excel export ('Awp0ExportToExcel').
        6. Falls back to interactive tree scraper + synthesis if export fails.
        7. Returns final Path to 'PLM_<part_number>.xlsx'.

        Args:
            part_number: Target machine assembly or part code (e.g. '110C103NL0').
            output_dir: Destination folder for the PLM file.

        Returns:
            Path to downloaded / synthesized 'PLM_<part_number>.xlsx'.
        """
        resolved_out_dir = Path(output_dir).resolve()
        resolved_out_dir.mkdir(parents=True, exist_ok=True)
        target_filename = f"PLM_{part_number}.xlsx"

        logger.info("Initiating download_bom_full for part '%s' into '%s'...", part_number, resolved_out_dir)

        # 1. Authenticate
        self.ensure_authenticated()

        # 2. Search
        search_results = self.search_part(part_number)

        # 3. Open primary result
        self.open_item(0, search_results=search_results)

        # 4. Content Tab
        self.navigate_to_content_tab()

        # 5. Attempt Native Export
        try:
            exported_path = self.export_bom_excel(
                output_dir=resolved_out_dir,
                timeout=45,
                target_filename=target_filename,
            )
            logger.info("Successfully exported native BOM Excel: %s", exported_path)
            return exported_path
        except TC14ExportError as exc:
            logger.warning("Native Excel export failed (%s). Falling back to interactive scraper...", exc)

        # 6. Fallback Scraper + Synthesis
        scraped_items = self.scrape_bom_tree_interactive(max_depth=6)
        fallback_path = resolved_out_dir / target_filename
        self.synthesize_plm_excel_from_tree(scraped_items, fallback_path)
        return fallback_path

    # -------------------------------------------------------------------------
    # Teardown & Context Management
    # -------------------------------------------------------------------------

    def close(self) -> None:
        """Close browser session and cleanup resources."""
        self.session.close()

    def __enter__(self) -> "TC14AutomationClient":
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()


class TC14Client(TC14AutomationClient):
    """Direct alias conforming to PROJECT.md interface contract."""
