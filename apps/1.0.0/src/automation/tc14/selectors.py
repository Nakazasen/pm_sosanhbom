"""Verified DOM selectors, XPath expressions, and URL routes for Teamcenter Active Workspace (TC14).

Based on verified live probes of Teamcenter Active Workspace Version 2412 at http://tcmp3gwb:3000/.
"""

from typing import Tuple
from urllib.parse import quote
from selenium.webdriver.common.by import By


class TC14URLs:
    """URL definitions and route templates for Active Workspace."""

    DEFAULT_BASE_URL: str = "http://tcmp3gwb:3000/"
    HOME_ROUTE: str = "#/showHome"
    SEARCH_ROUTE_TEMPLATE: str = (
        "#/teamcenter.search.search?searchCriteria={criteria}&secondaryCriteria=*&isGlobalSearch=true"
    )
    OBJECT_ROUTE_TEMPLATE: str = (
        "#/com.siemens.splm.clientapi.tcui.xrt.showObject?uid={uid}"
    )

    @classmethod
    def search_url(cls, search_criteria: str, base_url: str = DEFAULT_BASE_URL) -> str:
        """Alias for get_search_url accepting criteria first with default base_url."""
        return cls.get_search_url(base_url=base_url, search_criteria=search_criteria)

    @classmethod
    def get_search_url(cls, base_url: str, search_criteria: str) -> str:
        """Construct direct hash search navigation URL.

        Args:
            base_url: The root URL of TC14 (e.g. 'http://tcmp3gwb:3000/').
            search_criteria: The item/part code to search (e.g. '110C103NL0').

        Returns:
            Fully-qualified URL string with properly encoded criteria.
        """
        clean_base = base_url.rstrip("/") + "/"
        encoded_criteria = quote(search_criteria.strip())
        return clean_base + cls.SEARCH_ROUTE_TEMPLATE.format(criteria=encoded_criteria)

    @classmethod
    def get_object_url(cls, base_url: str, uid: str) -> str:
        """Construct direct hash object view navigation URL."""
        clean_base = base_url.rstrip("/") + "/"
        encoded_uid = quote(uid.strip())
        return clean_base + cls.OBJECT_ROUTE_TEMPLATE.format(uid=encoded_uid)

    @classmethod
    def get_home_url(cls, base_url: str) -> str:
        """Construct home page URL."""
        clean_base = base_url.rstrip("/") + "/"
        return clean_base + cls.HOME_ROUTE


class TC14Selectors:
    """DOM, CSS, and XPath selectors verified against live TC14 instance."""

    # --- Authentication ---
    USERNAME_INPUT: Tuple[str, str] = (By.CSS_SELECTOR, "input[name='username']")
    USERNAME_FALLBACK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[data-locator='User Name'], input[placeholder='User Name']",
    )
    PASSWORD_INPUT: Tuple[str, str] = (By.CSS_SELECTOR, "input[name='password']")
    PASSWORD_FALLBACK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[data-locator='Password'], input[placeholder='Password']",
    )
    LOGIN_BUTTON: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-login-signInButton button, .aw-login-signInButton button[type='submit'], button[type='submit']",
    )
    LOGIN_ERROR: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-login-error, .sw-login-error, .aw-error-reverse",
    )
    LOGIN_CONTAINER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-layout-loginContent, .aw-login-container, input[name='username']",
    )

    # --- Header & User Banner ---
    BANNER_HEADER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "header.aw-layout-header, .aw-layout-header, header[role='banner']",
    )
    BANNER_USER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "[class*='tcBanner_vn_pe03'], [data-locator='tcBanner'], .aw-header-subtitle",
    )
    APP_MAIN_VIEW: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "aw-layout-mainView, .aw-layout-workarea, .aw-layout-mainContainer",
    )

    # --- Global & Header Search ---
    SEARCH_INPUT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input.aw-uiwidgets-searchBox[name='searchBox'], input.aw-uiwidgets-searchBox",
    )
    SEARCH_CONTAINER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-uiwidgets-searchBoxContainer",
    )
    SEARCH_RESULT_ITEMS: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "li.aw-widgets-cellListItem",
    )
    SEARCH_ITEM_TITLE: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "a.aw-widgets-cellListCellTitle[data-locator='aw-clickable-title'], a.aw-widgets-cellListCellTitle",
    )
    SEARCH_ITEM_OPEN_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button.aw-commandId-Awp0ShowObjectCell, button[command-id='Awp0ShowObjectCell']",
    )
    NO_RESULTS_LABEL: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-widgets-noResultsLabel, .aw-layout-noResults",
    )

    # --- Object Navigation & Tabs ---
    NAVIGATION_TABS: Tuple[str, str] = (By.CSS_SELECTOR, "li.sw-tab, .sw-tab")
    CONTENT_TAB: Tuple[str, str] = (
        By.XPATH,
        "//li[contains(@class, 'sw-tab') and (.//span[normalize-space()='Content'] or contains(., 'Content'))]",
    )
    REPORTS_TAB: Tuple[str, str] = (
        By.XPATH,
        "//li[contains(@class, 'sw-tab') and (.//span[normalize-space()='Reports'] or contains(., 'Reports'))]",
    )
    OVERVIEW_TAB: Tuple[str, str] = (
        By.XPATH,
        "//li[contains(@class, 'sw-tab') and (.//span[normalize-space()='Overview'] or contains(., 'Overview'))]",
    )

    # --- Native Excel Export ---
    EXPORT_EXCEL_COMMAND: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0ExportToExcel']",
    )
    SELECT_ALL_COMMAND: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0SelectAllObjectSet']",
    )
    EXPORT_DIALOG: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.sw-right-dialog, div.aw-layout-popup, div.sw-popup, div.sw-popupContent",
    )
    EXPORT_CONFIRM_BTN: Tuple[str, str] = (
        By.XPATH,
        "//div[contains(@class,'sw-right-dialog') or contains(@class,'sw-popup') "
        "or contains(@class,'aw-layout-popup')]//button[normalize-space()='Export']",
    )
    EXPORT_DIALOG_CANCEL_BTN: Tuple[str, str] = (
        By.XPATH,
        "//div[contains(@class,'sw-right-dialog') or contains(@class,'sw-popup') "
        "or contains(@class,'aw-layout-popup')]//button[normalize-space()='Cancel']",
    )
    PROGRESS_INDICATOR: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-layout-progressIndicator, .sw-progress-indicator, .aw-widgets-progress",
    )
    TOAST_NOTIFICATION: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".noty_message, .sw-toast, .noty_text",
    )

    # --- Occurrence Tree & Interactive Scraper ---
    TREE_TABLE_CONTAINER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-table, div.occTreeTable, div.aw-jswidgets-grid, [data-locator='aw-splm-table']",
    )
    TREE_PINNED_ROWS: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-splm-tableRow.aw-splm-tablePinnedRow, .aw-splm-tableRow",
    )
    TREE_EXPAND_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-jswidgets-treeExpandCollapseCmd[title='Show Children'], "
        "[icon-id*='miscCollapsedTree'], [icon-id*='miscRightArrow']",
    )
    TREE_COLLAPSE_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-jswidgets-treeExpandCollapseCmd[title='Hide Children'], "
        "[icon-id*='miscExpandedTree'], [icon-id*='miscDownArrow']",
    )
    TREE_CELL_TEXT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-splm-tableCellText, .aw-jswidgets-tableCell",
    )
    TREE_ROW_LEVEL_ATTR: str = "aria-level"
    TREE_ROW_EXPANDED_ATTR: str = "aria-expanded"
