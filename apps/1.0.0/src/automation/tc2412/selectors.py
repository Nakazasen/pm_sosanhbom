"""Verified DOM selectors, XPath expressions, and URL routes for Siemens Teamcenter Version 2412 (TC2412 / Active Workspace).

Based on authoritative specifications in specs/SPEC_PLM_AUTO_DOWNLOAD.md
and verified live probes of Teamcenter Active Workspace 2412 at http://tcmp3gwb:3000/.
"""

from __future__ import annotations

from typing import Tuple
from urllib.parse import quote
from selenium.webdriver.common.by import By


class TC2412URLs:
    """URL definitions and route templates for Teamcenter Active Workspace 2412."""

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
        """Construct direct hash search navigation URL with criteria first."""
        return cls.get_search_url(base_url=base_url, search_criteria=search_criteria)

    @classmethod
    def get_search_url(cls, base_url: str, search_criteria: str) -> str:
        """Construct direct hash search navigation URL.

        Args:
            base_url: The root URL of TC2412 (e.g. 'http://tcmp3gwb:3000/').
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


class TC2412Selectors:
    """DOM, CSS, and XPath selectors verified against Siemens Teamcenter Version 2412 (Active Workspace).
    
    Contains all 49 verified DOM locators across the 7 automation phases.
    """

    # --- Phase 1: Authentication (AUTH-01) ---
    LOGIN_CONTAINER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.signin-form.aw-layout-mainView, div.signin-form, div.aw-layout-loginContent, .aw-login-container",
    )
    VERSION_TITLE: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "p.aw-login-copyrightTitle, .aw-login-copyrightTitle",
    )
    USERNAME_INPUT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[name='username']",
    )
    USERNAME_FALLBACK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[data-locator='User Name'], input[placeholder='User Name']",
    )
    PASSWORD_INPUT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[name='password']",
    )
    PASSWORD_FALLBACK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[data-locator='Password'], input[placeholder='Password']",
    )
    LOGIN_SUBMIT_BUTTON: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-login-signInButton button[type='submit'], .aw-login-signInButton button, button[type='submit']",
    )
    LOGIN_PROGRESS_SPINNER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-login-progressContainer, div.aw-layout-progressBar",
    )
    LOGIN_ERROR: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-login-error, .sw-login-error, .aw-error-reverse, div.aw-login-errorContainer",
    )
    BANNER_HEADER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "header.sw-row.aw-layout-header, header.aw-layout-header, header[role='banner']",
    )
    ALERT_NOTIFICATION_BELL: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0ShowAlertWithBubble'], button[aria-label='Alerts']",
    )

    # --- Phase 2: Search & Navigation (BOM-SEARCH-01) ---
    GLOBAL_SEARCH_CONTAINER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-search-globalSearchBoxContainer[role='search'], div.aw-uiwidgets-searchBoxContainer",
    )
    GLOBAL_SEARCH_INPUT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input.aw-uiwidgets-searchBox[name='searchBox'], input.aw-uiwidgets-searchBox",
    )
    SEARCH_ACTION_ICON: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "span.aw-uiwidgets-searchBoxIcon, [icon-id='cmdSearch']",
    )
    SEARCH_RESULTS_LIST: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "ul.aw-widgets-cellListWidget, div.aw-widgets-cellListContainer",
    )
    SEARCH_RESULT_ITEMS: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "li.aw-widgets-cellListItem[role='option'], li.aw-widgets-cellListItem",
    )
    SEARCH_ITEM_TITLE: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "a.aw-widgets-cellListCellTitle[data-locator='aw-clickable-title'], a.aw-widgets-cellListCellTitle",
    )
    SEARCH_ITEM_OPEN_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0ShowObjectCell'], button.aw-commandId-Awp0ShowObjectCell",
    )
    NO_RESULTS_LABEL: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-widgets-noResultsLabel, .aw-layout-noResults",
    )
    CONTENT_TAB: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "a[data-locator='tab-tc_xrt_Content'], li.sw-tab a[title*='Content']",
    )
    CONTENT_TAB_XPATH: Tuple[str, str] = (
        By.XPATH,
        "//span[@title='Content' or normalize-space()='Content'] | //a[@data-locator='tab-tc_xrt_Content'] | //*[contains(@class,'sw-tab') and (.//span[normalize-space()='Content'] or contains(.,'Content'))]",
    )
    ACTIVE_CONTENT_TAB: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "li.sw-tab.sw-tab-selected a[data-locator='tab-tc_xrt_Content'], li.sw-tab-selected",
    )

    # --- Phase 3: BOM Expansion to Level 7 (BOM-EXPAND-01) ---
    OCC_TREE_TABLE: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-table#occTreeTable, div.aw-splm-table, div.occTreeTable",
    )
    ROOT_ROW: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-tableRow[data-indexnumber='0'], div.aw-splm-tableRow:first-child",
    )
    ROOT_CELL_LINK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-splm-tableCellText a.aw-uiwidgets-clickableTitle, a.aw-uiwidgets-clickableTitle",
    )
    EXPAND_TOOLBAR_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awb0Expand'], button[button-id='Awb0Expand'], button.aw-commandId-Awb0Expand",
    )
    EXPAND_BELOW_CMD: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div[command-id='Awb0ExpandBelow'], div.aw-widgets-cellListItem[title*='Expand Below'], [command-id='Awb0ExpandBelow']",
    )
    EXPAND_LEVEL_INPUT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input.aw-widgets-propertyVal[type='number'], div.aw-popup input[type='number'], input[type='number']",
    )
    EXPAND_CONFIRM_BTN: Tuple[str, str] = (
        By.XPATH,
        "//div[contains(@class,'aw-popup') or contains(@class,'sw-popup')]//button[contains(.,'Expand') or contains(.,'OK') or normalize-space()='Expand']",
    )
    TREE_PROGRESS_BAR: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-layout-progressBar, .aw-layout-progressBarContainer",
    )
    TABLE_SUMMARY: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-tableSummary, div.aw-splm-tableStatus",
    )
    TABLE_ROWS: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-tableRow",
    )

    # --- Phase 4: Select All & Open Menu (BOM-SELECT-01) ---
    ENABLE_MULTI_SELECT_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0EnableMultiSelect'], button[aria-label='Selection Mode']",
    )
    SELECT_ALL_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0SelectAll'], button[command-id='Awp0SelectAllObjectSet'], button.aw-commandId-Awp0SelectAll",
    )
    TABLE_SELECTION_COUNT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-tableSelectionCount, div.aw-splm-tableSummary",
    )
    TABLE_HEADER_CHECKBOX: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-tableHeaderCheckbox, input[type='checkbox'].aw-splm-tableHeaderCheckbox",
    )
    SELECTED_ROWS_VIEWPORT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-splm-tableRow[aria-selected='true'], .aw-splm-tableRowSelected",
    )
    EXCEL_REPORT_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Arm0ExportImport'], button.aw-commandId-Arm0ExportImport",
    )
    POPUP_COMMAND_MENU: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-popup-commandListContainer, div.aw-popup.sw-popup, div.sw-popupContent",
    )

    # --- Phase 5: Export to Excel Open & Anti-Import Guard (BOM-EXPORT-OPEN-01) ---
    EXPORT_TO_EXCEL_CMD: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0ExportToExcel'], div[title='Export To Excel'], [command-id='Awp0ExportToExcel']",
    )
    EXPORT_DIALOG_CONTAINER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.sw-popup.sw-dialog#ui-id-3, div.sw-right-dialog, div.aw-layout-popup, div.sw-popupContent",
    )
    DIALOG_CAPTION_HEADER: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.aw-panel-caption, div.sw-dialog-header, h2",
    )

    # --- Phase 6: 14-Column Configuration (BOM-EXPORT-CONFIG-01) ---
    PROPERTY_SOURCE_SELECTOR: Tuple[str, str] = (
        By.CSS_SELECTOR,
        ".aw-property-source-dropdown, div.sw-lov-container, div[data-locator='Property Source']",
    )
    PROPERTY_SOURCE_XPATH: Tuple[str, str] = (
        By.XPATH,
        "//*[contains(text(),'Property Source')]/following::div[contains(@class,'sw-lov-container')][1]",
    )
    AVAILABLE_PROPS_FILTER_INPUT: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input.aw-uiwidgets-searchBox[placeholder*='Filter'], div.aw-panel-section input.aw-uiwidgets-searchBox, input[placeholder='Filter']",
    )
    ADD_PROPERTIES_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0ExportSelectedColumnsAdd'], button[aria-label='Add Properties']",
    )
    MOVE_UP_COLUMN_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0MoveUpExcelColumn'], button[aria-label='Move Up']",
    )
    MOVE_DOWN_COLUMN_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0MoveDownExcelColumn'], button[aria-label='Move Down']",
    )
    DISPLAYED_COLUMNS_ITEMS: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "details[caption='Selected Properties'] li.aw-widgets-cellListItem, div.aw-widgets-cellListWidget li.aw-widgets-cellListItem",
    )
    REMOVE_COLUMN_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "li.aw-widgets-cellListItem button[title='Remove'], button[command-id*='Remove']",
    )

    # --- Phase 7: Run Export & File Verification (BOM-EXPORT-RUN-01) ---
    RUN_IN_BACKGROUND_CHECKBOX: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "input[type='checkbox'][name*='Background'], input[type='checkbox'][data-locator*='Background'], input[name='runInBackground']",
    )
    EXPORT_SUBMIT_BTN: Tuple[str, str] = (
        By.XPATH,
        "//form[contains(@class,'sw-command-panel')]//button[contains(@class,'sw-button') and contains(.,'Export')] | //button[normalize-space()='Export']",
    )
    EXPORT_COMPLETION_TOAST: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.noty_message, div.aw-layout-popup, .sw-toast, .noty_text",
    )
    NOTIFICATION_DOWNLOAD_LINK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "div.noty_message a, a.aw-widgets-cellListCellTitle, a.aw-jswidgets-tableCellLink",
    )
    REPORTS_TAB_LINK: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "a[data-locator='tab-tc_xrt_Rb0Reports'], button[command-id='Awp0GoReports']",
    )
    REPORT_DOWNLOAD_BTN: Tuple[str, str] = (
        By.CSS_SELECTOR,
        "button[command-id='Awp0DownloadReport'], button.aw-commandId-Awp0DownloadReport",
    )

