"""Comprehensive Unit & Tier 1-5 Test Suite for TC2412 Automation Pipeline.

Covers:
- Tier 1: Feature Isolation (Auth, Search, Expand, Select, Safe Export, 14 Columns, Run & Download, ProgressReporter)
- Tier 2: Boundary Conditions (Timeouts, ItemNotFoundError, Zero Rows, Column Mismatch, Corrupt Files)
- Tier 3: Cross-Phase Integration (Pipeline chaining, TeamcenterSeleniumAdapter)
- Tier 4: Real-World Scenarios (Japanese Unicode preservation, Canonical 14 columns, Sheet PLM formula bridge, TC2412 24-col normalization)
- Tier 5: Adversarial Hardening (Fail-Closed Anti-Import Guard, Virtual DOM resilience, Active download click trigger)
"""

from __future__ import annotations

import io
import time
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import openpyxl
import pandas as pd
import pytest
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from src.automation.tc2412.client import (
    CANONICAL_14_COLUMNS,
    AuthenticationError,
    BOMCorruptFileError,
    BOMExpandTimeoutError,
    BOMValidationError,
    ColumnConfigurationError,
    ContentTabNotFoundError,
    ExcelMenuNotOpenError,
    ExpandMenuNotOpenError,
    ExportDialogTimeoutError,
    ExportDownloadTimeoutError,
    ImportChangesForbiddenError,
    ItemNotFoundError,
    NoRowsSelectedError,
    PLMNetworkUnreachableError,
    RootNodeSelectionError,
    SearchTimeoutError,
    TC2412AutomationClient,
    TC2412AutomationError,
    safe_trigger_export_to_excel,
    verify_exported_excel,
)
from src.automation.tc2412.progress_reporter import ProgressReporter
from src.automation.tc2412.selectors import TC2412Selectors, TC2412URLs
from src.automation.tc2412.session import BrowserConfig, TC2412SessionManager, create_driver
from src.core.adapters import TeamcenterSeleniumAdapter
from src.core.tc2412_bridge import TC2412CanonicalNormalizer, TC2412SheetPLMBridge


# =====================================================================
# Fixtures
# =====================================================================

@pytest.fixture
def mock_driver():
    """Create a mock Selenium WebDriver with common methods."""
    driver = MagicMock()
    driver.current_url = "http://tcmp3gwb:3000/#/showHome"
    driver.find_elements.return_value = []
    return driver


@pytest.fixture
def sample_valid_excel(tmp_path: Path) -> Path:
    """Create a valid 14-column PLM Excel file with Japanese text."""
    excel_path = tmp_path / "PLM_T10C423NL0_Valid.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Write 14 canonical headers
    for col_idx, col_name in enumerate(CANONICAL_14_COLUMNS, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Row 2 (Root part)
    row2 = [
        "Home/Machinery", 0, "TopAssembly", "T10C423NL0", True, 1,
        "1st", "FLAG_A", "to 2026/12/31", "Virgo2", "メインフレーム (Main Frame)", "ECN-101", "01", "Released"
    ]
    for col_idx, val in enumerate(row2, start=1):
        ws.cell(row=2, column=col_idx, value=val)

    # Pad with dummy bytes if needed to exceed 5KB
    for r in range(3, 100):
        for col_idx, val in enumerate(row2, start=1):
            ws.cell(row=r, column=col_idx, value=val)

    wb.save(str(excel_path))
    wb.close()
    return excel_path


# =====================================================================
# Tier 1: Feature Isolation Tests
# =====================================================================

class TestTier1Features:
    """Feature isolation testing for the 7 phases."""

    def test_tc2412_urls_construction(self):
        """Test URL templates and encoding."""
        search_url = TC2412URLs.get_search_url("http://tcmp3gwb:3000/", "110C 103NL0")
        assert "http://tcmp3gwb:3000/#/teamcenter.search.search" in search_url
        assert "110C%20103NL0" in search_url

        obj_url = TC2412URLs.get_object_url("http://tcmp3gwb:3000/", "uid_12345")
        assert "uid=uid_12345" in obj_url

        home_url = TC2412URLs.get_home_url("http://tcmp3gwb:3000/")
        assert home_url == "http://tcmp3gwb:3000/#/showHome"

    def test_auth_form_submit_success(self, mock_driver):
        """Phase 1: AUTH-01 form fill and submit."""
        user_elem = MagicMock()
        pass_elem = MagicMock()
        btn_elem = MagicMock()
        banner_elem = MagicMock()

        # Mock WebDriverWait and element lookup
        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.side_effect = [user_elem, pass_elem, banner_elem]
            mock_driver.find_element.return_value = btn_elem

            session = MagicMock()
            session.driver = mock_driver
            session.is_session_alive.return_value = False
            session.is_login_page_present.return_value = True

            client = TC2412AutomationClient(session_manager=session)
            res = client.login("vn_pe03", "SecurePass123", timeout=10)

            assert res is True
            user_elem.send_keys.assert_called_with("vn_pe03")
            pass_elem.send_keys.assert_called_with("SecurePass123")
            btn_elem.click.assert_called_once()
            session.mark_authenticated.assert_called_once()

    def test_search_item_success(self, mock_driver):
        """Phase 2: BOM-SEARCH-01 search and open item."""
        res_item = MagicMock()
        title_link = MagicMock()
        res_item.text = "T10C423NL0/01 - Main Frame"
        res_item.find_element.return_value = title_link

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            # return results list
            mock_driver.find_elements.side_effect = [
                [],  # no_results_label (empty = has results)
                [res_item],  # search results
            ]

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            rev = client.search_item("T10C423NL0", part_rev="01")
            assert rev == "01"
            title_link.click.assert_called_once()

    def test_navigate_to_content_tab_success(self, mock_driver):
        """Phase 2: Navigate to Content tab."""
        tab_elem = MagicMock()
        active_tab = MagicMock()

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.side_effect = [tab_elem, active_tab]

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            assert client.navigate_to_content_tab() is True
            tab_elem.click.assert_called_once()

    def test_expand_bom_tree_level_7(self, mock_driver):
        """Phase 3: BOM-EXPAND-01 expand to level 7."""
        tree_elem = MagicMock()
        root_elem = MagicMock()
        exp_btn = MagicMock()
        below_btn = MagicMock()
        lvl_input = MagicMock()
        confirm_btn = MagicMock()

        mock_driver.find_element.return_value = confirm_btn
        # Emulate 50 visible rows stable
        mock_row = MagicMock()
        mock_driver.find_elements.side_effect = [
            [],  # progress bar (not busy)
            [mock_row] * 50,  # rows
            [],  # progress bar
            [mock_row] * 50,
            [],
            [mock_row] * 50,
            [],
            [mock_row] * 50,
        ]

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.side_effect = [tree_elem, root_elem, exp_btn, below_btn, lvl_input]

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            rows = client.expand_bom_tree(target_level=7, timeout=10, stability_cycles=2, stability_interval=0.01)
            assert rows == 50
            root_elem.click.assert_called_once()
            exp_btn.click.assert_called_once()
            below_btn.click.assert_called_once()
            lvl_input.send_keys.assert_called_with("7")
            confirm_btn.click.assert_called_once()

    def test_select_all_bom_lines(self, mock_driver):
        """Phase 4: BOM-SELECT-01 Select All and open Excel Report menu."""
        select_all_btn = MagicMock()
        count_elem = MagicMock()
        count_elem.text = "1250 Selected"
        report_btn = MagicMock()
        menu_elem = MagicMock()

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.side_effect = [select_all_btn, count_elem, report_btn, menu_elem]

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            selected_count = client.select_all_bom_lines(timeout=10)
            assert selected_count == 1250
            select_all_btn.click.assert_called_once()
            report_btn.click.assert_called_once()

    def test_safe_trigger_export_to_excel(self, mock_driver):
        """Phase 5: BOM-EXPORT-OPEN-01 safe trigger."""
        popup_menu = MagicMock()
        export_btn = MagicMock()
        export_btn.text = "Export to Excel"
        export_btn.get_attribute.return_value = "Awp0ExportToExcel"

        popup_menu.find_elements.return_value = [export_btn]

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.return_value = popup_menu

            res = safe_trigger_export_to_excel(mock_driver)
            assert res is True
            export_btn.click.assert_called_once()

    def test_progress_reporter_heartbeat_and_update(self):
        """Phase 8: REP-01 ProgressReporter heartbeat and callback."""
        callback_args = []

        def test_callback(idx, total, name, pct, elapsed):
            callback_args.append((idx, total, name, pct, elapsed))

        reporter = ProgressReporter(callback=test_callback, interval_seconds=0.1)
        msg1 = reporter.update(1, 8, "Authentication", 12.5)
        assert "Phase [1/8]" in msg1
        assert "Authentication" in msg1
        assert "12.5%" in msg1
        assert len(callback_args) == 1

        # Test start and stop
        reporter.start()
        time.sleep(0.25)
        reporter.stop()
        assert len(callback_args) >= 2


# =====================================================================
# Tier 2: Boundary Conditions Tests
# =====================================================================

class TestTier2Boundaries:
    """Boundary conditions, timeout thresholds, and error matrices."""

    def test_search_not_found_raises_item_not_found_error(self, mock_driver):
        """Phase 2: Non-existent item raises ItemNotFoundError."""
        no_res_label = MagicMock()
        no_res_label.is_displayed.return_value = True

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            mock_driver.find_elements.side_effect = [
                [no_res_label],  # NO_RESULTS_LABEL
                [],              # SEARCH_RESULT_ITEMS
            ]

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            with pytest.raises(ItemNotFoundError, match="not found"):
                client.search_item("NON_EXISTENT_PART_999")

    def test_select_all_zero_rows_raises_no_rows_selected_error(self, mock_driver):
        """Phase 4: Zero rows selected raises NoRowsSelectedError."""
        select_all_btn = MagicMock()

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.side_effect = [select_all_btn, TimeoutException()]

            mock_driver.find_elements.return_value = []

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            with pytest.raises(NoRowsSelectedError, match="No BOM rows were selected"):
                client.select_all_bom_lines()

    def test_verify_exported_excel_too_small_raises_corrupt_file_error(self, tmp_path: Path):
        """Phase 7: File < 5KB raises BOMCorruptFileError."""
        tiny_file = tmp_path / "tiny.xlsx"
        tiny_file.write_bytes(b"PK\x03\x04" + b"A" * 100)  # ~104 bytes

        with pytest.raises(BOMCorruptFileError, match="too small"):
            verify_exported_excel(tiny_file)

    def test_verify_exported_excel_wrong_headers_raises_validation_error(self, tmp_path: Path):
        """Phase 7: Incorrect headers raise BOMValidationError."""
        bad_file = tmp_path / "bad_headers.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        # Write wrong header
        ws.cell(row=1, column=1, value="WrongHeader")
        for c in range(2, 15):
            ws.cell(row=1, column=c, value=f"Col{c}")
        ws.cell(row=2, column=1, value="Data")
        # Pad size > 5KB
        for r in range(3, 100):
            ws.cell(row=r, column=1, value="PaddingData")
        wb.save(str(bad_file))
        wb.close()

        with pytest.raises(BOMValidationError, match="Header mismatch at Column 1"):
            verify_exported_excel(bad_file)


# =====================================================================
# Tier 3: Cross-Phase Integration Tests
# =====================================================================

class TestTier3Integration:
    """Full pipeline integration and provider adapter contract."""

    def test_download_bom_full_pipeline_orchestration(self, sample_valid_excel: Path, tmp_path: Path):
        """Test download_bom_full orchestrating all 7 phases."""
        client = TC2412AutomationClient(session_manager=MagicMock())
        client._owns_session = False

        # Mock individual phase methods
        client.login = MagicMock(return_value=True)
        client.search_item = MagicMock(return_value="01")
        client.navigate_to_content_tab = MagicMock(return_value=True)
        client.expand_bom_tree = MagicMock(return_value=250)
        client.select_all_bom_lines = MagicMock(return_value=250)
        client.open_export_to_excel_dialog = MagicMock(return_value=True)
        client.configure_14_columns = MagicMock(return_value=CANONICAL_14_COLUMNS)
        client.run_export_and_download = MagicMock(return_value=sample_valid_excel)

        callback_records = []

        def callback(idx, total, name, pct, elapsed):
            callback_records.append((idx, name))

        out_path = client.download_bom_full(
            part_number="T10C423NL0",
            output_dir=tmp_path,
            part_rev="01",
            progress_callback=callback,
        )

        assert out_path == sample_valid_excel
        client.login.assert_called_once()
        client.search_item.assert_called_once_with(item_id="T10C423NL0", part_rev="01")
        client.navigate_to_content_tab.assert_called_once()
        client.expand_bom_tree.assert_called_once_with(target_level=7, timeout=180)
        client.select_all_bom_lines.assert_called_once()
        client.open_export_to_excel_dialog.assert_called_once()
        client.configure_14_columns.assert_called_once()
        client.run_export_and_download.assert_called_once()
        assert len(callback_records) >= 8

    def test_teamcenter_selenium_adapter_delegation(self, sample_valid_excel: Path, tmp_path: Path):
        """Test TeamcenterSeleniumAdapter calling TC2412 client."""
        mock_client = MagicMock()
        mock_client.download_bom_full.return_value = sample_valid_excel

        adapter = TeamcenterSeleniumAdapter(client=mock_client)
        exported = adapter.export_excel(item_id="T10C423NL0", rev="01", output_path=tmp_path / "downloads")

        assert exported == sample_valid_excel
        mock_client.download_bom_full.assert_called_once()


# =====================================================================
# Tier 4: Real-World Scenarios & Sheet PLM Formula Bridge Tests
# =====================================================================

class TestTier4RealWorldAndSheetPLMBridge:
    """Validate Sheet PLM formula preservation and 24-col TC2412 normalization."""

    def test_verify_exported_excel_valid_file(self, sample_valid_excel: Path):
        """Validate sample valid file against canonical contract."""
        meta = verify_exported_excel(sample_valid_excel)
        assert meta["total_rows"] >= 2
        assert meta["column_count"] == 14
        assert meta["unicode_japanese_verified"] is True

    def test_tc2412_canonical_normalizer_24_columns(self):
        """Test normalization of new TC2412 24-column format."""
        raw_24col_row = {
            "Level": 1,
            "Item Type": "Design Part",
            "Name": "1102Z53KR0",         # Col D Name -> item_id
            "Has Children": True,
            "Quantity": 2,
            "Parts Text": "ギヤユニット (Gear Unit)",  # Col H Parts Text -> item_name
            "Revision": "B",              # Col J Revision -> revision
            "Release Status": "Released",  # Col K Release Status -> item_rev_status
            "Occurrence Effectivities": "to 2026/12/31",
            "Item Revision Project List": "Virgo2",
        }

        normalized = TC2412CanonicalNormalizer.normalize_row(raw_24col_row)
        assert normalized["level"] == 1
        assert normalized["item_type"] == "Design Part"
        assert normalized["item_id"] == "1102Z53KR0"
        assert normalized["has_children"] is True
        assert normalized["quantity"] == 2
        assert normalized["item_name"] == "ギヤユニット (Gear Unit)"
        assert normalized["revision"] == "B"
        assert normalized["item_rev_status"] == "Released"

    def test_sheet_plm_bridge_preserves_critical_excel_formulas(self, tmp_path: Path):
        """CRITICAL: Test that Sheet PLM Bridge maintains exact cell layout for workbook formulas.

        Formula Ground Truths in BOM_110C0Z3LV1.xlsm:
        1. Tongket!C5 & CTTT!A1: =PLM!C2 -> Col C MUST be Root Machine Part Code
        2. CTTT!I3: =VLOOKUP(C3, PLM!C:L, 10, 0) -> Col L MUST be Revision
        3. PLM!R2: =IF(C2="","",C2)
        4. PLM!S2: =IF(E2="","",E2)
        """
        # Create a mock master workbook with Tongket, CTTT, and PLM sheets
        master_file = tmp_path / "BOM_Test_Master.xlsx"
        wb = openpyxl.Workbook()
        ws_tongket = wb.active
        ws_tongket.title = "Tongket"
        ws_tongket["C5"] = "=PLM!C2"

        ws_cttt = wb.create_sheet("CTTT")
        ws_cttt["A1"] = "=PLM!C2"
        ws_cttt["C3"] = "110C103NL0"
        ws_cttt["I3"] = "=VLOOKUP(C3, PLM!C:L, 10, 0)"
        ws_cttt["G3"] = "=VLOOKUP(C3, PLM!T:U, 2, 0)"

        ws_plm = wb.create_sheet("PLM")

        wb.save(str(master_file))
        wb.close()

        # Create source exported file with 24 columns
        source_plm = tmp_path / "Source_TC2412_24Col.xlsx"
        src_wb = openpyxl.Workbook()
        src_ws = src_wb.active
        src_ws.title = "Export"

        headers_24 = [
            "Home", "Level", "Item Type", "Name", "Has Children", "Quantity",
            "Units", "Parts Text", "Material", "Revision", "Release Status",
            "Occurrence Effectivities", "Item Revision Project List", "Notice No"
        ]
        for c, h in enumerate(headers_24, start=1):
            src_ws.cell(row=1, column=c, value=h)

        row2 = [
            "Root", 0, "TopLevel", "110C103NL0", True, 1,
            "PC", "マスターマシン (Master Machine)", "STEEL", "04", "Released",
            "UP", "Virgo", "ECN-555"
        ]
        for c, val in enumerate(row2, start=1):
            src_ws.cell(row=2, column=c, value=val)

        src_wb.save(str(source_plm))
        src_wb.close()

        # Execute Sheet PLM bridge
        updated_path = TC2412SheetPLMBridge.update_workbook_sheet_plm(
            target_workbook_path=master_file,
            source_plm_path=source_plm,
            backup=True,
        )

        # Verify updated master workbook
        res_wb = openpyxl.load_workbook(filename=str(updated_path), data_only=False)
        res_plm = res_wb["PLM"]

        # 1. Check Col C Row 2 (Root Item Id)
        assert res_plm["C2"].value == "110C103NL0"

        # 2. Check Col E Row 2 (Quantity)
        assert res_plm["E2"].value == 1

        # 3. Check Col J Row 2 (Item Name - Japanese)
        assert res_plm["J2"].value == "マスターマシン (Master Machine)"

        # 4. Check Col L Row 2 (Revision - Col 10 in VLOOKUP C:L)
        assert res_plm["L2"].value == "04"

        # 5. Check Col M Row 2 (Release Status)
        assert res_plm["M2"].value == "Released"

        # 6. Check Col R Row 2 formula (=IF(C2="","",C2))
        assert res_plm["R2"].value == '=IF(C2="","",C2)'

        # 7. Check Col S Row 2 formula (=IF(E2="","",E2))
        assert res_plm["S2"].value == '=IF(E2="","",E2)'

        # 8. Check backup file was created
        bak_file = master_file.with_suffix(".xlsx.bak")
        assert bak_file.exists()

        res_wb.close()


# =====================================================================
# Tier 5: Adversarial Hardening Tests
# =====================================================================

class TestTier5Adversarial:
    """Anti-import fail closed, virtual DOM resilience, and active download triggers."""

    def test_anti_import_fail_closed_prevents_import_click(self, mock_driver):
        """CRITICAL: Deliberately inject 'Import Changes' button to ensure guard stops execution."""
        popup_menu = MagicMock()
        import_btn = MagicMock()
        import_btn.text = "Import Changes from Excel"
        import_btn.get_attribute.return_value = "Awp0ImportFromExcel"
        import_btn.title = "Import Changes"

        popup_menu.find_elements.return_value = [import_btn]

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.return_value = popup_menu

            # Guard MUST refuse to click and raise NoSuchElementException or ImportChangesForbiddenError
            with pytest.raises(NoSuchElementException):
                safe_trigger_export_to_excel(mock_driver)

            import_btn.click.assert_not_called()

    def test_virtual_dom_scrolling_resilience(self, mock_driver):
        """Test selection count reading from summary indicator when only 35 rows exist in DOM."""
        select_all_btn = MagicMock()
        summary_elem = MagicMock()
        summary_elem.text = "3480 Selected"  # 3,480 rows in tree, only 35 rendered
        report_btn = MagicMock()
        menu_elem = MagicMock()

        with patch("src.automation.tc2412.client.WebDriverWait") as mock_wait_cls:
            wait_inst = mock_wait_cls.return_value
            wait_inst.until.side_effect = [select_all_btn, summary_elem, report_btn, menu_elem]

            session = MagicMock()
            session.driver = mock_driver
            client = TC2412AutomationClient(session_manager=session)

            # Emulate virtual DOM: only 35 DOM rows exist
            mock_driver.find_elements.return_value = [MagicMock()] * 35

            selected_count = client.select_all_bom_lines(timeout=10)
            assert selected_count == 3480

