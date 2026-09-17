"""Feature F19: SAP Spreadsheet Export & Path Routing Isolation Tests.

Verifies:
1. SAPLSPO5:0150 format selection dialog interaction (radSPOPLI-SELFLAG[1,0]).
2. Trailing backslash path formatting required by SAP GUI ctxtDY_PATH.
3. Automated handling of the file overwrite confirmation popup (btnSPOP-OPTION1).
4. Returning safely to main transaction screen via btn[3] (F3).
5. Post-export verification of file presence and non-zero size on disk.
"""

import datetime
from pathlib import Path
from unittest.mock import MagicMock
import pytest

from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import CS12Params


class TestF19SAPExportRouting:
    """Test suite for Feature F19: SAP Spreadsheet Export & Path Routing."""

    def test_f19_saplspo5_dialog_selection(self, tmp_path: Path, mock_sap_session):
        """Test 1: Select spreadsheet radio button in SAPLSPO5:0150 dialog."""
        radio_mock = MagicMock()
        btn_ok = MagicMock()

        def find_elem(eid):
            if "radSPOPLI-SELFLAG" in eid:
                return radio_mock
            if "wnd[1]/tbar[0]/btn[0]" in eid:
                return btn_ok
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem
        service = CS12Service(mock_sap_session)

        # Pre-create file to satisfy check
        out_file = tmp_path / "R3_MAT_17_09_2026.xls"
        out_file.write_text("CONTENT")

        params = CS12Params(material="MAT", destination_dir=tmp_path, valid_date=datetime.date(2026, 9, 17))
        service.execute_cs12_and_export(params)

        radio_mock.Select.assert_called_once()
        btn_ok.press.assert_called()

    def test_f19_trailing_backslash_path_formatting(self, tmp_path: Path, mock_sap_session):
        """Test 2: Ensure destination directory passed to ctxtDY_PATH ends with trailing backslash."""
        path_mock = MagicMock()
        mock_sap_session.findById.side_effect = lambda eid: path_mock if "ctxtDY_PATH" in eid else MagicMock()

        out_file = tmp_path / "R3_MAT_17_09_2026.xls"
        out_file.write_text("CONTENT")

        service = CS12Service(mock_sap_session)
        params = CS12Params(material="MAT", destination_dir=tmp_path, valid_date=datetime.date(2026, 9, 17))
        service.execute_cs12_and_export(params)

        assert path_mock.Text.endswith("\\")

    def test_f19_overwrite_confirmation_handling(self, tmp_path: Path, mock_sap_session):
        """Test 3: Confirm file overwrite popup (wnd[2]/usr/btnSPOP-OPTION1) when file exists."""
        replace_btn = MagicMock()

        def find_elem(eid):
            if "btnSPOP-OPTION1" in eid:
                return replace_btn
            return MagicMock()

        mock_sap_session.Children.Count = 2
        mock_sap_session.findById.side_effect = find_elem
        service = CS12Service(mock_sap_session)

        out_file = tmp_path / "R3_MAT_17_09_2026.xls"
        out_file.write_text("EXISTING CONTENT")

        params = CS12Params(material="MAT", destination_dir=tmp_path, valid_date=datetime.date(2026, 9, 17))
        service.execute_cs12_and_export(params)

        replace_btn.press.assert_called_once()

    def test_f19_return_to_main_screen_via_f3(self, tmp_path: Path, mock_sap_session):
        """Test 4: Press F3 (btn[3]) after export to return to the initial screen."""
        btn_back = MagicMock()
        mock_sap_session.findById.side_effect = lambda eid: btn_back if "wnd[0]/tbar[0]/btn[3]" in eid else MagicMock()

        out_file = tmp_path / "R3_MAT_17_09_2026.xls"
        out_file.write_text("CONTENT")

        service = CS12Service(mock_sap_session)
        params = CS12Params(material="MAT", destination_dir=tmp_path, valid_date=datetime.date(2026, 9, 17))
        service.execute_cs12_and_export(params)

        btn_back.press.assert_called_once()

    def test_f19_file_verification_on_disk(self, tmp_path: Path, mock_sap_session):
        """Test 5: Fails if target file was not written to disk or has 0 bytes."""
        service = CS12Service(mock_sap_session)
        # Target file NOT created on disk
        params = CS12Params(material="MAT_NO_FILE", destination_dir=tmp_path, valid_date=datetime.date(2026, 9, 17))

        res = service.execute_cs12_and_export(params)
        assert res.success is False
        assert "not found or empty" in res.error_message
