"""Feature F18: SAP Status Bar Fail-Closed Guard Isolation Tests.

Verifies:
1. Detecting Status Bar message type 'E' (Error) and halting execution immediately.
2. Detecting Status Bar message type 'A' (Abort) and halting execution.
3. Resetting navigation to top-level menu (/n) upon encountering an error.
4. Allowing export to proceed when status bar reports 'S' (Success) or 'W' (Warning).
5. Capturing verbatim SAP status bar error messages into ExportResult.
"""

import datetime
from pathlib import Path
from unittest.mock import MagicMock
import pytest

from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import CS12Params


class TestF18SAPFailClosedGuard:
    """Test suite for Feature F18: SAP Status Bar Fail-Closed Guard."""

    def test_f18_status_bar_error_type_e_aborts(self, tmp_path: Path, mock_sap_session):
        """Test 1: Status bar message type 'E' halts export and returns failure."""
        sbar_mock = MagicMock()
        sbar_mock.MessageType = "E"
        sbar_mock.Text = "Material 110C999NL0 not maintained in plant 2200"

        mock_sap_session.findById.side_effect = lambda eid: sbar_mock if "sbar" in eid else MagicMock()
        service = CS12Service(mock_sap_session)

        params = CS12Params(
            material="110C999NL0",
            valid_date=datetime.date(2026, 9, 17),
            destination_dir=tmp_path,
        )

        res = service.execute_cs12_and_export(params)
        assert res.success is False
        assert res.status_code == "E"
        assert "Material 110C999NL0 not maintained in plant 2200" in res.error_message

    def test_f18_status_bar_abort_type_a(self, tmp_path: Path, mock_sap_session):
        """Test 2: Status bar message type 'A' halts export."""
        sbar_mock = MagicMock()
        sbar_mock.MessageType = "A"
        sbar_mock.Text = "Critical database lock error"

        mock_sap_session.findById.side_effect = lambda eid: sbar_mock if "sbar" in eid else MagicMock()
        service = CS12Service(mock_sap_session)

        params = CS12Params(material="110C999NL0", destination_dir=tmp_path)
        res = service.execute_cs12_and_export(params)
        assert res.success is False
        assert res.status_code == "A"

    def test_f18_error_resets_navigation_to_home(self, tmp_path: Path, mock_sap_session):
        """Test 3: Reset navigation (/n) is called when an error is caught."""
        sbar_mock = MagicMock()
        sbar_mock.MessageType = "E"
        sbar_mock.Text = "BOM not found"

        okcd_mock = MagicMock()

        def find_elem(eid):
            if "sbar" in eid:
                return sbar_mock
            if "okcd" in eid:
                return okcd_mock
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem
        service = CS12Service(mock_sap_session)

        params = CS12Params(material="110C999NL0", destination_dir=tmp_path)
        service.execute_cs12_and_export(params)

        # /n should be sent to okcd
        assert okcd_mock.Text == "/n"

    def test_f18_warning_and_success_continue(self, tmp_path: Path, mock_sap_session):
        """Test 4: Warning (W) or Success (S) messages allow export execution to continue."""
        sbar_mock = MagicMock()
        sbar_mock.MessageType = "W"
        sbar_mock.Text = "Alternative 01 selected by default"

        dest_file = tmp_path / "R3_MAT_17_09_2026.xls"
        dest_file.write_text("DUMMY DATA")

        def find_elem(eid):
            if "sbar" in eid:
                return sbar_mock
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem
        service = CS12Service(mock_sap_session)

        params = CS12Params(
            material="MAT",
            valid_date=datetime.date(2026, 9, 17),
            destination_dir=tmp_path,
        )

        res = service.execute_cs12_and_export(params)
        # Should not fail early at status bar gate
        assert res.status_code != "E"

    def test_f18_structured_export_result_captures_error_message(self, tmp_path: Path, mock_sap_session):
        """Test 5: ExportResult accurately preserves error details and timing."""
        sbar_mock = MagicMock()
        sbar_mock.MessageType = "E"
        sbar_mock.Text = "No authorization for transaction CS12"

        mock_sap_session.findById.side_effect = lambda eid: sbar_mock if "sbar" in eid else MagicMock()
        service = CS12Service(mock_sap_session)

        params = CS12Params(material="MAT", destination_dir=tmp_path)
        res = service.execute_cs12_and_export(params)

        d = res.to_dict()
        assert d["success"] is False
        assert "No authorization" in d["error_message"]
        assert d["execution_time_sec"] >= 0
