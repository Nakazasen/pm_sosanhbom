"""Feature F17: SAP CS12 Transaction Execution Isolation Tests.

Verifies:
1. CS12Params default values (plant=2200, usage=pp01, alt=01).
2. Unconditional navigation using /nCS12 to reset transaction state.
3. Parameter binding (MATNR, WERKS, STLAL, CAPID, DATUV) and execution (btn[8]).
4. Fulfilling the Application Layer Contract download_bom.
5. Standard YYYY/MM/DD date formatting for SAP input screens.
"""

import datetime
from pathlib import Path
from unittest.mock import MagicMock
import pytest

from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import CS12Params, ExportResult


class TestF17SAPCS12Execution:
    """Test suite for Feature F17: SAP CS12 Transaction Execution."""

    def test_f17_cs12_parameter_contract(self):
        """Test 1: Mandatory CS12 defaults match Kyocera production requirements."""
        test_date = datetime.date(2026, 9, 17)
        params = CS12Params(material="110C103NL0", valid_date=test_date)

        assert params.plant == "2200"
        assert params.bom_usage == "pp01"
        assert params.alternative == "01"
        assert params.formatted_date_sap == "2026/09/17"
        assert params.formatted_date_filename == "17_09_2026"
        assert params.expected_filename == "R3_110C103NL0_17_09_2026.xls"

    def test_f17_navigate_with_reset_prefix(self, mock_sap_session):
        """Test 2: Navigation to CS12 prepends '/n' to avoid getting trapped in sub-screens."""
        service = CS12Service(mock_sap_session)
        service.navigate_to_cs12()

        # okcd text should be set to /nCS12
        okcd_call = mock_sap_session.findById("wnd[0]/tbar[0]/okcd")
        assert okcd_call.Text == "/nCS12"

    def test_f17_field_binding_and_f8_execution(self, tmp_path: Path, mock_sap_session):
        """Test 3: Bind parameters to RC29L fields and press F8 (btn[8])."""
        fields = {}

        def find_elem(elem_id):
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            fields[elem_id] = elem
            return elem

        mock_sap_session.findById.side_effect = find_elem
        service = CS12Service(mock_sap_session)

        # Mock successful file write
        dest_dir = tmp_path / "110C103NL0"
        dest_dir.mkdir()
        exported_file = dest_dir / "R3_110C103NL0_17_09_2026.xls"
        exported_file.write_text("R3 TEST DATA")

        params = CS12Params(
            material="110C103NL0",
            valid_date=datetime.date(2026, 9, 17),
            destination_dir=dest_dir,
        )

        res = service.execute_cs12_and_export(params)
        assert res.success is True
        assert fields["wnd[0]/usr/ctxtRC29L-MATNR"].Text == "110C103NL0"
        assert fields["wnd[0]/usr/ctxtRC29L-WERKS"].Text == "2200"
        assert fields["wnd[0]/usr/txtRC29L-STLAL"].Text == "01"
        assert fields["wnd[0]/usr/ctxtRC29L-CAPID"].Text == "pp01"
        assert fields["wnd[0]/usr/ctxtRC29L-DATUV"].Text == "2026/09/17"

    def test_f17_download_bom_interface_contract(self, tmp_path: Path, mock_sap_session):
        """Test 4: Application Layer Contract download_bom returns Path."""
        service = CS12Service(mock_sap_session)

        test_dir = tmp_path / "out"
        test_dir.mkdir()
        mock_file = test_dir / "R3_110C103NL0_17_09_2026.xls"
        mock_file.write_text("DATA")

        service.execute_cs12_and_export = MagicMock(
            return_value=ExportResult(success=True, material="110C103NL0", file_path=mock_file)
        )

        res_path = service.download_bom("110C103NL0", datetime.date(2026, 9, 17), test_dir)
        assert res_path == mock_file
        assert res_path.exists()

    def test_f17_date_formatting_sap_standard(self):
        """Test 5: Leap year and standard dates format accurately."""
        # Leap year Feb 29
        leap_params = CS12Params(material="MAT", valid_date=datetime.date(2024, 2, 29))
        assert leap_params.formatted_date_sap == "2024/02/29"
        assert leap_params.formatted_date_filename == "29_02_2024"
