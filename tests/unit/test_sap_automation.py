"""
Comprehensive Unit & Mock Test Suite for SAP R3 CS12 Automation Module.
Verifies models, connection manager, CS12 execution service, fail-closed guards,
resilient parsing across diverse layouts, and application layer contracts.
"""

import datetime
from pathlib import Path
from typing import Any, Dict, Optional
from unittest.mock import MagicMock, call, patch

import pandas as pd
import pytest

from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.cs12 import CS12Service, SAPCS12Client
from src.automation.sap.models import (
    CS12Params,
    ExportResult,
    R3ComponentRow,
    SAPConnectionError,
    SAPCS12Error,
    SAPCredentials,
    SAPError,
    SAPParseError,
)
from src.automation.sap.parser import ResilientR3Parser, parse_r3_cs12_file


# ============================================================================
# Section 1: Data Models & Exception Hierarchy Tests
# ============================================================================

class TestSAPModels:
    """Tests for SAPCredentials, CS12Params, ExportResult, R3ComponentRow."""

    def test_sap_credentials_defaults_and_masking(self):
        creds = SAPCredentials()
        assert creds.username == "v130474"
        assert creds.password == "0123456789"
        assert creds.system == "P1J(ERP60-AWS)-VN"
        assert creds.language == "EN"

        d = creds.to_dict()
        assert d["password"] == "***"
        assert d["username"] == "v130474"

    def test_sap_credentials_custom(self):
        creds = SAPCredentials(
            username="custom_user",
            password="secret_password",
            system="CUSTOM_SYS",
            language="VN",
        )
        assert creds.username == "custom_user"
        assert creds.system == "CUSTOM_SYS"
        assert creds.to_dict()["password"] == "***"

    def test_cs12_params_date_and_filename_formatting(self):
        test_date = datetime.date(2026, 5, 12)
        params = CS12Params(
            material="110K123450",
            plant="2200",
            bom_usage="pp01",
            alternative="1",  # should be zfilled to 01
            valid_date=test_date,
            destination_dir=Path("D:/output/110K123450"),
        )
        assert params.alternative == "01"
        assert params.formatted_date_sap == "2026/05/12"
        assert params.formatted_date_filename == "12_05_2026"
        assert params.expected_filename == "R3_110K123450_12_05_2026.xls"
        assert params.target_file_path == Path("D:/output/110K123450/R3_110K123450_12_05_2026.xls")

    def test_export_result_serialization(self):
        now = datetime.datetime(2026, 9, 17, 10, 0, 0)
        res = ExportResult(
            success=True,
            material="T10K99999",
            file_path=Path("D:/test.xls"),
            status_code="S",
            execution_time_sec=2.3456,
            timestamp=now,
        )
        d = res.to_dict()
        assert d["success"] is True
        assert d["material"] == "T10K99999"
        assert "test.xls" in d["file_path"]
        assert d["execution_time_sec"] == 2.346
        assert d["status_code"] == "S"

    def test_r3_component_row(self):
        row = R3ComponentRow(
            part_code="  302K12345  ",
            quantity=" 4.5 ",
            rev_r3="A",
            item_num="0010",
            description="COVER FRONT",
            level=1,
        )
        assert row.part_code == "302K12345"
        assert row.quantity == 4.5
        assert row.rev_r3 == "A"
        assert row.item_num == "0010"
        assert row.level == 1
        d = row.to_dict()
        assert d["quantity"] == 4.5
        assert d["description"] == "COVER FRONT"

    def test_r3_component_row_revision_coercion(self):
        """Verify rev_r3 coerces integer or float revision numbers (e.g. 1, 2, 1.0) cleanly to string ("01", "02")."""
        # Integers
        r_int1 = R3ComponentRow(part_code="P1", quantity=1.0, rev_r3=1)
        assert r_int1.rev_r3 == "01"
        r_int2 = R3ComponentRow(part_code="P2", quantity=2.0, rev_r3=2)
        assert r_int2.rev_r3 == "02"
        r_int10 = R3ComponentRow(part_code="P10", quantity=1.0, rev_r3=10)
        assert r_int10.rev_r3 == "10"

        # Floats
        r_flt1 = R3ComponentRow(part_code="P1", quantity=1.0, rev_r3=1.0)
        assert r_flt1.rev_r3 == "01"
        r_flt2 = R3ComponentRow(part_code="P2", quantity=1.0, rev_r3=2.0)
        assert r_flt2.rev_r3 == "02"
        r_flt_nan = R3ComponentRow(part_code="PNAN", quantity=1.0, rev_r3=float("nan"))
        assert r_flt_nan.rev_r3 == ""

        # Strings
        r_str1 = R3ComponentRow(part_code="P1", quantity=1.0, rev_r3="1")
        assert r_str1.rev_r3 == "01"
        r_str01 = R3ComponentRow(part_code="P1", quantity=1.0, rev_r3="01")
        assert r_str01.rev_r3 == "01"
        r_str_none = R3ComponentRow(part_code="P1", quantity=1.0, rev_r3=None)
        assert r_str_none.rev_r3 == ""
        r_str_alpha = R3ComponentRow(part_code="P1", quantity=1.0, rev_r3="B")
        assert r_str_alpha.rev_r3 == "B"

    def test_exception_inheritance(self):
        assert issubclass(SAPConnectionError, SAPError)
        assert issubclass(SAPCS12Error, SAPError)
        assert issubclass(SAPParseError, SAPError)


# ============================================================================
# Section 2: Connection Manager Unit & Mock Tests
# ============================================================================

class MockSAPElement:
    """Helper to mock SAP GUI COM elements with findById support."""

    def __init__(self, name: str = "", text: str = "", message_type: str = ""):
        self.Name = name
        self.Text = text
        self.MessageType = message_type
        self.Busy = False
        self._children = {}
        self.selected = False
        self.pressed = False
        self.focused = False

    def findById(self, path: str) -> Any:
        if path in self._children:
            return self._children[path]
        raise KeyError(f"Element not found: {path}")

    def add_child(self, path: str, element: "MockSAPElement"):
        self._children[path] = element
        return element

    def Select(self):
        self.selected = True

    def SetFocus(self):
        self.focused = True

    def press(self):
        self.pressed = True

    def sendVKey(self, key_code: int):
        pass


class TestSAPConnectionManager:
    """Tests for SAP Logon discovery, ROT waiting, multi-logon, and authentication."""

    def test_is_saplogon_process_running_custom_checker(self):
        manager = SAPConnectionManager(process_checker=lambda: True)
        assert manager.is_saplogon_process_running() is True

        manager_false = SAPConnectionManager(process_checker=lambda: False)
        assert manager_false.is_saplogon_process_running() is False

    def test_launch_saplogon_missing_executable(self):
        creds = SAPCredentials(saplogon_path="Z:/non/existent/saplogon.exe")
        manager = SAPConnectionManager(credentials=creds, process_checker=lambda: False)
        with pytest.raises(SAPConnectionError, match="SAP Logon executable not found"):
            manager.launch_saplogon()

    def test_launch_saplogon_custom_launcher(self):
        launched = []
        manager = SAPConnectionManager(
            process_checker=lambda: False,
            process_launcher=lambda path: launched.append(path),
        )
        manager.launch_saplogon()
        assert len(launched) == 1
        assert "saplogon.exe" in launched[0]

    def test_ensure_saplogon_running_rot_timeout(self):
        mock_com = MagicMock()
        mock_com.GetObject.side_effect = Exception("ROT object not found")

        manager = SAPConnectionManager(
            com_provider=mock_com,
            process_checker=lambda: True,
        )
        with pytest.raises(SAPConnectionError, match=r"Timed out.*SAPGUI.*ROT"):
            manager.ensure_saplogon_running(timeout_sec=0.5)

    def test_get_or_create_session_reuses_existing_connection(self):
        mock_app = MagicMock()
        mock_conn = MagicMock()
        mock_conn.Description = "Production P1J(ERP60-AWS)-VN"
        mock_conn.Children.Count = 1

        mock_session = MockSAPElement("Session0")
        # Add okcd so is_logged_in returns True
        mock_session.add_child("wnd[0]/tbar[0]/okcd", MockSAPElement("okcd"))
        mock_conn.Children.return_value = mock_session
        mock_conn.Children.side_effect = lambda idx: mock_session

        mock_app.Connections.Count = 1
        mock_app.Connections.return_value = mock_conn
        mock_app.Connections.side_effect = lambda idx: mock_conn

        mock_com = MagicMock()
        mock_gui_auto = MagicMock()
        mock_gui_auto.GetScriptingEngine.return_value = mock_app
        mock_com.GetObject.return_value = mock_gui_auto

        manager = SAPConnectionManager(
            com_provider=mock_com,
            process_checker=lambda: True,
        )
        sess = manager.get_or_create_session(timeout_sec=1.0)
        assert sess is mock_session
        assert manager.is_logged_in(sess) is True
        mock_app.OpenConnection.assert_not_called()

    def test_get_or_create_session_opens_new_and_handles_multi_logon(self):
        mock_app = MagicMock()
        mock_app.Connections.Count = 0

        mock_new_conn = MagicMock()
        mock_session = MockSAPElement("Session0")
        mock_new_conn.Children.Count = 1
        mock_new_conn.Children.side_effect = lambda idx: mock_session
        mock_app.OpenConnection.return_value = mock_new_conn

        mock_com = MagicMock()
        mock_gui_auto = MagicMock()
        mock_gui_auto.GetScriptingEngine.return_value = mock_app
        mock_com.GetObject.return_value = mock_gui_auto

        # Setup multi-logon dialog and login fields
        rad_opt2 = MockSAPElement("radMULTI_LOGON_OPT2")
        btn_dialog_ok = MockSAPElement("btn0")
        bname = MockSAPElement("bname")
        bcode = MockSAPElement("bcode")
        langu = MockSAPElement("langu")
        wnd0 = MockSAPElement("wnd0")
        okcd = MockSAPElement("okcd")

        mock_session.add_child("wnd[1]/usr/radMULTI_LOGON_OPT2", rad_opt2)
        mock_session.add_child("wnd[1]/tbar[0]/btn[0]", btn_dialog_ok)
        mock_session.add_child("wnd[0]/usr/txtRSYST-BNAME", bname)
        mock_session.add_child("wnd[0]/usr/pwdRSYST-BCODE", bcode)
        mock_session.add_child("wnd[0]/usr/txtRSYST-LANGU", langu)
        mock_session.add_child("wnd[0]", wnd0)

        # Once logged in, username field disappears and okcd appears
        def send_vkey(key):
            # remove login fields and add okcd
            mock_session._children.pop("wnd[0]/usr/txtRSYST-BNAME", None)
            mock_session.add_child("wnd[0]/tbar[0]/okcd", okcd)
        wnd0.sendVKey = send_vkey

        manager = SAPConnectionManager(
            com_provider=mock_com,
            process_checker=lambda: True,
        )
        sess = manager.get_or_create_session(timeout_sec=2.0)

        assert sess is mock_session
        assert rad_opt2.selected is True
        assert btn_dialog_ok.pressed is True
        assert bname.Text == "v130474"
        assert bcode.Text == "0123456789"
        assert langu.Text == "EN"
        assert manager.is_logged_in(sess) is True


# ============================================================================
# Section 3: CS12 Service & Fail-Closed Guard Tests
# ============================================================================

class TestCS12Service:
    """Tests for /nCS12 parameter binding, status bar error checks, and export dialogs."""

    def test_fail_closed_status_bar_error(self, tmp_path):
        """When SAP status bar displays MessageType='E', abort export and reset navigation."""
        mock_session = MockSAPElement("Session0")

        okcd = MockSAPElement("okcd")
        wnd0 = MockSAPElement("wnd0")
        matnr = MockSAPElement("matnr")
        werks = MockSAPElement("werks")
        stlal = MockSAPElement("stlal")
        capid = MockSAPElement("capid")
        datuv = MockSAPElement("datuv")
        btn8 = MockSAPElement("btn8")
        btn45 = MockSAPElement("btn45")
        sbar = MockSAPElement("sbar", text="BOM not found in plant 2200", message_type="E")

        mock_session.add_child("wnd[0]/tbar[0]/okcd", okcd)
        mock_session.add_child("wnd[0]", wnd0)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-MATNR", matnr)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-WERKS", werks)
        mock_session.add_child("wnd[0]/usr/txtRC29L-STLAL", stlal)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-CAPID", capid)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-DATUV", datuv)
        mock_session.add_child("wnd[0]/tbar[1]/btn[8]", btn8)
        mock_session.add_child("wnd[0]/tbar[1]/btn[45]", btn45)
        mock_session.add_child("wnd[0]/sbar", sbar)

        service = CS12Service(session=mock_session)
        params = CS12Params(
            material="INVALID_MATNR",
            valid_date=datetime.date(2026, 9, 17),
            destination_dir=tmp_path,
        )

        result = service.execute_cs12_and_export(params)

        assert result.success is False
        assert result.status_code == "E"
        assert "BOM not found" in result.error_message
        assert okcd.Text == "/n"  # Navigation reset triggered
        assert btn45.pressed is False  # btn[45] export NOT triggered!

    def test_successful_cs12_export(self, tmp_path):
        """Simulate successful CS12 execution and spreadsheet export."""
        mock_session = MockSAPElement("Session0")

        okcd = MockSAPElement("okcd")
        wnd0 = MockSAPElement("wnd0")
        matnr = MockSAPElement("matnr")
        werks = MockSAPElement("werks")
        stlal = MockSAPElement("stlal")
        capid = MockSAPElement("capid")
        datuv = MockSAPElement("datuv")
        btn8 = MockSAPElement("btn8")
        btn45 = MockSAPElement("btn45")
        btn3 = MockSAPElement("btn3")
        sbar = MockSAPElement("sbar", text="Data selected", message_type="S")

        # Dialog controls
        radio_format = MockSAPElement("radio_spreadsheet")
        btn_ok_wnd1 = MockSAPElement("btn0")
        path_field = MockSAPElement("path_field")
        file_field = MockSAPElement("file_field")
        btn_save_wnd1 = MockSAPElement("btn0")

        mock_session.add_child("wnd[0]/tbar[0]/okcd", okcd)
        mock_session.add_child("wnd[0]", wnd0)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-MATNR", matnr)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-WERKS", werks)
        mock_session.add_child("wnd[0]/usr/txtRC29L-STLAL", stlal)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-CAPID", capid)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-DATUV", datuv)
        mock_session.add_child("wnd[0]/tbar[1]/btn[8]", btn8)
        mock_session.add_child("wnd[0]/tbar[1]/btn[45]", btn45)
        mock_session.add_child("wnd[0]/tbar[0]/btn[3]", btn3)
        mock_session.add_child("wnd[0]/sbar", sbar)

        mock_session.add_child(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]",
            radio_format,
        )
        mock_session.add_child("wnd[1]/tbar[0]/btn[0]", btn_ok_wnd1)
        mock_session.add_child("wnd[1]/usr/ctxtDY_PATH", path_field)
        mock_session.add_child("wnd[1]/usr/ctxtDY_FILENAME", file_field)

        service = CS12Service(session=mock_session)
        params = CS12Params(
            material="110K123450",
            valid_date=datetime.date(2026, 5, 12),
            destination_dir=tmp_path,
        )

        # Simulate file created by SAP export
        expected_file = tmp_path / params.expected_filename
        expected_file.write_text("dummy export content", encoding="utf-8")

        result = service.execute_cs12_and_export(params)

        assert result.success is True
        assert result.status_code == "S"
        assert result.file_path == expected_file
        assert matnr.Text == "110K123450"
        assert werks.Text == "2200"
        assert stlal.Text == "01"
        assert capid.Text == "pp01"
        assert datuv.Text == "2026/05/12"
        assert btn8.pressed is True
        assert btn45.pressed is True
        assert radio_format.selected is True
        assert path_field.Text.endswith("\\")  # Trailing slash verified!
        assert file_field.Text == "R3_110K123450_12_05_2026.xls"
        assert btn3.pressed is True

    def test_application_layer_contract_download_bom(self, tmp_path):
        """Verify compliance with PROJECT.md download_bom contract."""
        mock_session = MockSAPElement("Session0")
        service = SAPCS12Client(session=mock_session)

        # Mock execute_cs12_and_export returning success
        mock_file = tmp_path / "R3_110K123450_12_05_2026.xls"
        mock_file.write_text("test", encoding="utf-8")

        with patch.object(service, "execute_cs12_and_export") as mock_exec:
            mock_exec.return_value = ExportResult(
                success=True,
                material="110K123450",
                file_path=mock_file,
            )
            returned_path = service.download_bom(
                material_code="110K123450",
                valid_date=datetime.date(2026, 5, 12),
                output_dir=tmp_path,
            )
            assert returned_path == mock_file

    def test_download_bom_raises_on_failure(self, tmp_path):
        mock_session = MockSAPElement("Session0")
        service = CS12Service(session=mock_session)

        with patch.object(service, "execute_cs12_and_export") as mock_exec:
            mock_exec.return_value = ExportResult(
                success=False,
                material="ERROR_MAT",
                error_message="BOM error",
            )
            with pytest.raises(SAPCS12Error, match="Failed to download BOM"):
                service.download_bom("ERROR_MAT", datetime.date.today(), tmp_path)

    def test_batch_download_and_archiving(self, tmp_path):
        """Verifies unified processing across 110* and T10* materials with archiving."""
        mock_session = MockSAPElement("Session0")
        service = CS12Service(session=mock_session)

        materials = ["110K123450", "T10K67890"]
        test_date = datetime.date(2026, 5, 12)

        # Pre-create old file to test archiving
        old_dir = tmp_path / "110K123450"
        old_dir.mkdir(parents=True)
        old_file = old_dir / "R3_110K123450_01_01_2026.xls"
        old_file.write_text("old version", encoding="utf-8")

        def mock_exec(params):
            out_file = params.destination_dir / params.expected_filename
            out_file.write_text("new content", encoding="utf-8")
            return ExportResult(
                success=True,
                material=params.material,
                file_path=out_file,
            )

        with patch.object(service, "execute_cs12_and_export", side_effect=mock_exec):
            results = service.batch_download(
                materials=materials,
                valid_date=test_date,
                base_destination_dir=tmp_path,
                archive_existing=True,
            )

            assert len(results) == 2
            assert all(r.success for r in results)

            # Check that old file was moved to capnhat/old/
            archive_dir = old_dir / "capnhat" / "old"
            assert archive_dir.exists()
            archived = list(archive_dir.glob("R3_110K123450_01_01_2026_*.xls"))
            assert len(archived) == 1


# ============================================================================
# Section 4: Resilient R3 Parser Unit Tests
# ============================================================================

class TestResilientR3Parser:
    """Tests for dynamic header matching, format decoding, layout variations, and cleaning."""

    @pytest.fixture
    def parser(self):
        return ResilientR3Parser()

    def test_parse_html_disguised_xls_with_revlev(self, parser):
        """Standard SAPLSPO5 HTML format with RevLev column present."""
        html_content = """
        <html>
        <body>
        <table>
            <tr><td colspan="5">Multilevel BOM Display for 110K123450</td></tr>
            <tr><td colspan="5">Plant: 2200, Valid: 2026/05/12</td></tr>
            <tr>
                <th>Level</th>
                <th>Item</th>
                <th>Component</th>
                <th>RevLev</th>
                <th>Quantity</th>
                <th>Description</th>
            </tr>
            <tr>
                <td>.1</td>
                <td>0010</td>
                <td>302K123450</td>
                <td>01</td>
                <td>2.000</td>
                <td>GEAR DRIVE</td>
            </tr>
            <tr>
                <td>..2</td>
                <td>0020</td>
                <td>302K678900</td>
                <td>A</td>
                <td>1.5</td>
                <td>SCREW M3</td>
            </tr>
        </table>
        </body>
        </html>
        """
        df = parser.parse_content(html_content)

        assert len(df) == 2
        assert list(df["part_code"]) == ["302K123450", "302K678900"]
        assert list(df["quantity"]) == [2.0, 1.5]
        assert list(df["rev_r3"]) == ["01", "A"]
        assert list(df["level"]) == [1, 2]
        assert list(df["item_num"]) == ["0010", "0020"]
        assert list(df["description"]) == ["GEAR DRIVE", "SCREW M3"]

    def test_parse_html_without_revlev_column(self, parser):
        """Layout variant WITHOUT RevLev column (proves no fragile column deletion needed)."""
        html_content = """
        <table>
            <tr>
                <th>Level</th>
                <th>PART CODE</th>
                <th>Q.TY</th>
                <th>DESCRIPTION</th>
            </tr>
            <tr>
                <td>1</td>
                <td>202K99999</td>
                <td>4.0</td>
                <td>BRACKET</td>
            </tr>
        </table>
        """
        df = parser.parse_content(html_content)

        assert len(df) == 1
        assert df.iloc[0]["part_code"] == "202K99999"
        assert df.iloc[0]["quantity"] == 4.0
        assert df.iloc[0]["rev_r3"] == ""

    def test_parse_tsv_with_vietnamese_headers(self, parser):
        """TSV format with Vietnamese headers (Mã linh kiện, Số lượng, Rev R3)."""
        tsv_content = (
            "Báo cáo BOM R3\n"
            "Ngày xuất: 17/09/2026\n"
            "Mã linh kiện\tSố lượng\tRev R3\tTên linh kiện\n"
            "110K12345\t10.0\t02\tBỘ TRỤC CHÍNH\n"
            "302K98765\t5.5\tB\tVÒNG ĐỆM\n"
        )
        df = parser.parse_content(tsv_content)

        assert len(df) == 2
        assert list(df["part_code"]) == ["110K12345", "302K98765"]
        assert list(df["quantity"]) == [10.0, 5.5]
        assert list(df["rev_r3"]) == ["02", "B"]

    def test_parse_quantity_formatting_nuances(self, parser):
        """Test thousand separators, comma decimals, and SAP trailing minus."""
        tsv_content = (
            "Component\tQuantity\tRevLev\n"
            "PART_A\t1,000.00\t00\n"
            "PART_B\t1.250,50\t00\n"
            "PART_C\t2,5\t00\n"
            "PART_D\t10-\t00\n"
            "PART_E\tinvalid_num\t00\n"
        )
        df = parser.parse_content(tsv_content)

        assert df.iloc[0]["quantity"] == 1000.0
        assert df.iloc[1]["quantity"] == 1250.5
        assert df.iloc[2]["quantity"] == 2.5
        assert df.iloc[3]["quantity"] == -10.0
        assert df.iloc[4]["quantity"] == 0.0

    def test_parse_filtering_separators_and_noise(self, parser):
        """Verify headers, dashed lines, and blank rows are discarded."""
        tsv_content = (
            "Component\tQuantity\tRevLev\n"
            "--------------------\t----\t------\n"
            "302K11111\t1.0\t01\n"
            "Component\tQuantity\tRevLev\n"
            "====================\t====\t======\n"
            "302K22222\t2.0\t02\n"
            "\t\t\n"
        )
        df = parser.parse_content(tsv_content)

        assert len(df) == 2
        assert list(df["part_code"]) == ["302K11111", "302K22222"]

    def test_parse_to_records_typed_output(self, tmp_path, parser):
        """Verify parse_to_records returns validated List[R3ComponentRow]."""
        test_file = tmp_path / "test_r3.xls"
        test_file.write_text(
            "<table><tr><th>Component</th><th>Quantity</th><th>RevLev</th></tr>"
            "<tr><td>PART100</td><td>3.0</td><td>01</td></tr></table>",
            encoding="utf-8",
        )

        records = parser.parse_to_records(test_file)
        assert len(records) == 1
        rec = records[0]
        assert isinstance(rec, R3ComponentRow)
        assert rec.part_code == "PART100"
        assert rec.quantity == 3.0
        assert rec.rev_r3 == "01"

    def test_convenience_function_parse_r3_cs12_file(self, tmp_path):
        test_file = tmp_path / "test_convenience.xls"
        test_file.write_text(
            "<table><tr><th>Component</th><th>Quantity</th><th>RevLev</th></tr>"
            "<tr><td>PART200</td><td>5.0</td><td>A</td></tr></table>",
            encoding="utf-8",
        )
        df = parse_r3_cs12_file(test_file)
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 1
        assert df.iloc[0]["part_code"] == "PART200"

    def test_empty_and_missing_file_errors(self, tmp_path, parser):
        empty_file = tmp_path / "empty.xls"
        empty_file.touch()

        with pytest.raises(SAPParseError, match="empty"):
            parser.parse(empty_file)

        with pytest.raises(SAPParseError, match="does not exist"):
            parser.parse(tmp_path / "does_not_exist.xls")

    def test_parse_fallback_heuristics_when_no_headers_found(self, parser):
        """When text contains data with no recognized header names, apply fallback columns."""
        raw_tsv = (
            "Title line\n"
            "Subtitle\n"
            "0010\tMAT001\t10\t2.0\tA\tDESC1\n"
            "0020\tMAT002\t20\t5.0\tB\tDESC2\n"
        )
        df = parser.parse_content(raw_tsv)
        assert len(df) >= 2

    def test_parse_excel_fallback(self, tmp_path, parser):
        """Verify parsing when file is binary excel."""
        df_dummy = pd.DataFrame({
            "Component": ["PART_EXCEL_1", "PART_EXCEL_2"],
            "Quantity": [10.5, 20.0],
            "RevLev": ["01", "02"],
        })
        excel_path = tmp_path / "sample.xlsx"
        df_dummy.to_excel(excel_path, index=False)

        df_parsed = parser.parse(excel_path)
        assert len(df_parsed) == 2
        assert list(df_parsed["part_code"]) == ["PART_EXCEL_1", "PART_EXCEL_2"]
        assert list(df_parsed["quantity"]) == [10.5, 20.0]


class TestAdditionalEdgeCases:
    """Additional white-box and branch tests for connection manager and CS12 service."""

    def test_connection_process_iter_psutil(self):
        with patch("psutil.process_iter") as mock_iter:
            mock_p = MagicMock()
            mock_p.info = {"name": "saplogon.exe"}
            mock_iter.return_value = [mock_p]

            manager = SAPConnectionManager()
            assert manager.is_saplogon_process_running() is True

    def test_connection_disconnect_and_popups(self):
        mock_session = MockSAPElement("Session0")
        wnd1 = MockSAPElement("wnd1")
        btn0 = MockSAPElement("btn0")
        mock_session.add_child("wnd[1]", wnd1)
        mock_session.add_child("wnd[1]/tbar[0]/btn[0]", btn0)

        pressed = []
        def on_press():
            pressed.append(True)
            mock_session._children.pop("wnd[1]", None)
        btn0.press = on_press

        manager = SAPConnectionManager()
        manager.session = mock_session
        manager.dismiss_post_login_popups(mock_session)
        assert len(pressed) == 1

        manager.disconnect()
        assert manager.session is None
        assert manager.connection is None


    def test_login_failure_sbar_error(self):
        mock_session = MockSAPElement("Session0")
        sbar = MockSAPElement("sbar", text="User locked", message_type="E")
        bname = MockSAPElement("bname")
        wnd0 = MockSAPElement("wnd0")

        mock_session.add_child("wnd[0]/usr/txtRSYST-BNAME", bname)
        mock_session.add_child("wnd[0]/sbar", sbar)
        mock_session.add_child("wnd[0]", wnd0)

        manager = SAPConnectionManager()
        manager.session = mock_session
        with pytest.raises(SAPConnectionError, match="User locked"):
            manager.handle_multi_logon_and_login(mock_session)

    def test_cs12_overwrite_popup(self, tmp_path):
        mock_session = MockSAPElement("Session0")
        okcd = MockSAPElement("okcd")
        wnd0 = MockSAPElement("wnd0")
        matnr = MockSAPElement("matnr")
        werks = MockSAPElement("werks")
        stlal = MockSAPElement("stlal")
        capid = MockSAPElement("capid")
        datuv = MockSAPElement("datuv")
        btn8 = MockSAPElement("btn8")
        btn45 = MockSAPElement("btn45")
        btn3 = MockSAPElement("btn3")
        sbar = MockSAPElement("sbar", text="Success", message_type="S")
        radio_format = MockSAPElement("radio_spreadsheet")
        btn_ok_wnd1 = MockSAPElement("btn0")
        path_field = MockSAPElement("path_field")
        file_field = MockSAPElement("file_field")
        btn_replace = MockSAPElement("btn_replace")

        mock_session.add_child("wnd[0]/tbar[0]/okcd", okcd)
        mock_session.add_child("wnd[0]", wnd0)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-MATNR", matnr)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-WERKS", werks)
        mock_session.add_child("wnd[0]/usr/txtRC29L-STLAL", stlal)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-CAPID", capid)
        mock_session.add_child("wnd[0]/usr/ctxtRC29L-DATUV", datuv)
        mock_session.add_child("wnd[0]/tbar[1]/btn[8]", btn8)
        mock_session.add_child("wnd[0]/tbar[1]/btn[45]", btn45)
        mock_session.add_child("wnd[0]/tbar[0]/btn[3]", btn3)
        mock_session.add_child("wnd[0]/sbar", sbar)
        mock_session.add_child(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]",
            radio_format,
        )
        mock_session.add_child("wnd[1]/tbar[0]/btn[0]", btn_ok_wnd1)
        mock_session.add_child("wnd[1]/usr/ctxtDY_PATH", path_field)
        mock_session.add_child("wnd[1]/usr/ctxtDY_FILENAME", file_field)
        mock_session.add_child("wnd[2]/usr/btnSPOP-OPTION1", btn_replace)

        service = CS12Service(session=mock_session)
        params = CS12Params(
            material="110K99999",
            valid_date=datetime.date(2026, 5, 12),
            destination_dir=tmp_path,
        )
        out_file = tmp_path / params.expected_filename
        out_file.write_text("sample content", encoding="utf-8")

        res = service.execute_cs12_and_export(params)
        assert res.success is True
        assert btn_replace.pressed is True

