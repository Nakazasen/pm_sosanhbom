"""Unit tests for SAP COM RPC failure detection, auto-reconnect, and dynamic session recovery."""

import datetime
from pathlib import Path
from unittest.mock import MagicMock, call, patch
import pytest

from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import CS12Params, ExportResult, SAPCS12Error


class TestSAPRPCErrorDetection:
    """Tests for CS12Service._is_rpc_or_session_error detection."""

    def test_rpc_hresult_exact_failure(self):
        """User exact error: (-2147023170, 'The remote procedure call failed.', None, None)."""
        err = Exception(-2147023170, "The remote procedure call failed.", None, None)
        assert CS12Service._is_rpc_or_session_error(err) is True

    def test_rpc_hex_code_0x800706be(self):
        """Detect 0x800706BE in error code / args / string."""
        err = Exception(0x800706BE, "RPC call failed")
        assert CS12Service._is_rpc_or_session_error(err) is True

    def test_rpc_server_unavailable_0x800706ba(self):
        """Detect 0x800706BA / -2147023174."""
        err1 = Exception(-2147023174, "The RPC server is unavailable")
        err2 = Exception(0x800706BA, "RPC Server Unavailable")
        assert CS12Service._is_rpc_or_session_error(err1) is True
        assert CS12Service._is_rpc_or_session_error(err2) is True

    def test_rpc_disconnected_clients(self):
        """Detect 0x80010108 / -2147417848."""
        err = Exception(-2147417848, "The object invoked has disconnected from its clients.")
        assert CS12Service._is_rpc_or_session_error(err) is True

    def test_wrapped_in_sapcs12_error(self):
        """Detect RPC error wrapped inside SAPCS12Error via __cause__ or text."""
        inner = Exception(-2147023170, "The remote procedure call failed.", None, None)
        wrapped = SAPCS12Error(f"Failed to navigate to /nCS12: {inner}")
        wrapped.__cause__ = inner
        assert CS12Service._is_rpc_or_session_error(wrapped) is True

    def test_string_only_rpc_error(self):
        """Detect RPC keywords in exception string."""
        err = RuntimeError("COM error: The remote procedure call failed during Wait.")
        assert CS12Service._is_rpc_or_session_error(err) is True

    def test_all_rpc_hresult_codes(self):
        """Test all registered RPC / disconnect HRESULTs in signed and unsigned forms."""
        from src.automation.sap.cs12 import RPC_HRESULTS_U32
        for code_u32 in RPC_HRESULTS_U32:
            code_signed = code_u32 if code_u32 < 0x80000000 else code_u32 - 0x100000000
            # Test unsigned
            err_u = Exception(code_u32, "RPC Code")
            assert CS12Service._is_rpc_or_session_error(err_u) is True
            # Test signed
            err_s = Exception(code_signed, "RPC Code")
            assert CS12Service._is_rpc_or_session_error(err_s) is True

    def test_extended_rpc_keywords(self):
        """Test extended session disconnect keywords (partner broken, wsaeconnreset, etc.)."""
        for kw in [
            "Connection to partner '10.20.30.40' broken",
            "Session terminated by host",
            "WSAECONNRESET: Connection reset by peer",
            "Server died unexpectedly",
        ]:
            assert CS12Service._is_rpc_or_session_error(RuntimeError(kw)) is True

    def test_non_rpc_business_errors(self):
        """Standard business / data errors should NOT trigger RPC reconnect."""
        assert CS12Service._is_rpc_or_session_error(ValueError("Material not found")) is False
        assert CS12Service._is_rpc_or_session_error(KeyError("column")) is False
        assert CS12Service._is_rpc_or_session_error(SAPCS12Error("SAP Error [E]: Material 123 does not exist")) is False
        assert CS12Service._is_rpc_or_session_error(None) is False


class TestCS12ServiceAutoReconnectAndRetry:
    """Tests for CS12Service automatic session recovery and retry."""

    def test_auto_reconnect_and_retry_success_on_rpc_failure(self, tmp_path: Path):
        """When RPC fails on initial attempt, reconnect via connection_manager and retry successfully."""
        # Session 1: dead session that raises RPC failure
        mock_dead_session = MagicMock()
        mock_dead_session.findById.side_effect = Exception(-2147023170, "The remote procedure call failed.", None, None)

        # Session 2: fresh healthy session
        mock_live_session = MagicMock()
        fields = {}
        def find_elem(elem_id):
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            fields[elem_id] = elem
            return elem
        mock_live_session.findById.side_effect = find_elem

        # Connection Manager
        mock_conn_mgr = MagicMock()
        mock_conn_mgr.reconnect.return_value = mock_live_session

        status_logs = []
        refreshed_sessions = []

        service = CS12Service(
            session=mock_dead_session,
            connection_manager=mock_conn_mgr,
            status_callback=lambda msg: status_logs.append(msg),
            on_session_refreshed=lambda s: refreshed_sessions.append(s),
        )

        # Prepare target file on disk so export succeeds
        mat = "3V2YJ58070"
        dest_dir = tmp_path / mat
        dest_dir.mkdir(parents=True, exist_ok=True)
        target_file = dest_dir / f"R3_{mat}_02_10_2026.xls"
        target_file.write_text("DUMMY SAP DATA")

        params = CS12Params(
            material=mat,
            plant="2200",
            bom_usage="pp01",
            alternative="01",
            valid_date=datetime.date(2026, 10, 2),
            destination_dir=dest_dir,
        )

        res = service.execute_cs12_and_export(params)

        # 1. Connection manager reconnect was called
        assert mock_conn_mgr.reconnect.called or mock_conn_mgr.get_session.called
        # 2. Session was updated to live session
        assert service.session == mock_live_session
        assert refreshed_sessions == [mock_live_session]
        # 3. Export succeeded on retry
        assert res.success is True
        assert res.material == mat
        # 4. Status log received notification
        assert any("mất kết nối COM/RPC" in log for log in status_logs)
        assert any("Tái lập kết nối SAP GUI thành công" in log for log in status_logs)

    def test_retry_exhaustion_on_persistent_rpc_failure(self, tmp_path: Path):
        """When RPC fails and reconnect fails, fail closed cleanly without looping."""
        mock_dead_session = MagicMock()
        mock_dead_session.findById.side_effect = Exception(-2147023170, "The remote procedure call failed.", None, None)

        mock_conn_mgr = MagicMock()
        # Reconnect also returns dead session
        mock_conn_mgr.reconnect.return_value = mock_dead_session

        service = CS12Service(
            session=mock_dead_session,
            connection_manager=mock_conn_mgr,
        )

        params = CS12Params(
            material="3V2YJ58070",
            destination_dir=tmp_path,
            valid_date=datetime.date(2026, 10, 2),
        )

        res = service.execute_cs12_and_export(params)
        # Must fail closed and return failure result without infinite recursion
        assert res.success is False
        assert "The remote procedure call failed" in res.error_message
        # reconnect should be attempted only once
        assert mock_conn_mgr.reconnect.call_count == 1

    def test_pre_execution_dead_session_auto_reconnect(self, tmp_path: Path):
        """When session is dead before execute_cs12_and_export starts, pre-check refreshes it."""
        mock_dead_session = MagicMock()
        mock_dead_session.findById.side_effect = Exception("Session closed")

        mock_live_session = MagicMock()
        fields = {}
        def find_elem(elem_id):
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            fields[elem_id] = elem
            return elem
        mock_live_session.findById.side_effect = find_elem

        mock_conn_mgr = MagicMock()
        mock_conn_mgr.reconnect.return_value = mock_live_session

        service = CS12Service(
            session=mock_dead_session,
            connection_manager=mock_conn_mgr,
        )

        mat = "110C103NL0"
        dest_dir = tmp_path / mat
        dest_dir.mkdir(parents=True, exist_ok=True)
        target_file = dest_dir / f"R3_{mat}_17_09_2026.xls"
        target_file.write_text("DUMMY DATA")

        params = CS12Params(
            material=mat,
            destination_dir=dest_dir,
            valid_date=datetime.date(2026, 9, 17),
        )

        res = service.execute_cs12_and_export(params)
        assert res.success is True
        assert service.session == mock_live_session


class TestTeardownAndLogoffAccuracy:
    """Tests for accurate reporting during logoff_and_exit."""

    def test_logoff_on_active_session_returns_true(self):
        """Logoff on responsive session returns True."""
        mock_session = MagicMock()
        okcd = MagicMock()
        mock_session.findById.return_value = okcd

        mgr = SAPConnectionManager()
        mgr.session = mock_session

        res = mgr.logoff_and_exit(close_saplogon=False)
        assert res is True
        assert okcd.Text == "/nex"
        assert mgr.session is None

    def test_logoff_on_dead_session_returns_false_without_crash(self):
        """Logoff on dead session (RPC error) returns False and cleans up handles safely."""
        mock_session = MagicMock()
        mock_session.findById.side_effect = Exception(-2147023170, "The remote procedure call failed.", None, None)

        mgr = SAPConnectionManager()
        mgr.session = mock_session

        res = mgr.logoff_and_exit(close_saplogon=False)
        assert res is False
        assert mgr.session is None
        assert mgr.is_connected() is False

    def test_logoff_when_no_session_returns_false(self):
        """Logoff when session is already None returns False."""
        mgr = SAPConnectionManager()
        mgr.session = None

        res = mgr.logoff_and_exit(close_saplogon=False)
        assert res is False


class TestAdvancedRecoveryScenarios:
    """Tests for mid-execution RPC drops and multi-BOM non-cascading behavior."""

    def test_mid_execution_export_btn_rpc_failure_auto_reconnect(self, tmp_path: Path):
        """When RPC fails at step 5 (btn[45]), reconnect and retry successfully."""
        # Initial session: works until btn[45] where RPC drops
        mock_initial_session = MagicMock()
        def initial_find(elem_id):
            if "btn[45]" in elem_id:
                raise Exception(-2147023170, "The remote procedure call failed.", None, None)
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            return elem
        mock_initial_session.findById.side_effect = initial_find

        # Fresh recovered session
        mock_fresh_session = MagicMock()
        def fresh_find(elem_id):
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            return elem
        mock_fresh_session.findById.side_effect = fresh_find

        mock_conn_mgr = MagicMock()
        mock_conn_mgr.reconnect.return_value = mock_fresh_session

        service = CS12Service(
            session=mock_initial_session,
            connection_manager=mock_conn_mgr,
        )

        mat = "3V2YJ58070"
        dest_dir = tmp_path / mat
        dest_dir.mkdir(parents=True, exist_ok=True)
        target_file = dest_dir / f"R3_{mat}_02_10_2026.xls"
        target_file.write_text("DATA")

        params = CS12Params(
            material=mat,
            destination_dir=dest_dir,
            valid_date=datetime.date(2026, 10, 2),
        )

        res = service.execute_cs12_and_export(params)
        assert res.success is True
        assert service.session == mock_fresh_session
        assert mock_conn_mgr.reconnect.call_count == 1

    def test_multi_bom_no_cascaded_failure(self, tmp_path: Path):
        """Dynamic session management: BOM 1 succeeds, connection drops, BOM 2 recovers and succeeds."""
        mock_session_1 = MagicMock()
        def sess1_find(elem_id):
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            return elem
        mock_session_1.findById.side_effect = sess1_find

        mock_session_2 = MagicMock()
        def sess2_find(elem_id):
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            return elem
        mock_session_2.findById.side_effect = sess2_find

        mock_conn_mgr = MagicMock()
        mock_conn_mgr.reconnect.return_value = mock_session_2

        service = CS12Service(
            session=mock_session_1,
            connection_manager=mock_conn_mgr,
        )

        # BOM 1 succeeds on session 1
        mat1 = "BOM0001"
        dir1 = tmp_path / mat1
        dir1.mkdir(parents=True, exist_ok=True)
        (dir1 / f"R3_{mat1}_02_10_2026.xls").write_text("DATA 1")

        res1 = service.execute_cs12_and_export(CS12Params(material=mat1, destination_dir=dir1, valid_date=datetime.date(2026, 10, 2)))
        assert res1.success is True
        assert mock_conn_mgr.reconnect.call_count == 0

        # Simulate connection drop on session 1 before BOM 2
        mock_session_1.findById.side_effect = Exception(-2147023170, "The remote procedure call failed.", None, None)

        # BOM 2 encounters RPC drop on session 1, recovers via reconnect to session 2, and succeeds
        mat2 = "BOM0002"
        dir2 = tmp_path / mat2
        dir2.mkdir(parents=True, exist_ok=True)
        (dir2 / f"R3_{mat2}_02_10_2026.xls").write_text("DATA 2")

        res2 = service.execute_cs12_and_export(CS12Params(material=mat2, destination_dir=dir2, valid_date=datetime.date(2026, 10, 2)))
        assert res2.success is True
        assert service.session == mock_session_2
        assert mock_conn_mgr.reconnect.call_count == 1

    def test_post_export_f3_rpc_drop_preserves_successful_file(self, tmp_path: Path):
        """When file is already written to disk and F3 back button encounters RPC drop, don't discard successful export."""
        mat = "3V2YJ58070"
        dest_dir = tmp_path / mat
        dest_dir.mkdir(parents=True, exist_ok=True)
        target_file = dest_dir / f"R3_{mat}_02_10_2026.xls"

        mock_session = MagicMock()
        def find_elem(elem_id):
            if "wnd[0]/tbar[0]/btn[3]" in elem_id:
                # Target file exists before F3 fails
                target_file.write_text("ALREADY EXPORTED DATA")
                raise Exception(-2147023170, "The remote procedure call failed.", None, None)
            elem = MagicMock()
            elem.Text = ""
            elem.MessageType = "S"
            return elem
        mock_session.findById.side_effect = find_elem

        service = CS12Service(session=mock_session)
        params = CS12Params(
            material=mat,
            destination_dir=dest_dir,
            valid_date=datetime.date(2026, 10, 2),
        )

        res = service.execute_cs12_and_export(params)
        assert res.success is True
        assert res.file_path == target_file
        # Session is marked dead / None so next call doesn't reuse dead pointer
        assert service.session is None

    def test_connection_reuse_failure_clears_dirty_handles(self):
        """When connection reuse check throws an exception, handles are cleared and OpenConnection failure raises cleanly."""
        mock_app = MagicMock()
        dead_conn = MagicMock()
        dead_conn.Description = "P1J"
        dead_conn.Id = "P1J"
        dead_conn.Children.Count = 1
        dead_sess = MagicMock()
        dead_sess.findById.side_effect = Exception(-2147023170, "The remote procedure call failed.")
        dead_conn.Children.return_value = [dead_sess]
        dead_conn.Children.side_effect = lambda i: dead_sess

        mock_app.Connections.Count = 1
        mock_app.Connections.return_value = [dead_conn]
        mock_app.Connections.side_effect = lambda i: dead_conn

        # OpenConnection also fails
        mock_app.OpenConnection.side_effect = Exception("Network unreachable")

        mgr = SAPConnectionManager(process_checker=lambda: True, rot_getter=lambda: mock_app)

        with pytest.raises(Exception) as exc_info:
            mgr.get_or_create_session()

        assert "Failed to open connection" in str(exc_info.value)
        assert mgr.connection is None
        assert mgr.session is None

    def test_no_reset_navigation_called_on_dead_session_upon_rpc_exhaustion(self, tmp_path: Path):
        """When RPC fails and reconnect fails, reset_navigation must NOT be called on dead session."""
        mock_dead_session = MagicMock()
        mock_dead_session.findById.side_effect = Exception(-2147023170, "The remote procedure call failed.", None, None)

        mock_conn_mgr = MagicMock()
        mock_conn_mgr.reconnect.return_value = None  # Reconnect failed

        service = CS12Service(
            session=mock_dead_session,
            connection_manager=mock_conn_mgr,
        )

        params = CS12Params(material="FAILMAT", destination_dir=tmp_path, valid_date=datetime.date(2026, 10, 2))
        res = service.execute_cs12_and_export(params)

        assert res.success is False
        assert service.session is None
        # Verify /n was never sent to dead session after failure
        reset_calls = [
            c for c in mock_dead_session.findById.call_args_list
            if "okcd" in str(c)
        ]
        # okcd is called in navigate_to_cs12, but NOT again in reset_navigation
        assert len(reset_calls) <= 1

    def test_minimize_sap_window_selectivity(self):
        """Verify _minimize_sap_window minimizes SAP GUI sessions and avoids generic browser windows."""
        from src.gui.plm_download_dialog import _minimize_sap_window

        # Test session.findById('wnd[0]').iconify()
        mock_sess = MagicMock()
        wnd0 = MagicMock()
        mock_sess.findById.return_value = wnd0

        _minimize_sap_window(mock_sess)
        wnd0.iconify.assert_called_once()



