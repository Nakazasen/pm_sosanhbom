"""Feature F15: SAP GUI 770 COM Automation Isolation Tests.

Verifies:
1. SAPCredentials model and password masking security.
2. Process lifecycle monitoring for saplogon.exe.
3. Reusing existing open SAP connections and sessions.
4. Handling missing executable with SAPConnectionError.
5. COM ROT (Running Object Table) acquisition timeout detection.
"""

from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.models import SAPConnectionError, SAPCredentials


class TestF15SAPCOMAutomation:
    """Test suite for Feature F15: SAP GUI 770 COM Automation."""

    def test_f15_credentials_configuration(self):
        """Test 1: Credentials configuration and password masking."""
        creds = SAPCredentials(username="v130474", password="SecretPassword123")
        assert creds.username == "v130474"
        assert creds.system == "P1J(ERP60-AWS)-VN"

        # Password must be masked in dictionary/logging
        d = creds.to_dict()
        assert d["password"] == "***"
        assert d["username"] == "v130474"

    def test_f15_process_detection_mock(self):
        """Test 2: Check saplogon.exe process discovery logic."""
        mgr = SAPConnectionManager(process_checker=lambda: True)
        assert mgr.is_saplogon_process_running() is True

        mgr_offline = SAPConnectionManager(process_checker=lambda: False)
        assert mgr_offline.is_saplogon_process_running() is False

    def test_f15_connection_manager_reuse_existing_session(self, mock_sap_session):
        """Test 3: Reuse already active connection from application.Connections."""
        mock_app = MagicMock()
        mock_conn = MagicMock()
        mock_conn.Description = "P1J(ERP60-AWS)-VN"
        mock_conn.Children.Count = 1
        mock_conn.Children.return_value = mock_sap_session
        mock_conn.Children.side_effect = lambda idx: mock_sap_session

        mock_app.Connections.Count = 1
        mock_app.Connections.return_value = mock_conn
        mock_app.Connections.side_effect = lambda idx: mock_conn

        mock_com = MagicMock()
        mock_gui_auto = MagicMock()
        mock_gui_auto.GetScriptingEngine.return_value = mock_app
        mock_com.GetObject.return_value = mock_gui_auto

        mgr = SAPConnectionManager(
            com_provider=mock_com,
            process_checker=lambda: True,
        )
        mgr.is_logged_in = MagicMock(return_value=True)

        session = mgr.get_or_create_session(timeout_sec=1.0)
        assert session == mock_sap_session

    def test_f15_connection_manager_launch_failure_handling(self, tmp_path: Path):
        """Test 4: Raise SAPConnectionError when saplogon.exe does not exist."""
        nonexistent = tmp_path / "fake_saplogon.exe"
        creds = SAPCredentials(saplogon_path=str(nonexistent))

        mgr = SAPConnectionManager(credentials=creds, process_checker=lambda: False)
        with pytest.raises(SAPConnectionError, match="SAP Logon executable not found"):
            mgr.launch_saplogon()

    def test_f15_rot_engine_timeout(self):
        """Test 5: Raise SAPConnectionError when SAPGUI ROT registration times out."""
        mock_com = MagicMock()
        mock_com.GetObject.side_effect = Exception("ROT object not found")

        mgr = SAPConnectionManager(
            com_provider=mock_com,
            process_checker=lambda: True,
        )

        with pytest.raises(SAPConnectionError, match="Timed out"):
            mgr.ensure_saplogon_running(timeout_sec=0.5)
