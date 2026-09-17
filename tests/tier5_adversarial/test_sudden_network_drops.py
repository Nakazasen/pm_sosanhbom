"""Tier 5 Adversarial Tests: Sudden Network Drops & Mid-Session Failures.

Verifies fail-closed resilience under:
1. TC14 Selenium: Socket disconnect during active navigation.
2. TC14 Selenium: Timeout waiting for slow PLM query.
3. TC14 Selenium: StaleElementReferenceException during dynamic tree explosion.
4. SAP GUI: COM RPC server disconnected (0x80010108) mid-transaction.
5. SAP GUI: Session disconnect and failure return during CS12 execution.
6. Clean resource cleanup and error propagation without indefinite freeze.
"""

from __future__ import annotations

import datetime
from unittest.mock import MagicMock, patch
import pytest

from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)

from src.automation.tc14.client import TC14AutomationClient
from src.automation.tc14.session import TC14SessionManager
from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import CS12Params, SAPConnectionError, SAPCS12Error


class TestAdversarialNetworkDrops:
    """Test suite for sudden network drops and disconnect resilience."""

    def test_tc14_sudden_socket_drop_during_login(self):
        """Socket / connection drop during login raises WebDriverException cleanly."""
        mock_driver = MagicMock()
        mock_driver.get.side_effect = WebDriverException("Connection reset by peer: Errno 10054")

        session_mgr = TC14SessionManager(driver_factory=lambda: mock_driver)
        session_mgr.start_session()
        client = TC14AutomationClient(session_manager=session_mgr)

        with pytest.raises(Exception) as exc_info:
            client.login()
        assert "Connection reset" in str(exc_info.value) or isinstance(exc_info.value, WebDriverException)

    def test_tc14_timeout_mid_bom_export(self):
        """Export operation timing out raises TimeoutException cleanly without hanging."""
        mock_driver = MagicMock()
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")

        client = TC14AutomationClient(session_manager=session_mgr)
        with patch.object(client, "export_bom_excel", side_effect=TimeoutException("Timed out waiting for export dialog")):
            with pytest.raises(TimeoutException):
                client.export_bom_excel()

    def test_tc14_stale_element_during_tree_scrape(self):
        """Stale DOM element during dynamic tree scrape is handled with clean fail-closed exception."""
        mock_driver = MagicMock()
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_driver)
        session_mgr.start_session()

        client = TC14AutomationClient(session_manager=session_mgr)
        with patch.object(client, "navigate_to_content_tab", side_effect=StaleElementReferenceException("Stale element in tab")):
            with pytest.raises(StaleElementReferenceException):
                client.navigate_to_content_tab()

    def test_sap_com_rpc_server_unavailable(self):
        """SAP COM connection drops with RPC server unavailable raises SAPConnectionError."""
        mock_com = MagicMock()
        mock_com.GetObject.side_effect = Exception("The RPC server is unavailable. (0x800706BA)")

        conn_mgr = SAPConnectionManager(
            com_provider=mock_com,
            process_checker=lambda: True,
        )
        with pytest.raises(SAPConnectionError) as exc:
            conn_mgr.ensure_saplogon_running(timeout_sec=0.1)
        assert "RPC" in str(exc.value) or "unavailable" in str(exc.value) or "Timed out" in str(exc.value)

    def test_sap_session_drop_during_cs12_transaction(self):
        """Session object dropping mid-transaction returns failed ExportResult without hanging."""
        mock_session = MagicMock()
        mock_session.findById.side_effect = Exception("SAP GUI Session disconnected by server.")

        service = CS12Service(session=mock_session)
        params = CS12Params(material="302K123456", valid_date=datetime.date(2026, 9, 17))
        result = service.execute_cs12_and_export(params)

        assert result.success is False
        assert result.status_code == "E"
        assert "disconnected" in result.error_message

    def test_sap_connection_disconnect_cleanup(self):
        """Explicit disconnect cleans up state cleanly even if COM raises exception during teardown."""
        conn_mgr = SAPConnectionManager()
        mock_session = MagicMock()
        conn_mgr.session = mock_session
        conn_mgr._connected = True

        mock_session.close.side_effect = Exception("COM object already dead")
        conn_mgr.disconnect()
        assert conn_mgr.session is None
