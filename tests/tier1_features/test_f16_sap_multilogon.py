"""Feature F16: SAP Multi-Logon Resolution Isolation Tests.

Verifies:
1. Detecting multi-logon license popup dialog (wnd[1]).
2. Selecting option 2 (radMULTI_LOGON_OPT2) to terminate previous orphaned sessions.
3. Graceful pass-through when multi-logon dialog does not appear.
4. Completing login parameter binding (BNAME, BCODE, LANGU) post-resolution.
5. Verifying successful authentication via okcd transaction box.
"""

from unittest.mock import MagicMock
import pytest

from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.models import SAPCredentials


class TestF16SAPMultiLogon:
    """Test suite for Feature F16: SAP Multi-Logon Resolution."""

    def test_f16_detect_multi_logon_dialog(self, mock_sap_session):
        """Test 1: Identify radMULTI_LOGON_OPT2 radio button when dialog appears."""
        rad_btn = MagicMock()
        confirm_btn = MagicMock()

        def find_elem(elem_id):
            if "radMULTI_LOGON_OPT2" in elem_id:
                return rad_btn
            if "wnd[1]/tbar[0]/btn[0]" in elem_id:
                return confirm_btn
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem
        mgr = SAPConnectionManager(credentials=SAPCredentials())
        mgr.is_logged_in = MagicMock(return_value=True)

        mgr.handle_multi_logon_and_login(mock_sap_session)
        rad_btn.Select.assert_called_once()
        confirm_btn.press.assert_called_once()

    def test_f16_select_option2_and_confirm(self, mock_sap_session):
        """Test 2: Radio option 2 is explicitly focused and selected."""
        rad_btn = MagicMock()
        mock_sap_session.findById.side_effect = lambda eid: rad_btn if "radMULTI_LOGON_OPT2" in eid else MagicMock()
        mgr = SAPConnectionManager()
        mgr.is_logged_in = MagicMock(return_value=True)

        mgr.handle_multi_logon_and_login(mock_sap_session)
        rad_btn.Select.assert_called_once()
        rad_btn.SetFocus.assert_called_once()

    def test_f16_no_dialog_skips_cleanly(self, mock_sap_session):
        """Test 3: Normal login continues when multi-logon dialog is absent."""
        def find_elem(elem_id):
            if "wnd[1]" in elem_id:
                raise Exception("Control not found")
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem
        mgr = SAPConnectionManager()
        mgr.is_logged_in = MagicMock(return_value=True)

        # Should complete without error
        mgr.handle_multi_logon_and_login(mock_sap_session)

    def test_f16_enter_credentials_after_dialog(self, mock_sap_session):
        """Test 4: BNAME, BCODE, and LANGU are populated into wnd[0]/usr."""
        bname = MagicMock()
        bcode = MagicMock()
        langu = MagicMock()

        def find_elem(elem_id):
            if "txtRSYST-BNAME" in elem_id:
                return bname
            if "pwdRSYST-BCODE" in elem_id:
                return bcode
            if "txtRSYST-LANGU" in elem_id:
                return langu
            if "wnd[1]" in elem_id:
                raise Exception("No multi-logon popup")
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem
        creds = SAPCredentials(username="v130474", password="test_password", language="EN")
        mgr = SAPConnectionManager(credentials=creds)
        mgr.is_logged_in = MagicMock(return_value=True)

        mgr.handle_multi_logon_and_login(mock_sap_session)
        assert bname.Text == "v130474"
        assert bcode.Text == "test_password"
        assert langu.Text == "EN"

    def test_f16_verify_logged_in_state(self, mock_sap_session):
        """Test 5: is_logged_in verifies presence of transaction box (okcd)."""
        okcd = MagicMock()

        def find_elem(eid):
            if "txtRSYST-BNAME" in eid:
                raise Exception("Not on login screen")
            if "okcd" in eid:
                return okcd
            return MagicMock()

        mock_sap_session.findById.side_effect = find_elem

        mgr = SAPConnectionManager()
        assert mgr.is_logged_in(mock_sap_session) is True

        # When okcd raises exception -> False
        mock_sap_session.findById.side_effect = Exception("Not logged in")
        assert mgr.is_logged_in(mock_sap_session) is False
