"""Feature F12: TC14 Automated Authentication Isolation Tests.

Verifies:
1. Form input filling (vn_pe03/vn_pe03) and login button interaction.
2. Skipping login when session is already authenticated.
3. Automatic session expiration detection via login page presence.
4. Cookie serialization and loading for persistent sessions.
5. Error handling and timeout raising TC14AuthenticationError.
"""

from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.remote.webelement import WebElement

from src.automation.tc14.client import TC14AuthenticationError, TC14AutomationClient
from src.automation.tc14.session import BrowserConfig, TC14SessionManager


class TestF12TC14Authentication:
    """Test suite for Feature F12: TC14 Automated Authentication."""

    def test_f12_login_form_interaction(self, mock_tc14_driver):
        """Test 1: Fills username/password fields and clicks login button."""
        user_input = MagicMock(spec=WebElement)
        pass_input = MagicMock(spec=WebElement)
        login_btn = MagicMock(spec=WebElement)

        def mock_find(by, sel):
            if "username" in sel:
                return user_input
            if "password" in sel:
                return pass_input
            if "button" in sel or "signInButton" in sel:
                return login_btn
            return MagicMock()

        mock_tc14_driver.find_element.side_effect = mock_find
        mock_tc14_driver.find_elements.return_value = []
        mock_tc14_driver.current_url = "http://tcmp3gwb:3000/#/showHome"

        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        client = TC14AutomationClient(session_manager=session_mgr)

        with patch("src.automation.tc14.client.WebDriverWait") as mock_wait_cls:
            wait_inst = MagicMock()
            mock_wait_cls.return_value = wait_inst
            wait_inst.until.side_effect = [user_input, login_btn, True]

            success = client.login("vn_pe03", "vn_pe03")
            assert success is True
            user_input.send_keys.assert_called_with("vn_pe03")
            pass_input.send_keys.assert_called_with("vn_pe03")
            login_btn.click.assert_called_once()
            assert session_mgr.is_authenticated is True

    def test_f12_skip_login_when_already_authenticated(self, mock_tc14_driver):
        """Test 2: Skip login process if session is already verified and active."""
        mock_tc14_driver.current_url = "http://tcmp3gwb:3000/#/showHome"
        mock_tc14_driver.find_elements.return_value = []

        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")
        session_mgr.touch()

        client = TC14AutomationClient(session_manager=session_mgr)
        result = client.login("vn_pe03", "vn_pe03")
        assert result is True
        mock_tc14_driver.get.assert_not_called()

    def test_f12_session_expiration_detection(self, mock_tc14_driver):
        """Test 3: Detect when session expires and returns to login form."""
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")

        # Simulate redirection to login page
        mock_tc14_driver.current_url = "http://tcmp3gwb:3000/#/login"
        assert session_mgr.is_login_page_present() is True
        assert session_mgr.is_session_alive() is False

    def test_f12_cookie_persistence_save_load(self, tmp_path: Path, mock_tc14_driver):
        """Test 4: Extract and restore browser cookies."""
        cookie_file = tmp_path / "tc14_cookies.json"
        config = BrowserConfig(cookie_file=cookie_file)
        mock_tc14_driver.get_cookies.return_value = [
            {"name": "JSESSIONID", "value": "TEST_SESSION_123", "domain": "tcmp3gwb"},
            {"name": "XSRF-TOKEN", "value": "TEST_XSRF_TOKEN", "domain": "tcmp3gwb"},
            {"name": "sanSID", "value": "AUTH_COOKIE_SID", "domain": "tcmp3gwb"},
        ]
        session_mgr = TC14SessionManager(config=config, driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()

        # Save cookies
        saved = session_mgr.save_cookies(cookie_file)
        assert len(saved) == 3
        assert cookie_file.exists()

        # Load cookies back
        loaded = session_mgr.load_cookies("http://tcmp3gwb:3000/", source_path=cookie_file)
        assert loaded is True
        assert mock_tc14_driver.add_cookie.call_count >= 3

    def test_f12_login_timeout_raises_authentication_error(self, mock_tc14_driver):
        """Test 5: Raise TC14AuthenticationError when login form wait times out."""
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        client = TC14AutomationClient(session_manager=session_mgr)

        with patch("src.automation.tc14.client.WebDriverWait") as mock_wait_cls:
            wait_inst = MagicMock()
            mock_wait_cls.return_value = wait_inst
            wait_inst.until.side_effect = TimeoutException("Element not found")

            with pytest.raises(TC14AuthenticationError):
                client.login("vn_pe03", "vn_pe03")
