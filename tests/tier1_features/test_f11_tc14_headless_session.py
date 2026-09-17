"""Feature F11: TC14 Headless Browser Session Isolation Tests.

Verifies:
1. BrowserConfig defaults (--headless=new, 1920x1080, download prefs).
2. Session manager lifecycle management with driver factory.
3. Inactivity timeout detection when threshold is exceeded.
4. Activity touch method refreshes session lease.
5. Clean resource teardown on session close.
"""

import time
from pathlib import Path
from unittest.mock import MagicMock
import pytest

from src.automation.tc14.session import BrowserConfig, TC14SessionManager, _build_download_prefs


class TestF11TC14HeadlessSession:
    """Test suite for Feature F11: TC14 Headless Browser Session."""

    def test_f11_browser_config_defaults(self, tmp_path: Path):
        """Test 1: Default configuration adheres to modern headless standards."""
        config = BrowserConfig(download_dir=tmp_path)
        assert config.headless is True
        assert config.window_size == "1920,1080"
        assert config.browser_type == "edge"
        assert config.session_timeout_seconds == 1800

        # Verify download preferences
        prefs = _build_download_prefs(tmp_path)
        assert prefs["download.prompt_for_download"] is False
        assert prefs["download.default_directory"] == str(tmp_path.resolve())

    def test_f11_session_manager_initialization(self, mock_tc14_driver):
        """Test 2: TC14SessionManager initializes and injects driver factory."""
        factory = MagicMock(return_value=mock_tc14_driver)
        mgr = TC14SessionManager(driver_factory=factory)

        driver = mgr.driver
        assert driver == mock_tc14_driver
        factory.assert_called_once()
        assert mgr.last_activity_time > 0

    def test_f11_session_inactivity_timeout(self, mock_tc14_driver):
        """Test 3: Detect session timeout after inactivity."""
        config = BrowserConfig(session_timeout_seconds=5)
        mgr = TC14SessionManager(config=config, driver_factory=lambda: mock_tc14_driver)
        mgr.start_session()

        # Set last activity 10 seconds in the past
        mgr._last_activity_time = time.time() - 10.0
        assert mgr.is_session_timed_out() is True

    def test_f11_session_touch_refreshes_activity(self, mock_tc14_driver):
        """Test 4: Activity touch updates timestamp and prevents timeout."""
        config = BrowserConfig(session_timeout_seconds=10)
        mgr = TC14SessionManager(config=config, driver_factory=lambda: mock_tc14_driver)
        mgr.start_session()

        # Artificial past activity
        mgr._last_activity_time = time.time() - 8.0
        # Call touch to simulate user operation
        mgr.touch()
        assert mgr.is_session_timed_out() is False

    def test_f11_session_close_cleans_resources(self, mock_tc14_driver):
        """Test 5: Session close safely quits driver and resets states."""
        mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        mgr.start_session()
        mgr.mark_authenticated("vn_pe03")
        assert mgr.is_authenticated is True

        mgr.close()
        mock_tc14_driver.quit.assert_called_once()
        assert mgr.is_authenticated is False
        assert mgr._driver is None
