"""Feature F13: TC14 Hash-Based Search & Open Isolation Tests.

Verifies:
1. Deterministic hash-based search URL construction avoiding UI typing delay.
2. Parsing matched item revisions from search cell elements.
3. Raising TC14SearchError when no matching items are returned.
4. Direct object view navigation by UID.
5. Content / occurrence tree tab activation.
"""

from unittest.mock import MagicMock, patch
import pytest
from selenium.webdriver.remote.webelement import WebElement

from src.automation.tc14.client import (
    SearchResult,
    TC14AutomationClient,
    TC14NavigationError,
    TC14SearchError,
)
from src.automation.tc14.selectors import TC14URLs
from src.automation.tc14.session import TC14SessionManager


class TestF13TC14SearchNavigation:
    """Test suite for Feature F13: TC14 Hash-Based Search & Open."""

    def test_f13_direct_hash_url_construction(self):
        """Test 1: Search URL format matches TC14 SPA declarative routing."""
        url = TC14URLs.get_search_url("http://tcmp3gwb:3000/", "110C103NL0")
        assert "searchCriteria=110C103NL0" in url
        assert "isGlobalSearch=true" in url

        # Special characters url-encoded
        special_url = TC14URLs.get_search_url("http://tcmp3gwb:3000/", "PART 123/45#A")
        assert "PART%20123/45%23A" in special_url

    def test_f13_search_results_parsing(self, mock_tc14_driver):
        """Test 2: Parse SearchResult objects from matched DOM cell elements."""
        cell1 = MagicMock(spec=WebElement)
        cell1.is_displayed.return_value = True
        title_elem = MagicMock(spec=WebElement)
        title_elem.text = "110C103NL0 / PBA-00001711 / Revision:A"
        title_elem.get_attribute.return_value = "http://tcmp3gwb:3000/#/showObject?uid=SR_UID_001"
        cell1.find_element.return_value = title_elem

        def mock_find_elems(by, sel):
            if "username" in sel:
                return []
            return [cell1]

        mock_tc14_driver.find_elements.side_effect = mock_find_elems

        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")
        client = TC14AutomationClient(session_manager=session_mgr)

        with patch("src.automation.tc14.client.WebDriverWait") as mock_wait_cls:
            wait_inst = MagicMock()
            mock_wait_cls.return_value = wait_inst
            wait_inst.until.return_value = True

            results = client.search_part("110C103NL0")
            assert len(results) == 1
            assert results[0].item_id == "110C103NL0"
            assert results[0].revision == "A"
            assert results[0].uid == "SR_UID_001"

    def test_f13_search_no_results_raises_error(self, mock_tc14_driver):
        """Test 3: Raise TC14SearchError when query yields 0 results."""
        mock_tc14_driver.find_elements.return_value = []
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")

        client = TC14AutomationClient(session_manager=session_mgr)

        with patch("src.automation.tc14.client.WebDriverWait") as mock_wait_cls:
            wait_inst = MagicMock()
            mock_wait_cls.return_value = wait_inst
            wait_inst.until.return_value = True

            with pytest.raises(TC14SearchError, match="No search results found"):
                client.search_part("NONEXISTENT_PART_123")

    def test_f13_open_item_revision(self, mock_tc14_driver):
        """Test 4: Navigate directly to item view via object URL."""
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")

        client = TC14AutomationClient(session_manager=session_mgr)

        item = SearchResult(item_id="110C103NL0", title="MA4500ifx", uid="UID_TEST_99")
        with patch("src.automation.tc14.client.WebDriverWait") as mock_wait_cls:
            wait_inst = MagicMock()
            mock_wait_cls.return_value = wait_inst
            wait_inst.until.return_value = True

            client.open_item(item)

        expected_url = TC14URLs.get_object_url(client.base_url, "UID_TEST_99")
        mock_tc14_driver.get.assert_called_with(expected_url)

    def test_f13_navigate_to_content_tab(self, mock_tc14_driver):
        """Test 5: Navigate to occurrence table Content tab."""
        tab_elem = MagicMock(spec=WebElement)

        session_mgr = TC14SessionManager(driver_factory=lambda: mock_tc14_driver)
        session_mgr.start_session()
        session_mgr.mark_authenticated("vn_pe03")

        client = TC14AutomationClient(session_manager=session_mgr)

        with patch("src.automation.tc14.client.WebDriverWait") as mock_wait_cls:
            wait_inst = MagicMock()
            mock_wait_cls.return_value = wait_inst
            wait_inst.until.side_effect = [tab_elem, True]

            success = client.navigate_to_content_tab()
            assert success is True
            tab_elem.click.assert_called_once()
