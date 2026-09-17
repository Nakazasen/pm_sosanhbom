"""Headless browser driver factory, CDP download configuration, and session lifecycle management for TC14.

Supports Microsoft Edge and Google Chrome with modern headless mode (--headless=new),
CDP Page.setDownloadBehavior, cookie persistence, and session timeout detection.
"""

from __future__ import annotations

import json
import logging
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.remote.webdriver import WebDriver

from .selectors import TC14Selectors

logger = logging.getLogger(__name__)


@dataclass
class BrowserConfig:
    """Configuration options for headless browser sessions."""

    browser_type: str = "edge"  # "edge" or "chrome"
    headless: bool = True
    window_size: str = "1920,1080"
    download_dir: Optional[Path] = None
    page_load_timeout: int = 45
    implicit_wait: int = 5
    cookie_file: Optional[Path] = None
    session_timeout_seconds: int = 1800  # 30 minutes


def create_driver(
    browser_type: str = "edge",
    download_dir: Optional[Path] = None,
    headless: bool = True,
    window_size: str = "1920,1080",
    extra_arguments: Optional[List[str]] = None,
) -> WebDriver:
    """Factory creating configured headless WebDriver (Edge with Chrome fallback).

    Configures:
    - --headless=new for modern Chromium rendering
    - --window-size=1920,1080 to prevent virtual DOM clipping
    - CDP Page.setDownloadBehavior for background downloads
    - Experimental download preferences for automatic file handling

    Args:
        browser_type: Preferred browser: 'edge' or 'chrome'.
        download_dir: Target directory where exported files will be saved.
        headless: Whether to run headlessly. Defaults to True.
        window_size: Browser dimensions. Defaults to '1920,1080'.
        extra_arguments: Optional list of additional command line flags.

    Returns:
        Configured WebDriver instance.

    Raises:
        WebDriverException: If neither Edge nor Chrome could be initialized.
    """
    resolved_download_dir = Path(download_dir or Path.cwd() / "downloads").resolve()
    resolved_download_dir.mkdir(parents=True, exist_ok=True)

    normalized_browser = browser_type.strip().lower()
    browsers_to_try = (
        ["edge", "chrome"] if normalized_browser == "edge" else ["chrome", "edge"]
    )

    last_exception: Optional[Exception] = None
    driver: Optional[WebDriver] = None

    for candidate in browsers_to_try:
        try:
            logger.info("Attempting to initialize %s WebDriver...", candidate)
            if candidate == "edge":
                driver = _create_edge_driver(
                    resolved_download_dir, headless, window_size, extra_arguments
                )
            else:
                driver = _create_chrome_driver(
                    resolved_download_dir, headless, window_size, extra_arguments
                )
            logger.info("Successfully initialized %s WebDriver.", candidate)
            break
        except Exception as exc:
            logger.warning("Failed to initialize %s WebDriver: %s", candidate, exc)
            last_exception = exc

    if driver is None:
        raise WebDriverException(
            f"Failed to initialize any WebDriver (attempted {browsers_to_try}): {last_exception}"
        ) from last_exception

    # Apply CDP download behavior
    _enable_cdp_download_behavior(driver, resolved_download_dir)

    return driver


def _build_download_prefs(download_dir: Path) -> Dict[str, Any]:
    """Construct browser preferences dictionary for seamless file downloads."""
    return {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    }


def _create_edge_driver(
    download_dir: Path,
    headless: bool,
    window_size: str,
    extra_args: Optional[List[str]],
) -> WebDriver:
    """Create and configure Microsoft Edge WebDriver."""
    options = EdgeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument(f"--window-size={window_size}")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")

    if extra_args:
        for arg in extra_args:
            options.add_argument(arg)

    options.add_experimental_option("prefs", _build_download_prefs(download_dir))
    return webdriver.Edge(options=options)


def _create_chrome_driver(
    download_dir: Path,
    headless: bool,
    window_size: str,
    extra_args: Optional[List[str]],
) -> WebDriver:
    """Create and configure Google Chrome WebDriver."""
    options = ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument(f"--window-size={window_size}")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")

    if extra_args:
        for arg in extra_args:
            options.add_argument(arg)

    options.add_experimental_option("prefs", _build_download_prefs(download_dir))
    return webdriver.Chrome(options=options)


def _enable_cdp_download_behavior(driver: WebDriver, download_dir: Path) -> None:
    """Configure Chrome DevTools Protocol to allow file downloads headlessly."""
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": str(download_dir.resolve()),
            },
        )
        logger.debug("Configured CDP Page.setDownloadBehavior: %s", download_dir)
    except Exception as exc:
        logger.warning(
            "Could not set CDP download behavior (may not be supported on this remote driver): %s",
            exc,
        )


class TC14SessionManager:
    """Manages WebDriver lifecycle, session persistence, cookies, and timeout detection."""

    def __init__(
        self,
        config: Optional[BrowserConfig] = None,
        driver_factory: Optional[Callable[[], WebDriver]] = None,
        driver: Optional[WebDriver] = None,
    ) -> None:
        self.config = config or BrowserConfig()
        self._driver_factory = driver_factory
        self._driver: Optional[WebDriver] = driver
        self._last_activity_time: float = 0.0
        self._authenticated: bool = False
        self._current_username: Optional[str] = None

    @property
    def driver(self) -> WebDriver:
        """Get the active WebDriver instance, creating one if necessary."""
        if self._driver is None:
            self.start_session()
        assert self._driver is not None
        return self._driver

    @property
    def is_authenticated(self) -> bool:
        """Return whether current session has successfully authenticated."""
        return self._authenticated

    @property
    def last_activity_time(self) -> float:
        """Return timestamp of the most recent operation."""
        return self._last_activity_time

    def start_session(self) -> WebDriver:
        """Initialize or return the underlying WebDriver instance."""
        if self._driver is not None:
            return self._driver

        if self._driver_factory:
            self._driver = self._driver_factory()
        else:
            self._driver = create_driver(
                browser_type=self.config.browser_type,
                download_dir=self.config.download_dir,
                headless=self.config.headless,
                window_size=self.config.window_size,
            )

        self._driver.set_page_load_timeout(self.config.page_load_timeout)
        self._driver.implicitly_wait(self.config.implicit_wait)
        self.touch()
        logger.info("TC14SessionManager: browser session started.")
        return self._driver

    def touch(self) -> None:
        """Update the last activity timestamp to prevent premature timeout."""
        self._last_activity_time = time.time()

    def mark_authenticated(self, username: str) -> None:
        """Record successful authentication state."""
        self._authenticated = True
        self._current_username = username
        self.touch()

    def mark_unauthenticated(self) -> None:
        """Mark session as unauthenticated / logged out."""
        self._authenticated = False
        self._current_username = None

    def is_session_timed_out(self) -> bool:
        """Check if local session inactivity threshold has been exceeded."""
        if self._last_activity_time <= 0:
            return True
        elapsed = time.time() - self._last_activity_time
        return elapsed > self.config.session_timeout_seconds

    def is_login_page_present(self) -> bool:
        """Inspect current browser DOM to check if the login screen is displayed.

        Returns True if username input or login button is found, indicating expired session.
        """
        if self._driver is None:
            return False

        try:
            current_url = self._driver.current_url.lower()
            if "login" in current_url:
                return True

            by_type, selector = TC14Selectors.USERNAME_INPUT
            elements = self._driver.find_elements(by_type, selector)
            for elem in elements:
                if elem.is_displayed():
                    return True
        except Exception as exc:
            logger.debug("Error checking login page presence: %s", exc)
            return False

        return False

    def is_session_alive(self) -> bool:
        """Verify whether the browser session is currently active and authenticated.

        Returns False if the driver is dead, timed out, or showing the login screen.
        """
        if self._driver is None:
            return False

        if self.is_session_timed_out():
            return False

        try:
            curr_url = self._driver.current_url
            if callable(curr_url):
                curr_url = curr_url()
        except Exception:
            self.mark_unauthenticated()
            return False

        if self.is_login_page_present():
            self.mark_unauthenticated()
            return False

        return self._authenticated

    def save_cookies(self, target_path: Optional[Path] = None) -> List[Dict[str, Any]]:
        """Extract cookies from active browser and optionally persist to JSON file.

        Args:
            target_path: Optional path to save cookies. Defaults to config.cookie_file.

        Returns:
            List of cookie dictionaries.
        """
        if self._driver is None:
            return []

        cookies = self._driver.get_cookies()
        path = target_path or self.config.cookie_file
        if path:
            path = Path(path)
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cookies, f, indent=2)
            logger.info("Saved %d cookies to %s", len(cookies), path)

        return cookies

    def load_cookies(
        self,
        base_url: str,
        source_path: Optional[Path] = None,
        cookies: Optional[List[Dict[str, Any]]] = None,
    ) -> bool:
        """Load cookies into active browser session.

        Note: Navigates to base_url first because cookies can only be set
        for the domain currently loaded.

        Args:
            base_url: The application root URL.
            source_path: Path to JSON cookie file.
            cookies: Direct list of cookies (alternative to file).

        Returns:
            True if cookies were successfully added, False otherwise.
        """
        if self._driver is None:
            self.start_session()
        assert self._driver is not None

        cookie_list: List[Dict[str, Any]] = []
        if cookies is not None:
            cookie_list = cookies
        else:
            path = source_path or self.config.cookie_file
            if path and Path(path).is_file():
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        cookie_list = json.load(f)
                except Exception as exc:
                    logger.warning("Failed to read cookie file %s: %s", path, exc)
                    return False

        if not cookie_list:
            return False

        try:
            # Must be on the domain to add cookies
            self._driver.get(base_url)
            for cookie in cookie_list:
                # Selenium requires domain-safe cookie dict
                safe_cookie = {
                    k: v
                    for k, v in cookie.items()
                    if k in ("name", "value", "path", "domain", "secure", "expiry")
                }
                try:
                    self._driver.add_cookie(safe_cookie)
                except Exception as exc:
                    logger.debug("Could not add cookie %s: %s", safe_cookie.get("name"), exc)
            self.touch()
            logger.info("Loaded %d cookies into browser.", len(cookie_list))
            return True
        except Exception as exc:
            logger.warning("Error applying cookies: %s", exc)
            return False

    def close(self) -> None:
        """Safely terminate the browser session and release all system resources."""
        if self._driver is not None:
            try:
                self._driver.quit()
                logger.info("TC14SessionManager: WebDriver terminated.")
            except Exception as exc:
                logger.warning("Error during driver quit: %s", exc)
            finally:
                self._driver = None
                self._authenticated = False
                self._current_username = None

    def __enter__(self) -> "TC14SessionManager":
        self.start_session()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()
