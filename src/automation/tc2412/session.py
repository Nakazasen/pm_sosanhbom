"""Headless browser driver factory, CDP download configuration, and session lifecycle management for TC2412.

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

from .selectors import TC2412Selectors

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
            logger.info("Attempting to initialize %s WebDriver for TC2412...", candidate)
            if candidate == "edge":
                driver = _create_edge_driver(
                    resolved_download_dir, headless, window_size, extra_arguments
                )
            else:
                driver = _create_chrome_driver(
                    resolved_download_dir, headless, window_size, extra_arguments
                )
            logger.info("Successfully initialized %s WebDriver for TC2412.", candidate)
            break
        except Exception as exc:
            logger.warning("Failed to initialize %s WebDriver: %s", candidate, exc)
            last_exception = exc

    if driver is None:
        raise WebDriverException(
            f"Failed to initialize any WebDriver (attempted {browsers_to_try}): {last_exception}"
        ) from last_exception

    enable_cdp_download_behavior(driver, resolved_download_dir)
    return driver


def _build_download_prefs(download_dir: Path) -> Dict[str, Any]:
    """Build Chromium download preferences dictionary."""
    return {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,
    }


def _create_edge_driver(
    download_dir: Path,
    headless: bool,
    window_size: str,
    extra_arguments: Optional[List[str]] = None,
) -> WebDriver:
    """Create and configure Microsoft Edge WebDriver."""
    options = EdgeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument(f"--window-size={window_size}")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--ignore-certificate-errors")

    if extra_arguments:
        for arg in extra_arguments:
            options.add_argument(arg)

    prefs = _build_download_prefs(download_dir)
    options.add_experimental_option("prefs", prefs)

    return webdriver.Edge(options=options)


def _create_chrome_driver(
    download_dir: Path,
    headless: bool,
    window_size: str,
    extra_arguments: Optional[List[str]] = None,
) -> WebDriver:
    """Create and configure Google Chrome WebDriver."""
    options = ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument(f"--window-size={window_size}")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--ignore-certificate-errors")

    if extra_arguments:
        for arg in extra_arguments:
            options.add_argument(arg)

    prefs = _build_download_prefs(download_dir)
    options.add_experimental_option("prefs", prefs)

    return webdriver.Chrome(options=options)


def enable_cdp_download_behavior(driver: WebDriver, download_dir: Path) -> None:
    """Enable file downloads in headless Chromium via Chrome DevTools Protocol."""
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allowAndName",
                "downloadPath": str(download_dir.resolve()),
                "eventsEnabled": True,
            },
        )
        logger.debug("Configured CDP Page.setDownloadBehavior to %s", download_dir)
    except Exception as exc:
        logger.warning(
            "Could not set CDP download behavior (may not be supported): %s", exc
        )


class TC2412SessionManager:
    """Manages WebDriver lifecycle, session timeout, and cookie persistence for TC2412."""

    def __init__(
        self,
        config: Optional[BrowserConfig] = None,
        driver_factory: Callable[..., WebDriver] = create_driver,
    ):
        self.config = config or BrowserConfig()
        self._driver_factory = driver_factory
        self._driver: Optional[WebDriver] = None
        self._last_activity: float = time.time()
        self._authenticated: bool = False

    @property
    def driver(self) -> WebDriver:
        """Get active WebDriver, creating one if not yet initialized."""
        if self._driver is None:
            self._driver = self._driver_factory(
                browser_type=self.config.browser_type,
                download_dir=self.config.download_dir,
                headless=self.config.headless,
                window_size=self.config.window_size,
            )
            self._driver.implicitly_wait(self.config.implicit_wait)
            self._driver.set_page_load_timeout(self.config.page_load_timeout)
            self.touch()
        return self._driver

    def touch(self) -> None:
        """Update last activity timestamp."""
        self._last_activity = time.time()

    def mark_authenticated(self) -> None:
        """Mark session as authenticated."""
        self._authenticated = True
        self.touch()

    def mark_unauthenticated(self) -> None:
        """Mark session as unauthenticated/logged out."""
        self._authenticated = False

    def is_session_timed_out(self) -> bool:
        """Check if session has exceeded configured inactivity timeout."""
        elapsed = time.time() - self._last_activity
        return elapsed > self.config.session_timeout_seconds

    def is_login_page_present(self) -> bool:
        """Detect if browser is currently showing the TC2412 login page."""
        if self._driver is None:
            return False
        try:
            current_url = self._driver.current_url.lower()
            if "login" in current_url or "signin" in current_url:
                return True

            by, value = TC2412Selectors.LOGIN_CONTAINER
            elements = self._driver.find_elements(by, value)
            return len(elements) > 0
        except Exception as exc:
            logger.debug("Error checking login page presence: %s", exc)
            return False

    def is_session_alive(self) -> bool:
        """Check if session is authenticated and still valid."""
        if not self._authenticated:
            return False
        if self.is_session_timed_out():
            logger.info("Session timed out due to inactivity.")
            self._authenticated = False
            return False
        if self.is_login_page_present():
            logger.info("Login page detected; session expired.")
            self._authenticated = False
            return False
        return True

    def save_cookies(self, cookie_file: Optional[Path] = None) -> bool:
        """Save browser session cookies to JSON file."""
        target_file = cookie_file or self.config.cookie_file
        if target_file is None or self._driver is None:
            return False
        try:
            target_file.parent.mkdir(parents=True, exist_ok=True)
            cookies = self._driver.get_cookies()
            with open(target_file, "w", encoding="utf-8") as f:
                json.dump(cookies, f, indent=2)
            logger.debug("Saved %d cookies to %s", len(cookies), target_file)
            return True
        except Exception as exc:
            logger.warning("Failed to save cookies to %s: %s", target_file, exc)
            return False

    def load_cookies(self, cookie_file: Optional[Path] = None) -> bool:
        """Load session cookies from JSON file into browser."""
        source_file = cookie_file or self.config.cookie_file
        if source_file is None or not source_file.exists() or self._driver is None:
            return False
        try:
            with open(source_file, "r", encoding="utf-8") as f:
                cookies = json.load(f)
            for cookie in cookies:
                cookie_to_add = {
                    k: v
                    for k, v in cookie.items()
                    if k in ("name", "value", "path", "domain", "secure", "httpOnly")
                }
                try:
                    self._driver.add_cookie(cookie_to_add)
                except Exception as c_exc:
                    logger.debug("Could not add cookie %s: %s", cookie.get("name"), c_exc)
            logger.debug("Loaded cookies from %s", source_file)
            return True
        except Exception as exc:
            logger.warning("Failed to load cookies from %s: %s", source_file, exc)
            return False

    def close(self) -> None:
        """Close WebDriver and clean up resources."""
        if self._driver is not None:
            try:
                self._driver.quit()
                logger.info("Closed TC2412 WebDriver session.")
            except Exception as exc:
                logger.warning("Error quitting TC2412 WebDriver: %s", exc)
            finally:
                self._driver = None
                self._authenticated = False

    def __enter__(self) -> TC2412SessionManager:
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()

