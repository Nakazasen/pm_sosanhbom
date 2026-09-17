"""Siemens Teamcenter Active Workspace (TC14) Web Automation Module.

Provides headless browser automation for authentication, part search,
occurrence tree navigation, and native BOM Excel export.
"""

from .client import (
    BOMItem,
    SearchResult,
    TC14AuthenticationError,
    TC14AutomationClient,
    TC14Client,
    TC14Error,
    TC14ExportError,
    TC14NavigationError,
    TC14SearchError,
    TC14TimeoutError,
)
from .selectors import TC14Selectors, TC14URLs
from .session import BrowserConfig, TC14SessionManager, create_driver

__all__ = [
    "TC14AutomationClient",
    "TC14Client",
    "TC14SessionManager",
    "BrowserConfig",
    "create_driver",
    "TC14Selectors",
    "TC14URLs",
    "SearchResult",
    "BOMItem",
    "TC14Error",
    "TC14AuthenticationError",
    "TC14SearchError",
    "TC14NavigationError",
    "TC14ExportError",
    "TC14TimeoutError",
]
