"""Siemens Teamcenter Version 2412 (TC2412) Active Workspace Automation Module.

Exports the authoritative TC2412AutomationClient, ProgressReporter,
and associated data contracts and safety guards.
"""

from __future__ import annotations

from .client import (
    CANONICAL_14_COLUMNS,
    AuthenticationError,
    BOMCorruptFileError,
    BOMExpandTimeoutError,
    BOMValidationError,
    ColumnConfigurationError,
    ColumnOrderMismatchError,
    ContentTabNotFoundError,
    ExcelMenuNotOpenError,
    ExpandMenuNotOpenError,
    ExportDialogTimeoutError,
    ExportDownloadTimeoutError,
    ImportChangesForbiddenError,
    IncompleteDownloadError,
    ItemNotFoundError,
    NoRowsSelectedError,
    PLMNetworkUnreachableError,
    PropertyNotFoundError,
    RootNodeSelectionError,
    SearchTimeoutError,
    TC2412AutomationClient,
    TC2412AutomationError,
    safe_trigger_export_to_excel,
    verify_exported_excel,
)
from .progress_reporter import ProgressReporter
from .selectors import TC2412Selectors, TC2412URLs
from .session import (
    BrowserConfig,
    TC2412SessionManager,
    create_driver,
    enable_cdp_download_behavior,
)

__all__ = [
    # Master Client
    "TC2412AutomationClient",
    # Progress Reporter
    "ProgressReporter",
    # Exceptions
    "TC2412AutomationError",
    "AuthenticationError",
    "PLMNetworkUnreachableError",
    "ItemNotFoundError",
    "SearchTimeoutError",
    "ContentTabNotFoundError",
    "RootNodeSelectionError",
    "BOMExpandTimeoutError",
    "ExpandMenuNotOpenError",
    "NoRowsSelectedError",
    "ExcelMenuNotOpenError",
    "ImportChangesForbiddenError",
    "ExportDialogTimeoutError",
    "PropertyNotFoundError",
    "ColumnOrderMismatchError",
    "ColumnConfigurationError",
    "ExportDownloadTimeoutError",
    "IncompleteDownloadError",
    "BOMValidationError",
    "BOMCorruptFileError",
    # Contracts & Safety
    "CANONICAL_14_COLUMNS",
    "verify_exported_excel",
    "safe_trigger_export_to_excel",
    # Selectors & URLs
    "TC2412Selectors",
    "TC2412URLs",
    # Session Management
    "BrowserConfig",
    "TC2412SessionManager",
    "create_driver",
    "enable_cdp_download_behavior",
]

