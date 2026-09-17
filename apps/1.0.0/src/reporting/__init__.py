"""Consolidated Reporting and Notification Package.

Provides:
- ExcelReportGenerator: Multi-sheet Excel workbook generator matching form_ssbom.xlsm schema.
- OutlookMailer: Outlook COM HTML email automation for reconciliation notifications.
"""

from src.reporting.excel_generator import (
    COLOR_GREEN_BGR,
    COLOR_GREEN_HEX,
    COLOR_RED_BGR,
    COLOR_RED_HEX,
    ExcelReportGenerator,
    generate_consolidated_report,
)
from src.reporting.outlook_mailer import (
    EmailPreview,
    OutlookMailer,
    generate_reconciliation_email,
)

__all__ = [
    "COLOR_GREEN_BGR",
    "COLOR_GREEN_HEX",
    "COLOR_RED_BGR",
    "COLOR_RED_HEX",
    "EmailPreview",
    "ExcelReportGenerator",
    "OutlookMailer",
    "generate_consolidated_report",
    "generate_reconciliation_email",
]
