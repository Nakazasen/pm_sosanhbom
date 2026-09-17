"""
SAP R3 CS12 Automation Module.
Provides COM-based SAP Logon connection management, CS12 multilevel BOM execution,
status bar fail-closed guard, and resilient spreadsheet parsing.
"""

from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.cs12 import CS12Service, SAPCS12Client
from src.automation.sap.models import (
    CS12Params,
    ExportResult,
    R3ComponentRow,
    SAPConnectionError,
    SAPCS12Error,
    SAPCredentials,
    SAPError,
    SAPParseError,
)
from src.automation.sap.parser import ResilientR3Parser, parse_r3_cs12_file

__all__ = [
    "SAPCredentials",
    "CS12Params",
    "ExportResult",
    "R3ComponentRow",
    "SAPError",
    "SAPConnectionError",
    "SAPCS12Error",
    "SAPParseError",
    "SAPConnectionManager",
    "CS12Service",
    "SAPCS12Client",
    "ResilientR3Parser",
    "parse_r3_cs12_file",
]
