"""
SAP R3 Data Models & Domain Exceptions for CS12 Automation.
"""

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional, Union


# ============================================================================
# Exceptions
# ============================================================================

class SAPError(Exception):
    """Base exception for SAP automation operations."""
    pass


class SAPConnectionError(SAPError):
    """Raised when connecting or authenticating to SAP GUI / system fails."""
    pass


class SAPCS12Error(SAPError):
    """Raised when CS12 transaction execution or export encounters an error."""
    pass


class SAPParseError(SAPError):
    """Raised when parsing exported SAP R3 spreadsheet data fails."""
    pass


# ============================================================================
# Data Models
# ============================================================================

DEFAULT_SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
DEFAULT_SAP_SYSTEM = "P1J(ERP60-AWS)-VN"


@dataclass
class SAPCredentials:
    """Authentication and environment configuration for SAP Logon."""
    username: str = "v130474"
    password: str = "0123456789"
    language: str = "EN"
    system: str = DEFAULT_SAP_SYSTEM
    saplogon_path: str = DEFAULT_SAPLOGON_PATH

    def to_dict(self) -> Dict[str, Any]:
        """Dictionary representation with masked password for logging."""
        return {
            "username": self.username,
            "password": "***" if self.password else "",
            "language": self.language,
            "system": self.system,
            "saplogon_path": self.saplogon_path,
        }


@dataclass
class CS12Params:
    """Parameters for SAP Transaction CS12 (Multilevel BOM Display & Export)."""
    material: str
    plant: str = "2200"
    bom_usage: str = "pp01"
    alternative: str = "01"
    valid_date: date = field(default_factory=date.today)
    destination_dir: Optional[Path] = None

    def __post_init__(self):
        self.material = str(self.material).strip()
        self.plant = str(self.plant).strip()
        self.bom_usage = str(self.bom_usage).strip()
        self.alternative = str(self.alternative).strip().zfill(2)
        if isinstance(self.destination_dir, str):
            self.destination_dir = Path(self.destination_dir)

    @property
    def formatted_date_sap(self) -> str:
        """Date string formatted for SAP input (YYYY/MM/DD)."""
        return self.valid_date.strftime("%Y/%m/%d")

    @property
    def formatted_date_filename(self) -> str:
        """Date string formatted for file naming (DD_MM_YYYY)."""
        return self.valid_date.strftime("%d_%m_%Y")

    @property
    def expected_filename(self) -> str:
        """Expected export filename adhering to legacy naming conventions:
        R3_<material>_<dd>_<mm>_<yyyy>.xls
        """
        return f"R3_{self.material}_{self.formatted_date_filename}.xls"

    @property
    def target_file_path(self) -> Optional[Path]:
        """Complete destination file path if destination_dir is specified."""
        if self.destination_dir:
            return self.destination_dir / self.expected_filename
        return None


@dataclass
class ExportResult:
    """Outcome of an SAP CS12 BOM export operation."""
    success: bool
    material: str
    file_path: Optional[Path] = None
    error_message: Optional[str] = None
    rows_count: int = 0
    status_code: str = ""  # 'S' = Success, 'E' = Error, 'W' = Warning
    execution_time_sec: float = 0.0
    timestamp: datetime = field(default_factory=datetime.now)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "success": self.success,
            "material": self.material,
            "file_path": str(self.file_path) if self.file_path else None,
            "error_message": self.error_message,
            "rows_count": self.rows_count,
            "status_code": self.status_code,
            "execution_time_sec": round(self.execution_time_sec, 3),
            "timestamp": self.timestamp.isoformat(),
        }


@dataclass
class R3ComponentRow:
    """Typed representation of an extracted R3 BOM component row."""
    part_code: str
    quantity: float
    rev_r3: str = ""
    item_num: Optional[str] = None
    description: Optional[str] = None
    level: Optional[int] = None

    def __post_init__(self):
        self.part_code = str(self.part_code).strip() if self.part_code is not None else ""
        try:
            self.quantity = float(self.quantity)
        except (ValueError, TypeError):
            self.quantity = 0.0

        if self.rev_r3 is None:
            self.rev_r3 = ""
        elif isinstance(self.rev_r3, float):
            import math
            if math.isnan(self.rev_r3):
                self.rev_r3 = ""
            elif self.rev_r3.is_integer():
                int_val = int(self.rev_r3)
                self.rev_r3 = f"{int_val:02d}" if 0 <= int_val < 100 else str(int_val)
            else:
                self.rev_r3 = str(self.rev_r3).strip()
        elif isinstance(self.rev_r3, int):
            self.rev_r3 = f"{self.rev_r3:02d}" if 0 <= self.rev_r3 < 100 else str(self.rev_r3)
        else:
            s = str(self.rev_r3).strip()
            if s.lower() == "nan":
                self.rev_r3 = ""
            elif s.isdigit() and len(s) == 1:
                self.rev_r3 = f"0{s}"
            else:
                try:
                    f_val = float(s)
                    if f_val.is_integer():
                        int_val = int(f_val)
                        self.rev_r3 = f"{int_val:02d}" if 0 <= int_val < 100 else str(int_val)
                    else:
                        self.rev_r3 = s
                except ValueError:
                    self.rev_r3 = s

        if self.item_num is not None:
            self.item_num = str(self.item_num).strip()
        if self.description is not None:
            self.description = str(self.description).strip()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "part_code": self.part_code,
            "quantity": self.quantity,
            "rev_r3": self.rev_r3,
            "item_num": self.item_num,
            "description": self.description,
            "level": self.level,
        }
