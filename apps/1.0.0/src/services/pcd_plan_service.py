"""PCD Monthly Production Plan scanner service.

Scans PCD SharePoint / OneDrive folder structure:
Theo tháng (月別) / <Year> / <YearMonth> /
Picks end-of-period file (*定期計画後明細計画.xlsx) over beginning-of-period file (*月初明細計画.xlsx).
Filters sheet '詳細日程' where Column H (履歴) == 'NEW' and Column D (Material) starts with 'T1' or '11'.
"""

from __future__ import annotations

import datetime
import logging
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, List, Optional

import openpyxl

from src.services.machine_dict_service import MachineDictService, MachineInfo

logger = logging.getLogger(__name__)

DEFAULT_ONEDRIVE_BASE = Path(r"C:\Users\tvn183660\OneDrive - KYOCERA Document Solutions")


@dataclass
class PCDPlanItem:
    """Represents a filtered production plan item."""

    material_code: str          # Cột D (Material): e.g. T10C0TZUS0, 1102Y43AX0
    machine_code_4char: str     # 4-character machine code e.g. 0C0T, 02Y4
    history_status: str         # Cột H (履歴): e.g. NEW
    is_trial: bool              # True if starts with T1 (DMT/PMT), False if starts with 11 (MP)
    suggested_phase: str        # 'DMT/PMT' for T1, 'MP' for 11
    machine_info: Optional[MachineInfo] = None  # Matched from dictionary
    machine_name: str = ""      # e.g. 6th Next, Virgo, Libra 2...
    raw_data: dict[str, Any] = field(default_factory=dict)


@dataclass
class PCDPlanScanResult:
    """Result of scanning a monthly plan folder/file."""

    year: int
    month: int
    plan_file: Path
    file_type: str              # 'end_of_period' (*定期計画後明細計画) or 'beginning_of_period' (*月初明細計画)
    total_rows_scanned: int = 0
    items: list[PCDPlanItem] = field(default_factory=list)


class PCDPlanService:
    """Service to discover and parse PCD monthly production plans."""

    def __init__(
        self,
        base_dir: str | Path | None = None,
        dict_service: MachineDictService | None = None,
    ) -> None:
        self.base_dir = Path(base_dir or DEFAULT_ONEDRIVE_BASE)
        self.dict_service = dict_service or MachineDictService()

    def find_monthly_plan_file(
        self,
        year: int | None = None,
        month: int | None = None,
        custom_folder: str | Path | None = None,
    ) -> Optional[tuple[Path, str]]:
        """Find the plan file for given year and month.

        Priority:
        1. End-of-period (*定期計画後明細計画.xlsx)
        2. Beginning-of-period (*月初明細計画.xlsx)
        """
        now = datetime.datetime.now()
        target_year = year or now.year
        target_month = month or now.month
        year_str = str(target_year)
        month_str = f"{target_year}{target_month:02d}"

        search_dirs: list[Path] = []

        if custom_folder:
            search_dirs.append(Path(custom_folder))

        # Check in OneDrive sync locations
        candidate_paths = [
            self.base_dir / "Theo tháng (月別)" / year_str / month_str,
            self.base_dir / year_str / month_str,
            self.base_dir / "ProductionPlan" / "Theo tháng (月別)" / year_str / month_str,
            self.base_dir,
        ]
        for p in candidate_paths:
            if p.exists() and p.is_dir():
                search_dirs.append(p)

        logger.info("Searching for monthly plan in %d candidate directories for %s/%s", len(search_dirs), year_str, month_str)

        # Look in candidate directories
        for sdir in search_dirs:
            try:
                files = list(sdir.glob("*.xlsx"))
                # Priority 1: End of period
                for f in files:
                    if "定期計画後明細計画" in f.name:
                        return f, "end_of_period"

                # Priority 2: Beginning of period
                for f in files:
                    if "月初明細計画" in f.name:
                        return f, "beginning_of_period"

                # Fallback: any plan file matching month pattern
                for f in files:
                    if f"{target_month}月" in f.name or month_str in f.name:
                        return f, "custom_plan"
            except Exception as e:
                logger.debug("Error checking directory %s: %s", sdir, e)

        return None

    def parse_plan_file(self, file_path: str | Path) -> PCDPlanScanResult:
        """Parse Excel plan file, read sheet '詳細日程', filter Column H='NEW' & Column D starts with 'T1' or '11'."""
        path = Path(file_path).resolve()
        if not path.exists():
            raise FileNotFoundError(f"Plan file does not exist: {path}")

        # Ensure dictionary is loaded
        self.dict_service.load()

        file_type = "end_of_period" if "定期計画後明細計画" in path.name else "beginning_of_period"
        now = datetime.datetime.now()

        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

        target_sheet = None
        for sname in wb.sheetnames:
            if "詳細日程" in sname:
                target_sheet = sname
                break

        if not target_sheet:
            target_sheet = wb.sheetnames[0]
            logger.warning("Sheet '詳細日程' not found in %s, falling back to '%s'", path.name, target_sheet)

        ws = wb[target_sheet]

        # Scan header row to identify Material (Col D) and History (Col H)
        header_row = 1
        mat_col_idx = 4  # Column D (1-indexed)
        history_col_idx = 8  # Column H (1-indexed)

        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            for c_idx, val in enumerate(row, 1):
                sval = str(val or "").strip()
                if sval in ("Material", "品目", "Mã hàng", "Mã vật tư"):
                    mat_col_idx = c_idx
                    header_row = r_idx
                elif sval in ("履歴", "History", "Trạng thái", "Status"):
                    history_col_idx = c_idx

        items: list[PCDPlanItem] = []
        total_rows = 0

        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            total_rows += 1
            if not row or len(row) < max(mat_col_idx, history_col_idx):
                continue

            val_mat = str(row[mat_col_idx - 1] or "").strip().upper()
            val_history = str(row[history_col_idx - 1] or "").strip().upper()

            # Filter Condition: History == 'NEW' and Material starts with 'T1' or '11'
            if val_history == "NEW" and (val_mat.startswith("T1") or val_mat.startswith("11")):
                is_trial = val_mat.startswith("T1")
                suggested_phase = "DMT/PMT" if is_trial else "MP"

                # Extract 4-char machine code (usually characters at index 2..6: e.g. T1 0C0T ZUS0 -> 0C0T)
                machine_code_4char = val_mat[2:6] if len(val_mat) >= 6 else ""

                # Lookup in machine dictionary
                matched_machine = self.dict_service.extract_and_lookup_material(val_mat)
                machine_name = matched_machine.machine_name if matched_machine else ""

                item = PCDPlanItem(
                    material_code=val_mat,
                    machine_code_4char=machine_code_4char,
                    history_status=val_history,
                    is_trial=is_trial,
                    suggested_phase=suggested_phase,
                    machine_info=matched_machine,
                    machine_name=machine_name,
                    raw_data={"row": total_rows, "material": val_mat, "history": val_history},
                )
                items.append(item)

        wb.close()

        logger.info("Parsed %d rows from %s, found %d NEW materials (T1/11)", total_rows, path.name, len(items))

        return PCDPlanScanResult(
            year=now.year,
            month=now.month,
            plan_file=path,
            file_type=file_type,
            total_rows_scanned=total_rows,
            items=items,
        )
