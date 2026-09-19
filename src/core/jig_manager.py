"""JIG Master Catalog and 4M Change Assessment Management Engine.

Authoritative implementation for Feature F20/F21:
- Loads JIG master catalogs across 15 Kyocera machine series from 'List JIG thay doi, khi bo sung ma hang.xlsx'.
- Preserves formulas, cell values, cell styles, merged cells, and column widths for A3:G50.
- Manages 4M change assessments (Man, Machine, Material, Method) and KTSX sign-offs in V36:V41.
- Conforms to legacy uf_jig.frm, md_jig.bas, and form_ssbom.xlsm Sheet 'List JIG'.
"""

from __future__ import annotations

import copy
import datetime
from enum import Enum
import logging
from pathlib import Path

import openpyxl
import pandas as pd
from pydantic import BaseModel, ConfigDict, Field

logger = logging.getLogger(__name__)

# Canonical list of 15 supported machine series from Sheet 'List JIG' V3:V17
STANDARD_JIG_SERIES: list[str] = [
    "6thA4",
    "Polaris",
    "6thA3",
    "Brazil",
    "Libra",
    "Libra2",
    "Mebius",
    "Kairos",
    "Virgo",
    "Iris",
    "Sirius",
    "EH Iris",
    "Spica",
    "DF Iris",
    "Pictor",
]


class Decision4M(str, Enum):
    """Evaluation decision for 4M change assessment."""

    UNCHECKED = "Hãy xác nhận"
    OK = "OK"
    NG = "NG"


class Presence4M(str, Enum):
    """Presence of 4M change."""

    UNCHECKED = "Hãy xác nhận"
    YES = "Có"
    NO = "Không"


class Assessment4M(BaseModel):
    """4M Change Evaluation and KTSX Sign-off Record."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    man_changed: Presence4M = Field(
        default=Presence4M.UNCHECKED, description="Man (Con người) thay đổi"
    )
    machine_changed: Presence4M = Field(
        default=Presence4M.UNCHECKED, description="Machine (Máy móc/Thiết bị/JIG) thay đổi"
    )
    material_changed: Presence4M = Field(
        default=Presence4M.UNCHECKED, description="Material (Vật liệu/Linh kiện) thay đổi"
    )
    method_changed: Presence4M = Field(
        default=Presence4M.UNCHECKED, description="Method (Phương pháp thao tác) thay đổi"
    )
    overall_evaluation: Decision4M = Field(
        default=Decision4M.UNCHECKED, description="Kết quả đánh giá 4M (OK/NG)"
    )
    ktsx_confirmed: bool = Field(
        default=False, description="Đã xác nhận với KTSX hay chưa"
    )
    ktsx_confirmer: str = Field(
        default="", description="Tên kỹ sư KTSX xác nhận"
    )
    confirmation_date: str = Field(
        default_factory=lambda: datetime.date.today().strftime("%d/%m/%Y"),
        description="Ngày xác nhận KTSX (dd/mm/yyyy)",
    )
    notes: str = Field(default="", description="Ghi chú đánh giá 4M")

    def has_any_change(self) -> bool:
        """Check if any of the 4M dimensions has changed."""
        return any(
            x == Presence4M.YES
            for x in (
                self.man_changed,
                self.machine_changed,
                self.material_changed,
                self.method_changed,
            )
        )

    def presence_summary(self) -> Presence4M:
        """Derive overall presence (Có / Không / Hãy xác nhận)."""
        if self.has_any_change():
            return Presence4M.YES
        # If all 4 are NO, summary is NO
        if all(
            x == Presence4M.NO
            for x in (
                self.man_changed,
                self.machine_changed,
                self.material_changed,
                self.method_changed,
            )
        ):
            return Presence4M.NO
        return Presence4M.UNCHECKED


class JIGManager:
    """Manages JIG master catalog ingestion, sheet population, and 4M verification."""

    def __init__(self, master_catalog_path: str | Path | None = None) -> None:
        self.master_catalog_path: Path | None = (
            Path(master_catalog_path) if master_catalog_path else None
        )

    @staticmethod
    def get_supported_series() -> list[str]:
        """Return canonical list of 15 supported machine series."""
        return list(STANDARD_JIG_SERIES)

    @staticmethod
    def is_supported_series(series_name: str) -> bool:
        """Check if series name belongs to standard 15 series."""
        target = series_name.strip().lower()
        return any(s.lower() == target for s in STANDARD_JIG_SERIES)

    @staticmethod
    def find_matching_sheet_name(
        available_sheets: list[str], series_name: str
    ) -> str | None:
        """Find matching sheet name case-insensitively and ignoring whitespace."""
        target = series_name.strip().lower()
        for s in available_sheets:
            if s.strip().lower() == target:
                return s
        return None

    def load_series_catalog(
        self,
        series_name: str,
        master_path: str | Path | None = None,
    ) -> pd.DataFrame:
        """Load DataFrame of JIG fixtures for a specific machine series from master file.

        Reads range equivalent to A3:G50.
        """
        target_path = Path(master_path) if master_path else self.master_catalog_path
        if not target_path or not target_path.exists():
            logger.warning("Master catalog file does not exist: %s", target_path)
            return pd.DataFrame()

        try:
            excel_file = pd.ExcelFile(target_path)
            matched_sheet = self.find_matching_sheet_name(
                excel_file.sheet_names, series_name
            )
            if not matched_sheet:
                logger.warning(
                    "Sheet for series '%s' not found in %s", series_name, target_path
                )
                return pd.DataFrame()

            # Read rows 3 to 50 (header at row 3 = index 2), up to 48 rows
            df = pd.read_excel(
                target_path,
                sheet_name=matched_sheet,
                header=2,
                nrows=48,
                usecols="A:G",
            )
            return df
        except Exception as exc:
            logger.error("Failed to load JIG catalog for series '%s': %s", series_name, exc)
            return pd.DataFrame()

    def populate_jig_sheet(
        self,
        target_workbook_path: str | Path,
        series_name: str,
        master_path: str | Path | None = None,
    ) -> bool:
        """Populate Sheet 'List JIG' in target workbook preserving formulas and formatting.

        Recreates legacy uf_jig.frm cmd_taijig_Click behavior:
        1. Open target workbook, access Sheet 'List JIG'.
        2. UnMerge and ClearContents on Range A3:H50.
        3. Open master workbook, read sheet for `series_name` (with formulas preserved).
        4. Copy A3:G50 with full formulas, values, and cell styles.
        5. Copy merged cell configurations within A3:G50.
        6. Copy or autofit column widths for A:G.
        7. Save target workbook.
        """
        src_path = Path(master_path) if master_path else self.master_catalog_path
        dest_path = Path(target_workbook_path)

        if not dest_path.exists():
            logger.error("Target workbook does not exist: %s", dest_path)
            return False
        if not src_path or not src_path.exists():
            logger.error("Source master catalog does not exist: %s", src_path)
            return False

        src_wb = None
        dest_wb = None
        try:
            # data_only=False ensures formulas are preserved as formulas (e.g. '=VLOOKUP(...)')
            src_wb = openpyxl.load_workbook(src_path, data_only=False)
            matched_sheet = self.find_matching_sheet_name(src_wb.sheetnames, series_name)
            if not matched_sheet:
                logger.warning("Series sheet '%s' not found in master catalog.", series_name)
                src_wb.close()
                return False

            ws_src = src_wb[matched_sheet]
            dest_wb = openpyxl.load_workbook(dest_path)

            if "List JIG" not in dest_wb.sheetnames:
                ws_dest = dest_wb.create_sheet("List JIG")
            else:
                ws_dest = dest_wb["List JIG"]

            # 1. Unmerge cells intersecting A3:H50
            merged_ranges = list(ws_dest.merged_cells.ranges)
            for m_range in merged_ranges:
                if (
                    m_range.min_row >= 3
                    and m_range.max_row <= 50
                    and m_range.min_col >= 1
                    and m_range.max_col <= 8
                ):
                    ws_dest.unmerge_cells(str(m_range))

            # 2. Clear contents and styles for A3:H50
            for r in range(3, 51):
                for c in range(1, 9):
                    cell = ws_dest.cell(row=r, column=c)
                    cell.value = None

            # 3. Copy cells A3:G50 from source sheet preserving formula and styles
            for r in range(3, 51):
                for c in range(1, 8):
                    src_cell = ws_src.cell(row=r, column=c)
                    dest_cell = ws_dest.cell(row=r, column=c)
                    dest_cell.value = src_cell.value

                    if src_cell.has_style:
                        if src_cell.font:
                            dest_cell.font = copy.copy(src_cell.font)
                        if src_cell.border:
                            dest_cell.border = copy.copy(src_cell.border)
                        if src_cell.fill:
                            dest_cell.fill = copy.copy(src_cell.fill)
                        if src_cell.number_format:
                            dest_cell.number_format = src_cell.number_format
                        if src_cell.alignment:
                            dest_cell.alignment = copy.copy(src_cell.alignment)

            # 4. Replicate merged cells from source A3:G50
            for m_range in ws_src.merged_cells.ranges:
                if (
                    m_range.min_row >= 3
                    and m_range.max_row <= 50
                    and m_range.min_col >= 1
                    and m_range.max_col <= 7
                ):
                    try:
                        ws_dest.merge_cells(str(m_range))
                    except Exception as merge_err:
                        logger.debug("Merged range notice: %s (%s)", m_range, merge_err)

            # 5. Copy column dimensions for A:G
            col_letters = ["A", "B", "C", "D", "E", "F", "G"]
            for col_letter in col_letters:
                if col_letter in ws_src.column_dimensions:
                    src_col_dim = ws_src.column_dimensions[col_letter]
                    if src_col_dim.width:
                        ws_dest.column_dimensions[col_letter].width = src_col_dim.width

            src_wb.close()
            dest_wb.save(dest_path)
            dest_wb.close()
            return True
        except Exception as exc:
            logger.error("Error populating JIG sheet: %s", exc)
            if src_wb:
                try:
                    src_wb.close()
                except Exception:
                    pass
            if dest_wb:
                try:
                    dest_wb.close()
                except Exception:
                    pass
            return False

    def clear_jig_sheet(self, target_workbook_path: str | Path) -> bool:
        """Clear contents and unmerge cells in A3:H50 of Sheet 'List JIG'."""
        dest_path = Path(target_workbook_path)
        if not dest_path.exists():
            return False

        try:
            wb = openpyxl.load_workbook(dest_path)
            if "List JIG" not in wb.sheetnames:
                wb.close()
                return False

            ws = wb["List JIG"]

            # Unmerge cells in A3:H50
            merged_ranges = list(ws.merged_cells.ranges)
            for m_range in merged_ranges:
                if (
                    m_range.min_row >= 3
                    and m_range.max_row <= 50
                    and m_range.min_col >= 1
                    and m_range.max_col <= 8
                ):
                    ws.unmerge_cells(str(m_range))

            # Clear contents in A3:H50
            for r in range(3, 51):
                for c in range(1, 9):
                    ws.cell(row=r, column=c).value = None

            wb.save(dest_path)
            wb.close()
            return True
        except Exception as exc:
            logger.error("Failed to clear JIG sheet: %s", exc)
            return False

    def write_4m_assessment(
        self,
        target_workbook_path: str | Path,
        assessment: Assessment4M,
    ) -> bool:
        """Write 4M assessment and KTSX sign-off values into Sheet 'List JIG' V36:V41.

        Cells:
        - V36: 'Hãy xác nhận' (Header)
        - V37: 'Có' if change exists, else 'Không' (or 'Hãy xác nhận')
        - V39: 'Hãy xác nhận' (Header)
        - V40: 'OK' / 'NG' / 'Hãy xác nhận' (Overall evaluation)
        - V41: Sign-off text with engineer name and date
        """
        dest_path = Path(target_workbook_path)
        if not dest_path.exists():
            logger.error("Target workbook does not exist: %s", dest_path)
            return False

        try:
            wb = openpyxl.load_workbook(dest_path)
            if "List JIG" not in wb.sheetnames:
                ws = wb.create_sheet("List JIG")
            else:
                ws = wb["List JIG"]

            # Headers
            ws["V36"] = "Hãy xác nhận"
            ws["V39"] = "Hãy xác nhận"

            # V37: Has 4M change
            has_change = assessment.presence_summary()
            ws["V37"] = has_change.value

            # V40: Assessment result (OK / NG / Hãy xác nhận)
            ws["V40"] = assessment.overall_evaluation.value

            # V41: KTSX Confirmation note
            if assessment.ktsx_confirmed:
                confirmer = assessment.ktsx_confirmer or "KTSX"
                date_str = assessment.confirmation_date or datetime.date.today().strftime("%d/%m/%Y")
                ws["V41"] = f"KTSX: {confirmer} ({date_str})"
            else:
                ws["V41"] = "Chưa xác nhận KTSX"

            wb.save(dest_path)
            wb.close()
            return True
        except Exception as exc:
            logger.error("Failed to write 4M assessment: %s", exc)
            return False

    def read_4m_assessment(
        self,
        target_workbook_path: str | Path,
    ) -> Assessment4M | None:
        """Read back 4M assessment and KTSX sign-off values from Sheet 'List JIG'."""
        dest_path = Path(target_workbook_path)
        if not dest_path.exists():
            return None

        try:
            wb = openpyxl.load_workbook(dest_path, data_only=True)
            if "List JIG" not in wb.sheetnames:
                wb.close()
                return None

            ws = wb["List JIG"]
            v37_val = str(ws["V37"].value or "").strip()
            v40_val = str(ws["V40"].value or "").strip()
            v41_val = str(ws["V41"].value or "").strip()

            wb.close()

            # Map V37 to presence
            if v37_val == Presence4M.YES.value:
                presence = Presence4M.YES
            elif v37_val == Presence4M.NO.value:
                presence = Presence4M.NO
            else:
                presence = Presence4M.UNCHECKED

            # Map V40 to decision
            if v40_val == Decision4M.OK.value:
                decision = Decision4M.OK
            elif v40_val == Decision4M.NG.value:
                decision = Decision4M.NG
            else:
                decision = Decision4M.UNCHECKED

            # Parse V41 confirmation: check for explicit positive confirmation
            ktsx_confirmed = (
                bool(v41_val)
                and "Chưa xác nhận" not in v41_val
                and (
                    v41_val.startswith("KTSX:")
                    or "Đã xác nhận" in v41_val
                    or "OK" in v41_val.upper()
                )
            )
            confirmer = ""
            date_str = ""

            if ktsx_confirmed and "KTSX:" in v41_val:
                parts = v41_val.replace("KTSX:", "").strip()
                if "(" in parts and ")" in parts:
                    confirmer = parts.split("(")[0].strip()
                    date_str = parts.split("(")[1].replace(")", "").strip()
                else:
                    confirmer = parts

            return Assessment4M(
                man_changed=presence,
                machine_changed=presence,
                material_changed=presence,
                method_changed=presence,
                overall_evaluation=decision,
                ktsx_confirmed=ktsx_confirmed,
                ktsx_confirmer=confirmer,
                confirmation_date=date_str or datetime.date.today().strftime("%d/%m/%Y"),
                notes=v41_val,
            )
        except Exception as exc:
            logger.error("Failed to read 4M assessment: %s", exc)
            return None

    def extract_master_path_from_sheet(
        self, workbook_path: str | Path
    ) -> Path | None:
        """Extract master catalog path from cell V24 of Sheet 'List JIG' if configured."""
        p = Path(workbook_path)
        if not p.exists():
            return None

        try:
            wb = openpyxl.load_workbook(p, data_only=True)
            if "List JIG" not in wb.sheetnames:
                wb.close()
                return None
            val = wb["List JIG"]["V24"].value
            wb.close()
            if val and str(val).strip():
                return Path(str(val).strip())
            return None
        except Exception:
            return None
