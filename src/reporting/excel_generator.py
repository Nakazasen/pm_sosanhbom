"""Feature F24: Consolidated Report Generation Module.

Generates consolidated Excel comparison report matching the exact workbook schema
of form_ssbom.xlsm with 7 canonical worksheets:
1. Tongket: Leader overview checklist of all sub-units and overall KPI metrics.
2. List JIG: Required 3-character fixed codes for JIGs and fixture parts.
3. MSI_7980_7990: MSI serialized component & barcode evaluation (9-branch status).
4. CTTT: Detailed line-by-line reconciliation (CTTT vs PLM vs R3).
5. PLM: Full multi-level PLM engineering BOM with reverse omission check.
6. R3: SAP CS12 multi-level ERP BOM.
7. CTTT_Total: Cross-substation aggregated part totals vs machine grand totals.

Applies exact conditional formatting and color vectors matching legacy VBA:
- NG: Red 255 (BGR: 255, Hex Fill: #FFC7CE, Font: #9C0006, Raw RGB: #FF0000)
- OK: Green 6750054 (BGR: 6750054 = 0x66FF66, Hex Fill: #C6EFCE, Font: #006100, Raw RGB: #66FF66)
"""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.core.models import BOMNode, BOMTree
from src.core.reconciliation import ReconciliationResult

logger = logging.getLogger(__name__)

# Exact VBA Color Vector Constants
COLOR_RED_BGR: int = 255          # VBA integer 255: RGB(255, 0, 0)
COLOR_GREEN_BGR: int = 6750054    # VBA integer 6750054: RGB(102, 255, 102) -> 0x66FF66

# Hex representations for Excel openpyxl styling
COLOR_RED_FILL_HEX: str = "FFC7CE"
COLOR_RED_FONT_HEX: str = "9C0006"
COLOR_RED_HEX: str = "FF0000"

COLOR_GREEN_FILL_HEX: str = "C6EFCE"
COLOR_GREEN_FONT_HEX: str = "006100"
COLOR_GREEN_HEX: str = "66FF66"

COLOR_YELLOW_FILL_HEX: str = "FFEB9C"
COLOR_YELLOW_FONT_HEX: str = "9C6500"

COLOR_HEADER_BG_HEX: str = "1F497D"      # Classic Corporate Dark Blue
COLOR_HEADER_FG_HEX: str = "FFFFFF"
COLOR_SUBHEADER_BG_HEX: str = "D9E1F2"   # Light Blue Header
COLOR_BORDER_HEX: str = "D9D9D9"

# Default sub-units monitored by Leader workspace
STANDARD_SUB_UNITS: list[str] = [
    "LSU",
    "DLP",
    "DRUM",
    "IMAGE",
    "FUSER",
    "DP",
    "ISU",
    "HONTAI",
    "FRAME UNIT",
    "ĐIỀU CHỈNH",
]


class ExcelReportGenerator:
    """Consolidated Excel workbook generator matching form_ssbom.xlsm schema."""

    def __init__(self) -> None:
        # Reusable styles
        self.fill_green = PatternFill(start_color=COLOR_GREEN_FILL_HEX, end_color=COLOR_GREEN_FILL_HEX, fill_type="solid")
        self.font_green = Font(name="Calibri", size=10, bold=True, color=COLOR_GREEN_FONT_HEX)

        self.fill_red = PatternFill(start_color=COLOR_RED_FILL_HEX, end_color=COLOR_RED_FILL_HEX, fill_type="solid")
        self.font_red = Font(name="Calibri", size=10, bold=True, color=COLOR_RED_FONT_HEX)

        self.fill_yellow = PatternFill(start_color=COLOR_YELLOW_FILL_HEX, end_color=COLOR_YELLOW_FILL_HEX, fill_type="solid")
        self.font_yellow = Font(name="Calibri", size=10, bold=True, color=COLOR_YELLOW_FONT_HEX)

        self.fill_header = PatternFill(start_color=COLOR_HEADER_BG_HEX, end_color=COLOR_HEADER_BG_HEX, fill_type="solid")
        self.font_header = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_FG_HEX)

        self.fill_subheader = PatternFill(start_color=COLOR_SUBHEADER_BG_HEX, end_color=COLOR_SUBHEADER_BG_HEX, fill_type="solid")
        self.font_subheader = Font(name="Calibri", size=10, bold=True, color="000000")

        thin = Side(style="thin", color=COLOR_BORDER_HEX)
        self.border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.align_center = Alignment(horizontal="center", vertical="center")
        self.align_left = Alignment(horizontal="left", vertical="center")
        self.align_right = Alignment(horizontal="right", vertical="center")

    def generate_report(
        self,
        output_path: str | Path,
        reconciliation: ReconciliationResult | None = None,
        cttt_data: pd.DataFrame | list[dict[str, Any]] | None = None,
        plm_data: BOMTree | pd.DataFrame | list[dict[str, Any]] | None = None,
        r3_data: pd.DataFrame | list[dict[str, Any]] | None = None,
        msi_data: pd.DataFrame | list[dict[str, Any]] | None = None,
        jig_data: pd.DataFrame | list[dict[str, Any]] | None = None,
        model_name: str = "Virgo",
        target_date: str | None = None,
        sub_unit_statuses: dict[str, str] | None = None,
    ) -> Path:
        """Create and save the 7-sheet consolidated comparison workbook.

        Args:
            output_path: Destination path for the generated .xlsx file.
            reconciliation: Existing ReconciliationResult object (if available).
            cttt_data: Line items from CTTT work instructions.
            plm_data: PLM BOM items or BOMTree.
            r3_data: SAP R3 BOM items.
            msi_data: MSI serialized components evaluation data.
            jig_data: Optional JIG list records.
            model_name: Machine model name (Virgo, Libra2, Iris2024, etc.).
            target_date: Validity / comparison date (defaults to today).
            sub_unit_statuses: Dict of sub-unit name -> status (e.g. {'LSU': 'OK', ...}).

        Returns:
            Path to the saved report file.
        """
        dest_path = Path(output_path).resolve()
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        if target_date is None:
            target_date = datetime.date.today().strftime("%d/%m/%Y")

        # Resolve DataFrames from reconciliation or inputs
        df_cttt = self._ensure_dataframe(
            reconciliation.cttt_rows if reconciliation else cttt_data
        )
        df_missing = self._ensure_dataframe(
            reconciliation.plm_missing_rows if reconciliation else None
        )
        df_totals = self._ensure_dataframe(
            reconciliation.cttt_totals if reconciliation else None
        )
        df_msi = self._ensure_dataframe(
            reconciliation.msi_results if reconciliation else msi_data
        )
        df_plm = self._extract_plm_df(plm_data)
        df_r3 = self._ensure_dataframe(r3_data)
        df_jig = self._ensure_dataframe(jig_data)

        overall_status = reconciliation.overall_status if reconciliation else self._compute_overall_status(df_cttt, df_msi, df_totals)

        # Build Workbook
        wb = openpyxl.Workbook()

        # Sheet 1: Tongket
        ws_tongket = wb.active
        ws_tongket.title = "Tongket"
        self._build_sheet_tongket(
            ws_tongket,
            model_name=model_name,
            target_date=target_date,
            overall_status=overall_status,
            df_cttt=df_cttt,
            df_msi=df_msi,
            df_missing=df_missing,
            sub_unit_statuses=sub_unit_statuses,
        )

        # Sheet 2: List JIG
        ws_jig = wb.create_sheet(title="List JIG")
        self._build_sheet_jig(ws_jig, df_jig, model_name=model_name)

        # Sheet 3: MSI_7980_7990
        ws_msi = wb.create_sheet(title="MSI_7980_7990")
        self._build_sheet_msi(ws_msi, df_msi)

        # Sheet 4: CTTT
        ws_cttt = wb.create_sheet(title="CTTT")
        self._build_sheet_cttt(ws_cttt, df_cttt)

        # Sheet 5: PLM
        ws_plm = wb.create_sheet(title="PLM")
        self._build_sheet_plm(ws_plm, df_plm, df_missing=df_missing)

        # Sheet 6: R3
        ws_r3 = wb.create_sheet(title="R3")
        self._build_sheet_r3(ws_r3, df_r3)

        # Sheet 7: CTTT_Total
        ws_total = wb.create_sheet(title="CTTT_Total")
        self._build_sheet_cttt_total(ws_total, df_totals=df_totals, df_cttt=df_cttt)

        wb.save(dest_path)
        logger.info("Successfully generated consolidated report at: %s", dest_path)
        return dest_path

    # =========================================================================
    # Sheet Builders
    # =========================================================================

    def _build_sheet_tongket(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        model_name: str,
        target_date: str,
        overall_status: str,
        df_cttt: pd.DataFrame,
        df_msi: pd.DataFrame,
        df_missing: pd.DataFrame,
        sub_unit_statuses: dict[str, str] | None = None,
    ) -> None:
        """Build Sheet 'Tongket' matching legacy form_ssbom summary layout."""
        ws.views.sheetView[0].showGridLines = True

        # Header Title
        ws.merge_cells("B2:J2")
        ws["B2"] = "BÁO CÁO TỔNG KẾT ĐỐI SOÁT BOM TỰ ĐỘNG (CTTT vs PLM vs R3)"
        ws["B2"].font = Font(name="Calibri", size=14, bold=True, color=COLOR_HEADER_BG_HEX)
        ws["B2"].alignment = self.align_left

        # Project Info Cards
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        metadata = [
            ("Mã Model Máy:", model_name),
            ("Ngày hiệu lực đối soát:", target_date),
            ("Thời gian xuất báo cáo:", now_str),
            ("Phán định tổng thể:", "HOÀN TẤT - OK" if overall_status == "OK" else "CẦN GIẢI TRÌNH - NG"),
        ]

        row = 4
        for label, val in metadata:
            ws.cell(row=row, column=2, value=label).font = Font(name="Calibri", size=11, bold=True)
            c_val = ws.cell(row=row, column=4, value=val)
            c_val.font = Font(name="Calibri", size=11, bold=(label == "Phán định tổng thể:"))
            c_val.border = self.border_thin
            if label == "Phán định tổng thể:":
                if overall_status == "OK":
                    c_val.fill = self.fill_green
                    c_val.font = self.font_green
                else:
                    c_val.fill = self.fill_red
                    c_val.font = self.font_red
            row += 1

        # Metrics Breakdown
        total_parts = len(df_cttt) if not df_cttt.empty else 0
        ng_cttt = 0
        if not df_cttt.empty and "Check" in df_cttt.columns:
            ng_cttt = (df_cttt["Check"] == "NG").sum()
        ok_cttt = total_parts - ng_cttt

        missing_plm = len(df_missing) if not df_missing.empty else 0
        msi_ng = 0
        if not df_msi.empty:
            status_col = "KẾT QUẢ" if "KẾT QUẢ" in df_msi.columns else ("status" if "status" in df_msi.columns else None)
            if status_col:
                msi_ng = (df_msi[status_col] == "NG").sum()

        metrics = [
            ("Tổng số linh kiện CTTT:", total_parts),
            ("Số linh kiện khớp hoàn toàn (OK):", ok_cttt),
            ("Số linh kiện sai khác (NG):", ng_cttt),
            ("Số linh kiện thiếu trên CTTT (PLM Missing):", missing_plm),
            ("Số cảnh báo sai khác Barcode / MSI:", msi_ng),
        ]

        row += 1
        ws.cell(row=row, column=2, value="CHỈ SỐ THỐNG KÊ CHI TIẾT").font = Font(name="Calibri", size=11, bold=True, color=COLOR_HEADER_BG_HEX)
        row += 1

        for label, count in metrics:
            ws.cell(row=row, column=2, value=label).font = Font(name="Calibri", size=10)
            c_count = ws.cell(row=row, column=4, value=count)
            c_count.font = Font(name="Calibri", size=10, bold=True)
            c_count.alignment = self.align_center
            c_count.border = self.border_thin
            if "NG" in label or "thiếu" in label or "cảnh báo" in label:
                if count > 0:
                    c_count.fill = self.fill_red
                    c_count.font = self.font_red
            row += 1

        # Sub-unit status matrix
        row += 2
        sub_unit_title = "TIẾN ĐỘ & TRẠNG THÁI THEO TỪNG CÔNG ĐOẠN / SUB-UNIT"
        ws.cell(row=row, column=2, value=sub_unit_title).font = Font(
            name="Calibri", size=11, bold=True, color=COLOR_HEADER_BG_HEX
        )
        row += 1

        table_headers = ["STT", "Công đoạn / Sub-Unit", "Trạng thái CTTT", "Linh kiện đối soát", "Kết quả"]
        for col_idx, h in enumerate(table_headers, start=2):
            cell = ws.cell(row=row, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row += 1
        unit_map = sub_unit_statuses or {}
        for idx, unit in enumerate(STANDARD_SUB_UNITS, start=1):
            ws.cell(row=row, column=2, value=idx).alignment = self.align_center
            ws.cell(row=row, column=3, value=unit).font = Font(name="Calibri", size=10, bold=True)

            # Count rows for this unit
            unit_count = 0
            unit_has_ng = False
            if not df_cttt.empty and "SUB" in df_cttt.columns:
                sub_df = df_cttt[df_cttt["SUB"].astype(str).str.upper() == unit.upper()]
                unit_count = len(sub_df)
                if "Check" in sub_df.columns:
                    unit_has_ng = (sub_df["Check"] == "NG").any()

            sub_status = unit_map.get(unit, "Đã nộp" if unit_count > 0 else "Chưa nộp")
            c_sub = ws.cell(row=row, column=4, value=sub_status)
            c_sub.alignment = self.align_center

            c_items = ws.cell(row=row, column=5, value=unit_count)
            c_items.alignment = self.align_center

            res_val = "NG" if unit_has_ng else ("OK" if unit_count > 0 else "-")
            c_res = ws.cell(row=row, column=6, value=res_val)
            c_res.alignment = self.align_center

            for c in range(2, 7):
                ws.cell(row=row, column=c).border = self.border_thin

            if res_val == "OK":
                c_res.fill = self.fill_green
                c_res.font = self.font_green
            elif res_val == "NG":
                c_res.fill = self.fill_red
                c_res.font = self.font_red

            row += 1

        self._auto_fit_columns(ws, min_col=2, max_col=7)

    def _build_sheet_jig(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        df_jig: pd.DataFrame,
        model_name: str,
    ) -> None:
        """Build Sheet 'List JIG' for fixture and JIG code tracking."""
        ws.views.sheetView[0].showGridLines = True

        ws["B2"] = f"DANH MỤC LIST JIG THAY ĐỔI & BỔ SUNG (Model: {model_name})"
        ws["B2"].font = Font(name="Calibri", size=13, bold=True, color=COLOR_HEADER_BG_HEX)

        headers = ["STT", "MÃ JIG / DỤNG CỤ", "TÊN JIG / FIXTURE", "MÃ CỐ ĐỊNH 3 KÝ TỰ", "CÔNG ĐOẠN", "GHI CHÚ"]
        row = 4
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=row, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row = 5
        if not df_jig.empty:
            for idx, r_data in df_jig.iterrows():
                ws.cell(row=row, column=2, value=idx + 1).alignment = self.align_center
                ws.cell(row=row, column=3, value=r_data.get("jig_code", r_data.get("MÃ JIG", "")))
                ws.cell(row=row, column=4, value=r_data.get("jig_name", r_data.get("TÊN JIG", "")))
                ws.cell(row=row, column=5, value=r_data.get("fixed_code", r_data.get("MSI", ""))).alignment = self.align_center
                ws.cell(row=row, column=6, value=r_data.get("unit", r_data.get("CÔNG ĐOẠN", "")))
                ws.cell(row=row, column=7, value=r_data.get("note", r_data.get("GHI CHÚ", "")))
                for c in range(2, 8):
                    ws.cell(row=row, column=c).border = self.border_thin
                row += 1
        else:
            # Provide standard baseline placeholder row
            sample_jigs = [
                ("JIG-001", "JIG CANH CHINH LSU", "1HN", "LSU", "Jig định vị laser"),
                ("JIG-002", "JIG DO DIEN AP FUSER", "2NL", "FUSER", "Jig đo điện trở nhiệt"),
            ]
            for idx, (j_code, j_name, m_code, u_name, note) in enumerate(sample_jigs, start=1):
                ws.cell(row=row, column=2, value=idx).alignment = self.align_center
                ws.cell(row=row, column=3, value=j_code)
                ws.cell(row=row, column=4, value=j_name)
                ws.cell(row=row, column=5, value=m_code).alignment = self.align_center
                ws.cell(row=row, column=6, value=u_name)
                ws.cell(row=row, column=7, value=note)
                for c in range(2, 8):
                    ws.cell(row=row, column=c).border = self.border_thin
                row += 1

        self._auto_fit_columns(ws, min_col=2, max_col=7)

    def _build_sheet_msi(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        df_msi: pd.DataFrame,
    ) -> None:
        """Build Sheet 'MSI_7980_7990' with exact 15 columns (A..O) and 9-branch decision formatting."""
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Mã LK Barcode",                    # Col A (1)
            "Mã UNIT/linh kiện bản mạch",        # Col B (2)
            "Tên UNIT",                          # Col C (3)
            "3 ký tự MSI",                       # Col D (4)
            "SEVICE",                            # Col E (5)
            "ABS",                               # Col F (6)
            "Trang CTTT",                        # Col G (7)
            "Điện áp",                           # Col H (8)
            "Loại Label",                        # Col I (9)
            "Tên phụ trách",                     # Col J (10)
            "Kết quả so sánh",                   # Col K (11)
            "Ghi chú",                           # Col L (12)
            "Mã linh kiện trong PLM",            # Col M (13)
            "MSI trong File Fix Serial DLTool",  # Col N (14)
            "SERVICE (Master)",                  # Col O (15)
        ]

        # Write Headers on Row 1
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row = 2
        if not df_msi.empty:
            for _, r in df_msi.iterrows():
                ws.cell(row=row, column=1, value=str(r.get("barcode_part", r.get("Mã LK Barcode", ""))))
                c_unit_code = ws.cell(row=row, column=2, value=str(r.get("unit_code", r.get("Mã UNIT/linh kiện bản mạch", ""))))
                ws.cell(row=row, column=3, value=str(r.get("unit_name", r.get("Tên UNIT", ""))))
                c_msi = ws.cell(row=row, column=4, value=str(r.get("member_code", r.get("3 ký tự MSI", ""))))
                c_service = ws.cell(row=row, column=5, value=str(r.get("member_service", r.get("SEVICE", "-"))))
                ws.cell(row=row, column=6, value=str(r.get("abs", r.get("ABS", "OK"))))
                ws.cell(row=row, column=7, value=str(r.get("page", r.get("Trang CTTT", ""))))
                ws.cell(row=row, column=8, value=str(r.get("voltage", r.get("Điện áp", ""))))
                ws.cell(row=row, column=9, value=str(r.get("label_type", r.get("Loại Label", ""))))
                ws.cell(row=row, column=10, value=str(r.get("person", r.get("Tên phụ trách", ""))))

                # Status and lookups
                status = str(r.get("status", r.get("Kết quả so sánh", "OK"))).upper()
                c_status = ws.cell(row=row, column=11, value=status)
                ws.cell(row=row, column=12, value=str(r.get("reason", r.get("Ghi chú", ""))))
                ws.cell(row=row, column=13, value=str(r.get("plm_part", r.get("Mã linh kiện trong PLM", ""))))
                ws.cell(row=row, column=14, value=str(r.get("master_code", r.get("MSI trong File Fix Serial DLTool", ""))))
                ws.cell(row=row, column=15, value=str(r.get("master_service", r.get("SERVICE (Master)", ""))))

                for c in range(1, 16):
                    ws.cell(row=row, column=c).border = self.border_thin

                # Alignments
                c_msi.alignment = self.align_center
                c_status.alignment = self.align_center

                # Apply exact conditional formatting colors
                # Branch color rules from legacy msi.bas
                raw_red = r.get("highlight_red")
                highlight_red = raw_red if isinstance(raw_red, (list, tuple, set)) else []
                raw_green = r.get("highlight_green")
                highlight_green = raw_green if isinstance(raw_green, (list, tuple, set)) else []

                if status == "OK":
                    c_status.fill = self.fill_green
                    c_status.font = self.font_green
                elif status == "NG":
                    c_status.fill = self.fill_red
                    c_status.font = self.font_red

                # Check explicit highlights from MSIEvaluationResult
                if "B" in highlight_red:
                    c_unit_code.fill = self.fill_red
                    c_unit_code.font = self.font_red
                elif "B" in highlight_green:
                    c_unit_code.fill = self.fill_green
                    c_unit_code.font = self.font_green

                if "D" in highlight_red:
                    c_msi.fill = self.fill_red
                    c_msi.font = self.font_red
                elif "D" in highlight_green:
                    c_msi.fill = self.fill_green
                    c_msi.font = self.font_green

                if "E" in highlight_red:
                    c_service.fill = self.fill_red
                    c_service.font = self.font_red
                elif "E" in highlight_green:
                    c_service.fill = self.fill_green
                    c_service.font = self.font_green

                row += 1

        self._apply_openpyxl_conditional_rule(ws, col_letter="K", start_row=2, end_row=max(row, 2))
        self._auto_fit_columns(ws, min_col=1, max_col=15)

    def _build_sheet_cttt(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        df_cttt: pd.DataFrame,
    ) -> None:
        """Build Sheet 'CTTT' with columns A..R matching legacy three-way reconciliation layout."""
        ws.views.sheetView[0].showGridLines = True

        # Row 1 subheader blocks
        ws.cell(row=1, column=7, value="work instruction and PLM").font = self.font_subheader
        ws.cell(row=1, column=11, value="work instruction and R3").font = self.font_subheader
        ws.cell(row=1, column=15, value="REV LIST").font = self.font_subheader

        headers = [
            "SUB",               # Col A (1)
            "TRANG CTTT",        # Col B (2)
            "MÃ LINH KIỆN",      # Col C (3)
            "TÊN LINH KIỆN",     # Col D (4)
            "SỐ LƯỢNG",          # Col E (5)
            "PHỤ TRÁCH",         # Col F (6)
            "Q.ty (PLM)",         # Col G (7)
            "Compare (PLM Qty)",  # Col H (8)
            "Rev PLM",            # Col I (9)
            "Spare1",            # Col J (10)
            "Qty (R3)",          # Col K (11)
            "Compare (R3 Qty)",  # Col L (12)
            "Rev R3",            # Col M (13)
            "Compare (Rev)",     # Col N (14)
            "Spare2",            # Col O (15)
            "Spare3",            # Col P (16)
            "Giải thích",        # Col Q (17)
            "Check",             # Col R (18)
        ]

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row = 3
        if not df_cttt.empty:
            for _, r in df_cttt.iterrows():
                ws.cell(row=row, column=1, value=str(r.get("SUB", r.get("sub", ""))))
                ws.cell(row=row, column=2, value=str(r.get("TRANG CTTT", r.get("page", ""))))
                ws.cell(row=row, column=3, value=str(r.get("MÃ LINH KIỆN", r.get("part_code", ""))))
                ws.cell(row=row, column=4, value=str(r.get("TÊN LINH KIỆN", r.get("part_name", ""))))

                # Quantities
                c_qty = ws.cell(row=row, column=5, value=float(r.get("SỐ LƯỢNG", r.get("quantity", 0.0))))
                c_qty.alignment = self.align_right

                ws.cell(row=row, column=6, value=str(r.get("PHỤ TRÁCH", r.get("person", ""))))

                c_plm_qty = ws.cell(row=row, column=7, value=float(r.get("Q.ty (PLM)", r.get("plm_qty", 0.0))))
                c_plm_qty.alignment = self.align_right

                comp_plm = str(r.get("Compare (PLM Qty)", r.get("comp_plm_qty", "OK"))).upper()
                c_comp_plm = ws.cell(row=row, column=8, value=comp_plm)
                c_comp_plm.alignment = self.align_center

                ws.cell(row=row, column=9, value=str(r.get("Rev PLM", r.get("plm_rev", ""))))
                ws.cell(row=row, column=10, value="")

                c_r3_qty = ws.cell(row=row, column=11, value=float(r.get("Qty (R3)", r.get("r3_qty", 0.0))))
                c_r3_qty.alignment = self.align_right

                comp_r3 = str(r.get("Compare (R3 Qty)", r.get("comp_r3_qty", "OK"))).upper()
                c_comp_r3 = ws.cell(row=row, column=12, value=comp_r3)
                c_comp_r3.alignment = self.align_center

                ws.cell(row=row, column=13, value=str(r.get("Rev R3", r.get("r3_rev", ""))))

                comp_rev = str(r.get("Compare (Rev)", r.get("comp_rev", "OK"))).upper()
                c_comp_rev = ws.cell(row=row, column=14, value=comp_rev)
                c_comp_rev.alignment = self.align_center

                ws.cell(row=row, column=15, value="")
                ws.cell(row=row, column=16, value="")
                ws.cell(row=row, column=17, value=str(r.get("Giải thích", r.get("explanation", ""))))

                overall_check = str(r.get("Check", r.get("overall_check", "OK"))).upper()
                c_check = ws.cell(row=row, column=18, value=overall_check)
                c_check.alignment = self.align_center

                for c in range(1, 19):
                    ws.cell(row=row, column=c).border = self.border_thin

                # Cell color styling
                self._apply_ok_ng_style(c_comp_plm, comp_plm)
                self._apply_ok_ng_style(c_comp_r3, comp_r3)
                self._apply_ok_ng_style(c_comp_rev, comp_rev)
                self._apply_ok_ng_style(c_check, overall_check)

                row += 1

        # Apply conditional formatting rules to columns H, L, N, R
        for col_letter in ["H", "L", "N", "R"]:
            self._apply_openpyxl_conditional_rule(ws, col_letter=col_letter, start_row=3, end_row=max(row, 3))

        self._auto_fit_columns(ws, min_col=1, max_col=18)

    def _build_sheet_plm(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        df_plm: pd.DataFrame,
        df_missing: pd.DataFrame | None = None,
    ) -> None:
        """Build Sheet 'PLM' containing 14-column TC14 BOM plus Column N Reverse Check."""
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Level",                        # Col A (1)
            "Item Type",                    # Col B (2)
            "Item Id",                      # Col C (3)
            "Has Children",                 # Col D (4)
            "Quantity",                     # Col E (5)
            "1st Parts",                    # Col F (6)
            "2nd BOM Flag",                 # Col G (7)
            "Occurrence Effectivities",     # Col H (8)
            "Item Revision Projects List",  # Col I (9)
            "Item Name",                    # Col J (10)
            "Notice No",                    # Col K (11)
            "Revision",                     # Col L (12)
            "Item Rev Status",              # Col M (13)
            "Check CTTT (Reverse Check)",   # Col N (14)
        ]

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row = 2
        missing_parts_set = set()
        if df_missing is not None and not df_missing.empty:
            for _, mr in df_missing.iterrows():
                p = str(mr.get("part_code", mr.get("Item Id", ""))).strip().upper()
                if p:
                    missing_parts_set.add(p)

        source_df = df_plm if not df_plm.empty else (df_missing if df_missing is not None else pd.DataFrame())

        if not source_df.empty:
            for _, r in source_df.iterrows():
                ws.cell(row=row, column=1, value=r.get("level", 1)).alignment = self.align_center
                ws.cell(row=row, column=2, value=str(r.get("item_type", "Part")))
                part_id = str(r.get("item_id", r.get("part_code", ""))).strip().upper()
                c_id = ws.cell(row=row, column=3, value=part_id)
                ws.cell(row=row, column=4, value=str(r.get("has_children", "False"))).alignment = self.align_center
                ws.cell(row=row, column=5, value=float(r.get("quantity", 1.0))).alignment = self.align_right
                ws.cell(row=row, column=6, value=str(r.get("first_parts", "")))
                ws.cell(row=row, column=7, value=str(r.get("second_bom_flag", "")))
                ws.cell(row=row, column=8, value=str(r.get("occurrence_effectivities", "")))
                ws.cell(row=row, column=9, value=str(r.get("item_revision_projects_list", "")))
                ws.cell(row=row, column=10, value=str(r.get("item_name", "")))
                ws.cell(row=row, column=11, value=str(r.get("notice_no", "")))
                ws.cell(row=row, column=12, value=str(r.get("revision", ""))).alignment = self.align_center
                ws.cell(row=row, column=13, value=str(r.get("item_rev_status", "")))

                # Reverse check value
                is_missing = part_id in missing_parts_set or (df_plm.empty and df_missing is not None and not df_missing.empty)
                check_val = "#N/A (Thiếu CTTT)" if is_missing else "OK"
                c_check = ws.cell(row=row, column=14, value=check_val)
                c_check.alignment = self.align_center

                for c in range(1, 15):
                    ws.cell(row=row, column=c).border = self.border_thin

                if is_missing:
                    c_id.fill = self.fill_red
                    c_id.font = self.font_red
                    c_check.fill = self.fill_red
                    c_check.font = self.font_red
                else:
                    c_check.fill = self.fill_green
                    c_check.font = self.font_green

                row += 1

        self._auto_fit_columns(ws, min_col=1, max_col=14)

    def _build_sheet_r3(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        df_r3: pd.DataFrame,
    ) -> None:
        """Build Sheet 'R3' containing SAP CS12 multi-level ERP BOM."""
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Level",                    # Col A
            "Component (Material)",     # Col B
            "Object Description",       # Col C
            "Component Quantity",       # Col D
            "Unit of Measure",          # Col E
            "Revision Level (RevLev)",  # Col F
            "Valid From",               # Col G
            "Valid To",                 # Col H
        ]

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row = 2
        if not df_r3.empty:
            for _, r in df_r3.iterrows():
                ws.cell(row=row, column=1, value=r.get("level", r.get("Level", 1))).alignment = self.align_center
                ws.cell(row=row, column=2, value=str(r.get("material", r.get("part_code", r.get("Component", "")))))
                ws.cell(row=row, column=3, value=str(r.get("description", r.get("part_name", r.get("Object Description", "")))))
                c_qty = ws.cell(row=row, column=4, value=float(r.get("quantity", r.get("Component Quantity", 0.0))))
                c_qty.alignment = self.align_right
                ws.cell(row=row, column=5, value=str(r.get("uom", r.get("Unit of Measure", "PC")))).alignment = self.align_center
                ws.cell(row=row, column=6, value=str(r.get("revision", r.get("RevLev", "")))).alignment = self.align_center
                ws.cell(row=row, column=7, value=str(r.get("valid_from", r.get("Valid From", "")))).alignment = self.align_center
                ws.cell(row=row, column=8, value=str(r.get("valid_to", r.get("Valid To", "")))).alignment = self.align_center

                for c in range(1, 9):
                    ws.cell(row=row, column=c).border = self.border_thin
                row += 1

        self._auto_fit_columns(ws, min_col=1, max_col=8)

    def _build_sheet_cttt_total(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        df_totals: pd.DataFrame,
        df_cttt: pd.DataFrame,
    ) -> None:
        """Build Sheet 'CTTT_Total' for cross-substation shared parts aggregation."""
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "MÃ LINH KIỆN",                  # Col A (1)
            "TỔNG SL CTTT",                  # Col B (2)
            "TỔNG SL R3",                    # Col C (3)
            "CHÊNH LỆCH (Diff)",             # Col D (4)
            "SO SÁNH SL R3 (Compare)",       # Col E (5)
            "TRẠNG THÁI",                    # Col F (6)
            "GIẢI THÍCH BÊN SHEET CTTT",     # Col G (7)
            "TÊN LINH KIỆN",                 # Col H (8)
            "PHỤ TRÁCH",                     # Col I (9)
        ]

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_thin

        row = 2
        if not df_totals.empty:
            for _, r in df_totals.iterrows():
                ws.cell(row=row, column=1, value=str(r.get("part_code", r.get("Row Labels", ""))))

                c_cttt = ws.cell(row=row, column=2, value=float(r.get("cttt_total_qty", r.get("CTTT_TOTAL", 0.0))))
                c_cttt.alignment = self.align_right

                c_r3 = ws.cell(row=row, column=3, value=float(r.get("r3_total_qty", r.get("R3_TOTAL", 0.0))))
                c_r3.alignment = self.align_right

                diff = float(r.get("diff_qty", r.get("DIFF", 0.0)))
                c_diff = ws.cell(row=row, column=4, value=diff)
                c_diff.alignment = self.align_right

                comp = str(r.get("comp_r3_total", r.get("STATUS", "OK"))).upper()
                c_comp = ws.cell(row=row, column=5, value=comp)
                c_comp.alignment = self.align_center

                status = str(r.get("STATUS", comp)).upper()
                c_status = ws.cell(row=row, column=6, value=status)
                c_status.alignment = self.align_center

                ws.cell(row=row, column=7, value=str(r.get("explanation", "")))
                ws.cell(row=row, column=8, value=str(r.get("part_name", "")))
                ws.cell(row=row, column=9, value=str(r.get("person", "")))

                for c in range(1, 10):
                    ws.cell(row=row, column=c).border = self.border_thin

                self._apply_ok_ng_style(c_comp, comp)
                self._apply_ok_ng_style(c_status, status)
                row += 1

        self._apply_openpyxl_conditional_rule(ws, col_letter="E", start_row=2, end_row=max(row, 2))
        self._apply_openpyxl_conditional_rule(ws, col_letter="F", start_row=2, end_row=max(row, 2))
        self._auto_fit_columns(ws, min_col=1, max_col=9)

    # =========================================================================
    # Helpers & Formatting Utilities
    # =========================================================================

    def _apply_ok_ng_style(self, cell: openpyxl.cell.cell.Cell, status: str) -> None:
        """Apply Green for OK, Red for NG, or Yellow for warning."""
        if status == "OK":
            cell.fill = self.fill_green
            cell.font = self.font_green
        elif status == "NG":
            cell.fill = self.fill_red
            cell.font = self.font_red
        elif "WARNING" in status or "CẢNH BÁO" in status:
            cell.fill = self.fill_yellow
            cell.font = self.font_yellow

    def _apply_openpyxl_conditional_rule(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        col_letter: str,
        start_row: int,
        end_row: int,
    ) -> None:
        """Register native OpenXML conditional formatting rules into worksheet."""
        cell_range = f"{col_letter}{start_row}:{col_letter}{end_row}"

        rule_ok = CellIsRule(
            operator="equal",
            formula=['"OK"'],
            stopIfTrue=True,
            fill=self.fill_green,
            font=self.font_green,
        )
        rule_ng = CellIsRule(
            operator="equal",
            formula=['"NG"'],
            stopIfTrue=True,
            fill=self.fill_red,
            font=self.font_red,
        )

        ws.conditional_formatting.add(cell_range, rule_ok)
        ws.conditional_formatting.add(cell_range, rule_ng)

    def _auto_fit_columns(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        min_col: int = 1,
        max_col: int = 20,
    ) -> None:
        """Adjust column widths based on cell content length."""
        for col in range(min_col, max_col + 1):
            col_letter = get_column_letter(col)
            max_len = 0
            for row in range(1, min(ws.max_row + 1, 100)):  # sample first 100 rows for speed
                val = ws.cell(row=row, column=col).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    @staticmethod
    def _ensure_dataframe(data: Any) -> pd.DataFrame:
        """Ensure input is converted to a pandas DataFrame."""
        if data is None:
            return pd.DataFrame()
        if isinstance(data, pd.DataFrame):
            return data.copy()
        if isinstance(data, list):
            return pd.DataFrame(data)
        return pd.DataFrame()

    @staticmethod
    def _extract_plm_df(plm_data: Any) -> pd.DataFrame:
        """Extract flat DataFrame from BOMTree, list of nodes, or DataFrame."""
        if plm_data is None:
            return pd.DataFrame()
        if isinstance(plm_data, pd.DataFrame):
            return plm_data.copy()
        if isinstance(plm_data, BOMTree):
            return plm_data.to_dataframe()
        if isinstance(plm_data, list):
            records = []
            for item in plm_data:
                if isinstance(item, BOMNode):
                    records.append({
                        "level": item.level,
                        "item_type": "Assembly" if item.has_children else "Part",
                        "item_id": item.item_id,
                        "has_children": str(item.has_children),
                        "quantity": item.quantity,
                        "occurrence_effectivities": item.effectivity,
                        "item_name": item.item_name,
                        "revision": item.revision,
                    })
                elif isinstance(item, dict):
                    records.append(item)
            return pd.DataFrame(records)
        return pd.DataFrame()

    @staticmethod
    def _compute_overall_status(
        df_cttt: pd.DataFrame,
        df_msi: pd.DataFrame,
        df_totals: pd.DataFrame,
    ) -> str:
        """Determine overall reconciliation status ('OK' or 'NG')."""
        if not df_cttt.empty and "Check" in df_cttt.columns:
            if (df_cttt["Check"] == "NG").any():
                return "NG"
        if not df_msi.empty:
            status_col = "KẾT QUẢ" if "KẾT QUẢ" in df_msi.columns else ("status" if "status" in df_msi.columns else None)
            if status_col and (df_msi[status_col] == "NG").any():
                return "NG"
        if not df_totals.empty and "STATUS" in df_totals.columns:
            if (df_totals["STATUS"] == "NG").any():
                return "NG"
        return "OK"


# High-level procedural shortcut
def generate_consolidated_report(
    output_path: str | Path,
    reconciliation: ReconciliationResult | None = None,
    **kwargs: Any,
) -> Path:
    """Generate consolidated report workbook matching form_ssbom.xlsm."""
    generator = ExcelReportGenerator()
    return generator.generate_report(output_path=output_path, reconciliation=reconciliation, **kwargs)
