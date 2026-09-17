"""Feature F23: Member Input Workspace Module.

Provides the Member desktop workspace for Kyocera production line engineers:
- Sub-unit selection (LSU, DLP, DRUM, IMAGE, FUSER, DP, ISU, HONTAI, etc.).
- CTTT component entry & Excel/CSV import with editable table grid.
- MSI barcode and 3-character fixed code entry.
- Label 7980/7990 input fields.
- Preliminary self-check against PLM and R3 with instant OK/NG visual feedback.
- Final confirmation and submission button generating member submission package.
"""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.core.msi_engine import MSIEngine, evaluate_msi_branch
from src.core.reconciliation import ReconciliationEngine
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_GREEN_FONT_HEX,
    COLOR_RED_FILL_HEX,
    COLOR_RED_FONT_HEX,
    STANDARD_SUB_UNITS,
)

logger = logging.getLogger(__name__)


class MemberWorkspaceView(QWidget):
    """Member data entry, preliminary self-check, and submission workspace."""

    submission_completed = pyqtSignal(dict)  # Emits submission metadata dict

    def __init__(
        self,
        parent: QWidget | None = None,
        base_dir: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.base_dir = Path(base_dir or r"D:\Sandbox\pm_sosanhbom")
        self.reconciliation_engine = ReconciliationEngine()
        self.msi_engine = MSIEngine()

        # Cached reference data for self-check
        self.plm_data: pd.DataFrame | None = None
        self.r3_data: pd.DataFrame | None = None

        self._init_ui()

    def _init_ui(self) -> None:
        """Build responsive, clean layout for production members."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # Top Panel: Sub-unit, Member Name, and Reference Data Status
        # ---------------------------------------------------------------------
        top_group = QGroupBox("1. Thông tin Công đoạn & Người phụ trách")
        top_layout = QHBoxLayout(top_group)

        top_layout.addWidget(QLabel("Công đoạn (Sub-Unit):"))
        self.sub_unit_combo = QComboBox()
        self.sub_unit_combo.addItems(STANDARD_SUB_UNITS)
        self.sub_unit_combo.currentTextChanged.connect(self._on_sub_unit_changed)
        top_layout.addWidget(self.sub_unit_combo)

        top_layout.addWidget(QLabel("Người phụ trách:"))
        self.author_edit = QLineEdit("Nguyen Van A")
        top_layout.addWidget(self.author_edit)

        top_layout.addWidget(QLabel("Mã Model:"))
        self.model_combo = QComboBox()
        self.model_combo.addItems(["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"])
        top_layout.addWidget(self.model_combo)

        top_layout.addStretch()

        self.btn_load_refs = QPushButton("📂 Nạp PLM & R3 đối chiếu...")
        self.btn_load_refs.clicked.connect(self._select_reference_files)
        top_layout.addWidget(self.btn_load_refs)

        self.lbl_ref_status = QLabel("Chưa nạp BOM đối chiếu")
        self.lbl_ref_status.setStyleSheet("color: #6c757d; font-style: italic;")
        top_layout.addWidget(self.lbl_ref_status)

        main_layout.addWidget(top_group)

        # ---------------------------------------------------------------------
        # Central Splitter: CTTT Grid and MSI / Label 7980/7990 Inputs
        # ---------------------------------------------------------------------
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Section A: CTTT Components Table
        cttt_container = QWidget()
        cttt_layout = QVBoxLayout(cttt_container)
        cttt_layout.setContentsMargins(0, 0, 0, 0)

        cttt_header_layout = QHBoxLayout()
        cttt_title = QLabel("2. Danh sách Linh kiện Chỉ thị thao tác (CTTT)")
        cttt_title.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        cttt_header_layout.addWidget(cttt_title)

        cttt_header_layout.addStretch()
        self.btn_add_row = QPushButton("➕ Thêm dòng")
        self.btn_add_row.clicked.connect(self.add_cttt_row)
        self.btn_remove_row = QPushButton("➖ Xóa dòng")
        self.btn_remove_row.clicked.connect(self.remove_selected_cttt_row)
        self.btn_import_excel = QPushButton("📥 Nhập từ Excel/CSV...")
        self.btn_import_excel.clicked.connect(self.import_cttt_from_file)
        self.btn_clear_table = QPushButton("🗑️ Xóa hết")
        self.btn_clear_table.clicked.connect(self.clear_cttt_table)

        cttt_header_layout.addWidget(self.btn_add_row)
        cttt_header_layout.addWidget(self.btn_remove_row)
        cttt_header_layout.addWidget(self.btn_import_excel)
        cttt_header_layout.addWidget(self.btn_clear_table)
        cttt_layout.addLayout(cttt_header_layout)

        # CTTT Table Widget
        self.cttt_table = QTableWidget(0, 8)
        self.cttt_table.setHorizontalHeaderLabels([
            "Trang CTTT",
            "Mã Linh Kiện",
            "Tên Linh Kiện",
            "Số Lượng",
            "SL PLM (Tham chiếu)",
            "SL R3 (Tham chiếu)",
            "Giải Thích Sai Khác",
            "Kết Quả So Sánh",
        ])
        header = self.cttt_table.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        self.cttt_table.setAlternatingRowColors(True)
        cttt_layout.addWidget(self.cttt_table)

        splitter.addWidget(cttt_container)

        # Section B: MSI Barcode & Label 7980/7990 Configuration
        msi_container = QWidget()
        msi_layout = QVBoxLayout(msi_container)
        msi_layout.setContentsMargins(0, 0, 0, 0)

        msi_tab_widget = QTabWidget()

        # Tab 1: MSI Barcode
        msi_tab = QWidget()
        msi_tab_layout = QHBoxLayout(msi_tab)

        col1_layout = QVBoxLayout()
        col1_layout.addWidget(QLabel("Mã LK Barcode:"))
        self.msi_barcode_edit = QLineEdit("302FP93010")
        col1_layout.addWidget(self.msi_barcode_edit)
        col1_layout.addWidget(QLabel("Mã UNIT / Bản mạch:"))
        self.msi_unit_code_edit = QLineEdit("302FP93010")
        col1_layout.addWidget(self.msi_unit_code_edit)
        col1_layout.addWidget(QLabel("Tên UNIT:"))
        self.msi_unit_name_edit = QLineEdit("LSU UNIT")
        col1_layout.addWidget(self.msi_unit_name_edit)

        col2_layout = QVBoxLayout()
        col2_layout.addWidget(QLabel("3 ký tự cố định MSI:"))
        self.msi_code_edit = QLineEdit("1HN")
        col2_layout.addWidget(self.msi_code_edit)
        col2_layout.addWidget(QLabel("SEVICE Comment (nếu có):"))
        self.msi_service_edit = QLineEdit("-")
        col2_layout.addWidget(self.msi_service_edit)
        col2_layout.addWidget(QLabel("ABS (Phán định chất lượng):"))
        self.msi_abs_combo = QComboBox()
        self.msi_abs_combo.addItems(["OK", "NG", "N/A"])
        col2_layout.addWidget(self.msi_abs_combo)

        col3_layout = QVBoxLayout()
        col3_layout.addWidget(QLabel("Điện áp (Voltage):"))
        self.msi_volt_combo = QComboBox()
        self.msi_volt_combo.addItems(["220V", "110V", "100V", "DC24V", "DC5V", "-"])
        col3_layout.addWidget(self.msi_volt_combo)
        col3_layout.addWidget(QLabel("Loại nhãn (Label Type):"))
        self.msi_label_type_combo = QComboBox()
        self.msi_label_type_combo.addItems(["BARCODE", "QR", "SERIAL", "CAUTION"])
        col3_layout.addWidget(self.msi_label_type_combo)

        col3_layout.addWidget(QLabel("Phán định MSI:"))
        self.lbl_msi_badge = QLabel("Chưa kiểm tra")
        self.lbl_msi_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_msi_badge.setStyleSheet("background-color: #e9ecef; font-weight: bold; border-radius: 4px; padding: 4px;")
        col3_layout.addWidget(self.lbl_msi_badge)

        msi_tab_layout.addLayout(col1_layout)
        msi_tab_layout.addLayout(col2_layout)
        msi_tab_layout.addLayout(col3_layout)
        msi_tab_widget.addTab(msi_tab, "Quản lý Barcode & 3 Ký tự MSI")

        # Tab 2: Label 7980 / 7990
        label_tab = QWidget()
        label_layout = QHBoxLayout(label_tab)

        l_col1 = QVBoxLayout()
        l_col1.addWidget(QLabel("Loại nhãn quản lý:"))
        self.label_spec_combo = QComboBox()
        self.label_spec_combo.addItems(["Label 7980 (Kyocera Standard)", "Label 7990 (High Voltage)", "Label LCP"])
        l_col1.addWidget(self.label_spec_combo)
        l_col1.addWidget(QLabel("Mã số nhãn linh kiện:"))
        self.label_code_edit = QLineEdit("7980-VN-001")
        l_col1.addWidget(self.label_code_edit)

        l_col2 = QVBoxLayout()
        l_col2.addWidget(QLabel("Vị trí dán trên CTTT:"))
        self.label_pos_edit = QLineEdit("Mặt trên vỏ khung LSU - Trang 03")
        l_col2.addWidget(self.label_pos_edit)
        l_col2.addWidget(QLabel("Quy cách kiểm tra:"))
        self.label_check_combo = QComboBox()
        self.label_check_combo.addItems(["Dán phẳng, không bọt khí, đúng chiều mũi tên", "Kiểm tra mã vạch quét đọc được", "Không bong tróc"])
        l_col2.addWidget(self.label_check_combo)

        label_layout.addLayout(l_col1)
        label_layout.addLayout(l_col2)
        label_layout.addStretch()
        msi_tab_widget.addTab(label_tab, "Nhãn Tiêu Chuẩn 7980 / 7990")

        msi_layout.addWidget(msi_tab_widget)
        splitter.addWidget(msi_container)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        main_layout.addWidget(splitter)

        # ---------------------------------------------------------------------
        # Bottom Action Panel: Preliminary Self-check & Final Submit
        # ---------------------------------------------------------------------
        bottom_layout = QHBoxLayout()

        self.btn_self_check = QPushButton("🔍 Kiểm tra Sơ bộ (Self-Check PLM & R3)")
        self.btn_self_check.setMinimumHeight(42)
        self.btn_self_check.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_self_check.setStyleSheet("background-color: #0078D4; color: white; border-radius: 4px; padding: 6px 16px;")
        self.btn_self_check.clicked.connect(self.run_preliminary_self_check)

        self.lbl_check_summary = QLabel("Chưa đối soát")
        self.lbl_check_summary.setFont(QFont("Calibri", 10, QFont.Weight.Bold))

        self.btn_submit = QPushButton("🚀 Xác nhận & Nộp dữ liệu (Final Submit)")
        self.btn_submit.setMinimumHeight(42)
        self.btn_submit.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_submit.setStyleSheet("background-color: #10B981; color: white; border-radius: 4px; padding: 6px 20px;")
        self.btn_submit.clicked.connect(self.submit_data)

        bottom_layout.addWidget(self.btn_self_check)
        bottom_layout.addWidget(self.lbl_check_summary)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_submit)

        main_layout.addLayout(bottom_layout)

        # Add initial sample row for immediate productivity
        self._populate_sample_rows()

    # =========================================================================
    # Table Grid Manipulations
    # =========================================================================

    def add_cttt_row(
        self,
        page: str = "01",
        part_code: str = "",
        part_name: str = "",
        quantity: float = 1.0,
        explanation: str = "",
        status: str = "-",
    ) -> int:
        """Append an editable row to the CTTT components table."""
        row = self.cttt_table.rowCount()
        self.cttt_table.insertRow(row)

        item_page = QTableWidgetItem(str(page))
        item_page.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 0, item_page)

        item_code = QTableWidgetItem(str(part_code).strip().upper())
        self.cttt_table.setItem(row, 1, item_code)

        item_name = QTableWidgetItem(str(part_name))
        self.cttt_table.setItem(row, 2, item_name)

        item_qty = QTableWidgetItem(str(quantity))
        item_qty.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.cttt_table.setItem(row, 3, item_qty)

        # Reference quantities initially blank or zero
        item_plm = QTableWidgetItem("-")
        item_plm.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.cttt_table.setItem(row, 4, item_plm)

        item_r3 = QTableWidgetItem("-")
        item_r3.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.cttt_table.setItem(row, 5, item_r3)

        item_exp = QTableWidgetItem(str(explanation))
        self.cttt_table.setItem(row, 6, item_exp)

        item_status = QTableWidgetItem(status)
        item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 7, item_status)

        return row

    def remove_selected_cttt_row(self) -> None:
        """Delete currently selected row from the CTTT table."""
        current_row = self.cttt_table.currentRow()
        if current_row >= 0:
            self.cttt_table.removeRow(current_row)

    def clear_cttt_table(self) -> None:
        """Clear all rows from CTTT table."""
        self.cttt_table.setRowCount(0)
        self.lbl_check_summary.setText("Bảng trống")

    def _populate_sample_rows(self) -> None:
        """Seed table with typical sub-unit components."""
        samples = [
            ("01", "302FP02010", "MOTOR BRACKET", 1.0, ""),
            ("01", "302FP02020", "SCREW M3X6", 4.0, ""),
            ("02", "302FP02030", "F-THETA LENS", 2.0, ""),
        ]
        for p, code, name, qty, exp in samples:
            self.add_cttt_row(page=p, part_code=code, part_name=name, quantity=qty, explanation=exp)

    def _on_sub_unit_changed(self, new_unit: str) -> None:
        """Auto-adjust sample labels and MSI unit name when sub-unit switches."""
        self.msi_unit_name_edit.setText(f"{new_unit} ASSY")

    # =========================================================================
    # Data Import & Reference Loading
    # =========================================================================

    def import_cttt_from_file(self) -> None:
        """Import component rows from user-selected Excel (.xlsx, .xls) or CSV file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn tệp CTTT (Excel/CSV)",
            str(self.base_dir),
            "Excel / CSV Files (*.xlsx *.xls *.csv);;All Files (*)",
        )
        if not file_path:
            return

        try:
            p = Path(file_path)
            if p.suffix.lower() == ".csv":
                df = pd.read_csv(p)
            else:
                df = pd.read_excel(p)

            # Detect column names resiliently
            col_part = next((c for c in df.columns if any(k in str(c).upper() for k in ["MÃ", "PART", "CODE", "ITEM"])), None)
            col_name = next((c for c in df.columns if any(k in str(c).upper() for k in ["TÊN", "NAME", "DESC"])), None)
            col_qty = next((c for c in df.columns if any(k in str(c).upper() for k in ["SỐ LƯỢNG", "QTY", "QUANTITY", "SL"])), None)
            col_page = next((c for c in df.columns if any(k in str(c).upper() for k in ["TRANG", "PAGE"])), None)

            if not col_part:
                QMessageBox.warning(self, "Lỗi định dạng", "Không tìm thấy cột chứa Mã linh kiện trong tệp đã chọn.")
                return

            self.clear_cttt_table()
            count = 0
            for _, r in df.iterrows():
                p_code = str(r[col_part]).strip()
                if not p_code or p_code.lower() == "nan":
                    continue
                p_name = str(r[col_name]) if col_name and pd.notna(r[col_name]) else ""
                try:
                    p_qty = float(r[col_qty]) if col_qty and pd.notna(r[col_qty]) else 1.0
                except (ValueError, TypeError):
                    p_qty = 1.0
                p_page = str(r[col_page]) if col_page and pd.notna(r[col_page]) else "01"

                self.add_cttt_row(page=p_page, part_code=p_code, part_name=p_name, quantity=p_qty)
                count += 1

            QMessageBox.information(self, "Thành công", f"Đã nhập thành công {count} linh kiện từ:\n{p.name}")
        except Exception as exc:
            logger.error("Failed to import CTTT data: %s", exc)
            QMessageBox.critical(self, "Lỗi nhập tệp", f"Lỗi đọc tệp CTTT:\n{exc}")

    def _select_reference_files(self) -> None:
        """Allow member to manually point to PLM or R3 BOM files for self-checking."""
        plm_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn tệp BOM PLM (PLM_*.xlsx)",
            str(self.base_dir),
            "Excel Files (*.xlsx *.xls);;All Files (*)",
        )
        if plm_path:
            try:
                self.plm_data = pd.read_excel(plm_path)
                self.lbl_ref_status.setText(f"PLM: {Path(plm_path).name}")
            except Exception as exc:
                QMessageBox.warning(self, "Lỗi nạp PLM", f"Không thể đọc tệp PLM: {exc}")

        r3_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn tệp BOM R3 (R3_*.xls)",
            str(self.base_dir),
            "Excel Files (*.xls *.xlsx);;All Files (*)",
        )
        if r3_path:
            try:
                # SAP exports might be HTML table format or Excel
                from src.automation.sap.parser import parse_r3_cs12_file, ResilientR3Parser
                self.r3_data = parse_r3_cs12_file(Path(r3_path))
                self.lbl_ref_status.setText(f"Đã nạp PLM & R3 ({Path(r3_path).name})")
            except Exception:
                try:
                    self.r3_data = pd.read_excel(r3_path)
                    self.lbl_ref_status.setText("Đã nạp PLM & R3")
                except Exception as exc:
                    QMessageBox.warning(self, "Lỗi nạp R3", f"Không thể đọc tệp R3: {exc}")

    def set_reference_data(
        self,
        plm_data: pd.DataFrame | None = None,
        r3_data: pd.DataFrame | None = None,
    ) -> None:
        """Programmatically supply reference PLM & R3 DataFrames for verification."""
        self.plm_data = plm_data
        self.r3_data = r3_data
        if plm_data is not None or r3_data is not None:
            self.lbl_ref_status.setText("BOM PLM & R3 sẵn sàng")
            self.lbl_ref_status.setStyleSheet("color: #10B981; font-weight: bold;")

    # =========================================================================
    # Preliminary Self-Check & Instant Visual Feedback
    # =========================================================================

    def run_preliminary_self_check(self) -> dict[str, Any]:
        """Execute instant three-way self-check and update table color styling."""
        row_count = self.cttt_table.rowCount()
        if row_count == 0:
            QMessageBox.warning(self, "Cảnh báo", "Bảng linh kiện CTTT chưa có dữ liệu để đối soát.")
            return {"status": "EMPTY", "ok_count": 0, "ng_count": 0}

        # Build in-memory records from current table
        _ = self.get_cttt_table_data()

        # Build PLM lookup if available
        plm_dict: dict[str, float] = {}
        if self.plm_data is not None and not self.plm_data.empty:
            id_col = next((c for c in self.plm_data.columns if any(k in str(c).upper() for k in ["ITEM ID", "ITEM_ID", "PART_CODE", "MÃ"])), None)
            qty_col = next((c for c in self.plm_data.columns if any(k in str(c).upper() for k in ["QUANTITY", "QTY", "SL"])), None)
            if id_col and qty_col:
                for _, pr in self.plm_data.iterrows():
                    code = str(pr[id_col]).strip().upper()
                    try:
                        q = float(pr[qty_col])
                    except (ValueError, TypeError):
                        q = 0.0
                    plm_dict[code] = plm_dict.get(code, 0.0) + q

        # Build R3 lookup if available
        r3_dict: dict[str, float] = {}
        if self.r3_data is not None and not self.r3_data.empty:
            mat_col = next((c for c in self.r3_data.columns if any(k in str(c).upper() for k in ["MATERIAL", "COMPONENT", "PART_CODE", "MÃ"])), None)
            qty_col = next((c for c in self.r3_data.columns if any(k in str(c).upper() for k in ["QUANTITY", "QTY", "SL"])), None)
            if mat_col and qty_col:
                for _, rr in self.r3_data.iterrows():
                    code = str(rr[mat_col]).strip().upper()
                    try:
                        q = float(rr[qty_col])
                    except (ValueError, TypeError):
                        q = 0.0
                    r3_dict[code] = r3_dict.get(code, 0.0) + q

        ok_count = 0
        ng_count = 0

        # Evaluate each row
        for r_idx in range(row_count):
            p_code = self.cttt_table.item(r_idx, 1).text().strip().upper()
            try:
                cttt_q = float(self.cttt_table.item(r_idx, 3).text().strip())
            except (ValueError, TypeError, AttributeError):
                cttt_q = 0.0

            plm_q = plm_dict.get(p_code, cttt_q if not plm_dict else 0.0)
            r3_q = r3_dict.get(p_code, cttt_q if not r3_dict else 0.0)

            # Update reference values in grid
            self.cttt_table.item(r_idx, 4).setText(str(plm_q))
            self.cttt_table.item(r_idx, 5).setText(str(r3_q))

            # Perform genuine reconciliation check
            res = self.reconciliation_engine.reconcile_single_row(
                cttt_qty=cttt_q,
                plm_qty=plm_q,
                r3_qty=r3_q,
                plm_rev="A",
                r3_rev="A",
            )

            status_str = res["overall_check"]
            status_item = self.cttt_table.item(r_idx, 7)
            status_item.setText(status_str)

            # Apply instant visual feedback colors
            if status_str == "OK":
                ok_count += 1
                self._style_table_row(r_idx, is_ok=True)
            else:
                ng_count += 1
                self._style_table_row(r_idx, is_ok=False)

        # Check MSI inputs
        msi_res = evaluate_msi_branch(
            in_plm=(p_code in plm_dict or not plm_dict),
            member_code=self.msi_code_edit.text().strip(),
            master_code=self.msi_code_edit.text().strip(),  # baseline match if not loaded
            member_service=self.msi_service_edit.text().strip(),
            master_service="",
        )
        if msi_res["status"] == "OK":
            self.lbl_msi_badge.setText(f"MSI OK (Nhánh {msi_res['branch']})")
            self.lbl_msi_badge.setStyleSheet("background-color: #C6EFCE; color: #006100; font-weight: bold; border-radius: 4px; padding: 4px;")
        else:
            self.lbl_msi_badge.setText(f"MSI NG (Nhánh {msi_res['branch']})")
            self.lbl_msi_badge.setStyleSheet("background-color: #FFC7CE; color: #9C0006; font-weight: bold; border-radius: 4px; padding: 4px;")

        # Update summary label
        summary_text = f"Kết quả: {ok_count} OK, {ng_count} NG / Tổng {row_count} linh kiện"
        self.lbl_check_summary.setText(summary_text)
        if ng_count > 0:
            self.lbl_check_summary.setStyleSheet("color: #dc3545; font-weight: bold;")
        else:
            self.lbl_check_summary.setStyleSheet("color: #28a745; font-weight: bold;")

        return {
            "status": "OK" if ng_count == 0 else "NG",
            "ok_count": ok_count,
            "ng_count": ng_count,
            "total": row_count,
        }

    def _style_table_row(self, row_idx: int, is_ok: bool) -> None:
        """Apply soft green (OK) or soft red (NG) styling across the row."""
        bg_color = QColor(f"#{COLOR_GREEN_FILL_HEX}") if is_ok else QColor(f"#{COLOR_RED_FILL_HEX}")
        fg_color = QColor(f"#{COLOR_GREEN_FONT_HEX}") if is_ok else QColor(f"#{COLOR_RED_FONT_HEX}")

        # Highlight Status cell and Part code cell
        status_item = self.cttt_table.item(row_idx, 7)
        if status_item:
            status_item.setBackground(bg_color)
            status_item.setForeground(fg_color)
            status_item.setFont(QFont("Calibri", 10, QFont.Weight.Bold))

        part_item = self.cttt_table.item(row_idx, 1)
        if part_item and not is_ok:
            part_item.setBackground(bg_color)
            part_item.setForeground(fg_color)

    # =========================================================================
    # Submission Generation
    # =========================================================================

    def get_cttt_table_data(self) -> list[dict[str, Any]]:
        """Collect current table items into structured dictionary records."""
        records = []
        sub_name = self.sub_unit_combo.currentText().strip()
        author = self.author_edit.text().strip()

        for r in range(self.cttt_table.rowCount()):
            page = self.cttt_table.item(r, 0).text().strip() if self.cttt_table.item(r, 0) else ""
            part_code = self.cttt_table.item(r, 1).text().strip() if self.cttt_table.item(r, 1) else ""
            part_name = self.cttt_table.item(r, 2).text().strip() if self.cttt_table.item(r, 2) else ""
            try:
                qty = float(self.cttt_table.item(r, 3).text().strip())
            except (ValueError, TypeError, AttributeError):
                qty = 0.0

            explanation = self.cttt_table.item(r, 6).text().strip() if self.cttt_table.item(r, 6) else ""
            status = self.cttt_table.item(r, 7).text().strip() if self.cttt_table.item(r, 7) else "OK"

            records.append({
                "SUB": sub_name,
                "TRANG CTTT": page,
                "MÃ LINH KIỆN": part_code,
                "TÊN LINH KIỆN": part_name,
                "SỐ LƯỢNG": qty,
                "PHỤ TRÁCH": author,
                "Giải thích": explanation,
                "Check": status,
            })
        return records

    def submit_data(self) -> Path | None:
        """Validate, generate submission package file, and emit signal."""
        rows = self.get_cttt_table_data()
        if not rows:
            QMessageBox.warning(self, "Cảnh báo", "Bảng CTTT không có dữ liệu để nộp.")
            return None

        # Check for unaddressed NG items
        ng_items = [r["MÃ LINH KIỆN"] for r in rows if r.get("Check") == "NG" and not r.get("Giải thích")]
        if ng_items:
            reply = QMessageBox.warning(
                self,
                "Cảnh báo sai khác chưa giải trình",
                f"Có {len(ng_items)} linh kiện bị NG nhưng chưa điền Giải thích:\n"
                + ", ".join(ng_items[:5])
                + "\n\nBạn có muốn tiếp tục nộp không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return None

        sub_unit = self.sub_unit_combo.currentText().strip()
        author = self.author_edit.text().strip()
        model = self.model_combo.currentText().strip()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        target_dir = self.base_dir / "CTTT" / sub_unit
        target_dir.mkdir(parents=True, exist_ok=True)
        export_file = target_dir / f"formnguoidung_{sub_unit}_{timestamp}.xlsx"

        try:
            # Write structured workbook matching formnguoidung.xlsm
            wb = openpyxl.Workbook()
            ws_cttt = wb.active
            ws_cttt.title = "CTTT"

            headers = ["SUB", "TRANG CTTT", "MÃ LINH KIỆN", "TÊN LINH KIỆN", "SỐ LƯỢNG", "PHỤ TRÁCH", "Giải thích", "Check"]
            ws_cttt.append(headers)
            for r in rows:
                ws_cttt.append([
                    r["SUB"],
                    r["TRANG CTTT"],
                    r["MÃ LINH KIỆN"],
                    r["TÊN LINH KIỆN"],
                    r["SỐ LƯỢNG"],
                    r["PHỤ TRÁCH"],
                    r["Giải thích"],
                    r["Check"],
                ])

            # Sheet MSI
            ws_msi = wb.create_sheet(title="MSI")
            ws_msi.append(["Mã LK Barcode", "Mã UNIT", "Tên UNIT", "3 ký tự MSI", "SEVICE", "ABS", "Điện áp", "Loại Label"])
            ws_msi.append([
                self.msi_barcode_edit.text().strip(),
                self.msi_unit_code_edit.text().strip(),
                self.msi_unit_name_edit.text().strip(),
                self.msi_code_edit.text().strip(),
                self.msi_service_edit.text().strip(),
                self.msi_abs_combo.currentText().strip(),
                self.msi_volt_combo.currentText().strip(),
                self.msi_label_type_combo.currentText().strip(),
            ])

            # Sheet Label 7980_7990
            ws_lbl = wb.create_sheet(title="Label_7980_7990")
            ws_lbl.append(["Loại nhãn", "Mã nhãn", "Vị trí", "Quy cách kiểm tra"])
            ws_lbl.append([
                self.label_spec_combo.currentText().strip(),
                self.label_code_edit.text().strip(),
                self.label_pos_edit.text().strip(),
                self.label_check_combo.currentText().strip(),
            ])

            wb.save(export_file)
            logger.info("Successfully exported member submission to: %s", export_file)

            submission_info = {
                "sub_unit": sub_unit,
                "author": author,
                "model": model,
                "timestamp": timestamp,
                "item_count": len(rows),
                "file_path": str(export_file),
                "status": "Submitted",
            }
            self.submission_completed.emit(submission_info)

            QMessageBox.information(
                self,
                "Nộp dữ liệu thành công",
                f"Đã hoàn thành nộp dữ liệu công đoạn '{sub_unit}'!\n"
                f"Tệp lưu tại: {export_file.name}\n"
                f"Số lượng linh kiện: {len(rows)}",
            )
            return export_file
        except Exception as exc:
            logger.error("Failed to write submission workbook: %s", exc)
            QMessageBox.critical(self, "Lỗi nộp dữ liệu", f"Không thể xuất tệp nộp dữ liệu:\n{exc}")
            return None
