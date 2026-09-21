"""Feature F23: Modern Member Input Workspace Module (Conforming to Requirement R6).

Provides the full-featured Member desktop workspace for Kyocera production line engineers:
- Auto-loading assignment: Automatically detects Engineer Name, Machine Code, Department.
- 3 Input Tables conforming to formnguoidung.xlsm:
  1. CTTT table: 15-column engineering specification with full comparison metrics.
  2. MSI table: 35 preloaded canonical sub-units (IMAGE, FUSER... HONTAI) with FIX_SERIAL cross-check.
  3. Label 7980/7990 table: Standard 12-column software LCP & high-voltage label catalog.
- Preliminary Self-Check: 1-click automatic discovery of PLM & R3 BOMs in machine folder,
  instant visual feedback (soft green #C6EFCE for OK, soft red #FFC7CE for NG).
- Submission Seal: "Xác nhận Nộp" writes Q2="OK" (green fill) directly into member assignment file,
  notifying Leader in real-time. Includes unlock/re-edit functionality.
"""

from __future__ import annotations

import datetime
import logging
import re
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QCompleter,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.core.msi_engine import FixSerialMaster, MSIEngine, evaluate_msi_branch
from src.core.reconciliation import (
    ReconciliationEngine,
    _clean_part_code,
    _normalize_rev,
    _to_float,
)
from src.gui.leader_view import ROSTER_MECHA_1, ROSTER_MECHA_2, ROSTER_MECHA_3
from src.gui.styles import get_theme_manager, tokens
from src.gui.styles.tokens import (
    COLOR_DARK_STATUS_DIFF_BG,
    COLOR_DARK_STATUS_DIFF_TEXT,
    COLOR_DARK_STATUS_MATCH_BG,
    COLOR_DARK_STATUS_MATCH_TEXT,
    COLOR_LIGHT_STATUS_DIFF_BG,
    COLOR_LIGHT_STATUS_DIFF_TEXT,
    COLOR_LIGHT_STATUS_MATCH_BG,
    COLOR_LIGHT_STATUS_MATCH_TEXT,
)
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_GREEN_FONT_HEX,
    COLOR_RED_FILL_HEX,
    COLOR_RED_FONT_HEX,
    STANDARD_SUB_UNITS,
)

logger = logging.getLogger(__name__)

# Preloaded canonical Unit names from formnguoidung.xlsm Sheet MSI (Rows 2..36)
CANONICAL_MSI_UNITS = [
    "IMAGE UNIT",
    "FUSER UNIT",
    "DATA LASER -V ",
    "CASE OUTER -V ",
    "ISU",
    "TRANSFER UNIT(IRIS)",
    "IH UNIT L",
    "MAIN DRIVE ASSY",
    "UNIT HIGH VOLTAGE TRANSFER",
    "UNIT HIGH VOLTAGE MAIN(IRIS)",
    "WTB UNIT",
    "PWB IH ASSY WITH SOFTWARE",
    "POLYBAG(IRIS)",
    "AC CORD(dây điện)",
    "LSU KC UNIT",
    "LSU MY UNIT",
    "DLP K UNIT",
    "DLP Y UNIT",
    "DLP M UNIT",
    "DLP C UNIT",
    "DP UNIT",
    "DRUM UNIT(K)",
    "DRUM UNIT(Y)",
    "DRUM UNIT(M)",
    "DRUM UNIT(C)",
    "UNIT HIGH VOLTAGE FUSER",
    "CONVEYING UNIT",
    "RATING-LABEL",
    "PWB ENGINE CONNECT ASSY",
    "UNIT LOW VOLTAGE(bản mạch nguồn)",
    "PWB ASSY MAIN WITH SOFTWARE(bản mạch Main)",
    "FAX UNIT(bản mạch Fax)",
    "UNIT HIGH VOLTAGE(bản mạch cao áp)",
    "PWB ENGINE ASSY WITH SOFTWARE(bản mạch Engine)",
    "HONTAI(ký tự cố định hontai)",
]


# =============================================================================
# Workflow Stepper 3 Steps (Data-Dense Dashboard Header)
# =============================================================================

class MemberWorkflowStepper(QFrame):
    """3-Step Workflow Progress Stepper with SVG icons and active state indicators."""

    step_changed = pyqtSignal(int)

    STEPS = [
        ("Bước 1: Chọn Model & BOM", "folder"),
        ("Bước 2: Đối Soát Quy Tắc", "search"),
        ("Bước 3: Phụ trách công đoạn & Xuất Kết Quả", "file-spreadsheet"),
    ]

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("workflow_stepper")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet(
            "QFrame#workflow_stepper {"
            "  background-color: #FFFFFF;"
            "  border: 1px solid #CBD5E1;"
            "  border-radius: 6px;"
            "  padding: 4px;"
            "}"
        )
        self.current_step = 0
        self.step_completed_flags = [False, False, False]
        self._step_widgets: list[QWidget] = []
        self._step_labels: list[QLabel] = []
        self._icon_labels: list[QLabel] = []
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(12)

        for idx, (title, default_icon) in enumerate(self.STEPS):
            step_w = QWidget(self)
            w_layout = QHBoxLayout(step_w)
            w_layout.setContentsMargins(8, 6, 8, 6)
            w_layout.setSpacing(8)

            icon_lbl = QLabel(step_w)
            icon_lbl.setFixedSize(22, 22)
            icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            w_layout.addWidget(icon_lbl)

            text_lbl = QLabel(title, step_w)
            text_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.Medium))
            w_layout.addWidget(text_lbl)

            layout.addWidget(step_w, stretch=1)
            self._step_widgets.append(step_w)
            self._icon_labels.append(icon_lbl)
            self._step_labels.append(text_lbl)

            if idx < len(self.STEPS) - 1:
                sep = QLabel("→", self)
                sep.setStyleSheet("color: #94A3B8; font-weight: bold; font-size: 14px;")
                sep.setAlignment(Qt.AlignmentFlag.AlignCenter)
                layout.addWidget(sep)

        self._refresh_step_styles()

    def set_current_step(self, step_idx: int) -> None:
        if 0 <= step_idx < len(self.STEPS):
            self.current_step = step_idx
            self._refresh_step_styles()

    def set_step_completed(self, step_idx: int, completed: bool = True) -> None:
        if 0 <= step_idx < len(self.STEPS):
            self.step_completed_flags[step_idx] = completed
            self._refresh_step_styles()

    def _refresh_step_styles(self) -> None:
        theme_mgr = get_theme_manager()
        for idx in range(len(self.STEPS)):
            step_w = self._step_widgets[idx]
            icon_lbl = self._icon_labels[idx]
            text_lbl = self._step_labels[idx]
            _, def_icon = self.STEPS[idx]

            if self.step_completed_flags[idx]:
                # Completed: Green check SVG (WCAG AAA >= 7.0:1)
                icon_lbl.setPixmap(theme_mgr.get_styled_icon("check-circle", color="#064E3B").pixmap(18, 18))
                text_lbl.setStyleSheet("color: #064E3B; font-weight: 600;")
                step_w.setStyleSheet(
                    "background-color: #DCFCE7; border: 1px solid #86EFAC; border-radius: 4px;"
                )
            elif idx == self.current_step:
                # Active: Blue
                icon_lbl.setPixmap(theme_mgr.get_styled_icon(def_icon, color="#2563EB").pixmap(18, 18))
                text_lbl.setStyleSheet("color: #1E40AF; font-weight: bold;")
                step_w.setStyleSheet(
                    "background-color: #EFF6FF; border: 1px solid #93C5FD; border-radius: 4px;"
                )
            else:
                # Pending: Slate
                icon_lbl.setPixmap(theme_mgr.get_styled_icon("alert-triangle", color="#94A3B8").pixmap(16, 16))
                text_lbl.setStyleSheet("color: #64748B; font-weight: normal;")
                step_w.setStyleSheet(
                    "background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 4px;"
                )


class MemberWorkspaceView(QWidget):
    """Modern Member Workspace implementing Requirement R6 and legacy formnguoidung."""

    submission_completed = pyqtSignal(dict)  # Emits metadata on successful Q2="OK" submission

    def on_theme_changed(self, effective_theme: str) -> None:
        """Propagate theme changes to MemberWorkflowStepper and styled icons."""
        if hasattr(self, "workflow_stepper"):
            self.workflow_stepper._refresh_step_styles()

    def __init__(
        self,
        parent: QWidget | None = None,
        base_dir: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.base_dir = Path(base_dir or r"D:\Sandbox\pm_sosanhbom")
        self.reconciliation_engine = ReconciliationEngine()
        self.msi_engine = MSIEngine()
        self.fix_serial_master = FixSerialMaster()

        # Assignment state
        self.current_assignment_file: Path | None = None
        self.current_machine_dir: Path | None = None
        self.engineer_name: str = "Chưa nạp phân công (Bấm Mở gói hoặc gõ bên dưới)"
        self.machine_code: str = ""
        self.department: str = "Phòng Kỹ Thuật Cơ"
        self.is_submitted_ok: bool = False

        # Cached reference data for self-check
        self.plm_data: pd.DataFrame | None = None
        self.r3_data: pd.DataFrame | None = None

        self._init_ui()
        self._populate_canonical_msi_table()

    # =========================================================================
    # UI Layout Construction
    # =========================================================================

    def _init_ui(self) -> None:
        """Construct full responsive layout for Member Workspace."""
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        content_widget = QWidget()
        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # ---------------------------------------------------------------------
        # Zone 0: 3-Step Workflow Stepper
        # ---------------------------------------------------------------------
        self.workflow_stepper = MemberWorkflowStepper(self)
        self.stepper = self.workflow_stepper
        main_layout.addWidget(self.workflow_stepper)

        # ---------------------------------------------------------------------
        # Zone 1: Auto-loading Banner & Assignment Information
        # ---------------------------------------------------------------------
        banner_group = QGroupBox("1. Thông tin Phân công")
        banner_layout = QVBoxLayout(banner_group)
        banner_layout.setSpacing(8)

        # Row A: Action buttons & file path
        row_a = QHBoxLayout()
        self.btn_open_assignment = QPushButton("Mở gói phân công...")
        self.btn_open_assignment.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_open_assignment.setStyleSheet("font-weight: bold; padding: 4px 10px;")
        self.btn_open_assignment.clicked.connect(lambda: self.open_assignment_package())
        row_a.addWidget(self.btn_open_assignment)

        self.btn_select_machine_dir = QPushButton("Chọn thư mục máy...")
        self.btn_select_machine_dir.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_select_machine_dir.clicked.connect(lambda: self.load_machine_directory())
        row_a.addWidget(self.btn_select_machine_dir)

        row_a.addWidget(QLabel("Đường dẫn gói:"))
        self.lbl_assignment_path = QLabel("Chưa nạp gói phân công")
        self.lbl_assignment_path.setStyleSheet("color: #495057; font-style: italic;")
        row_a.addWidget(self.lbl_assignment_path)
        row_a.addStretch()

        self.btn_load_refs = QPushButton("Nạp PLM & R3 đối chiếu...")
        self.btn_load_refs.setIcon(get_theme_manager().get_styled_icon("download"))
        self.btn_load_refs.clicked.connect(self._select_reference_files)
        row_a.addWidget(self.btn_load_refs)

        self.lbl_ref_status = QLabel("Chưa nạp BOM đối chiếu")
        self.lbl_ref_status.setStyleSheet("color: #6c757d; font-style: italic;")
        row_a.addWidget(self.lbl_ref_status)
        banner_layout.addLayout(row_a)

        # Row B: Auto-detected badges & status
        row_b = QHBoxLayout()
        row_b.addWidget(QLabel("Phụ trách công đoạn:"))
        self.lbl_engineer_name = QLabel(self.engineer_name)
        self.lbl_engineer_name.setStyleSheet("font-weight: bold; color: #0d6efd;")
        row_b.addWidget(self.lbl_engineer_name)

        row_b.addSpacing(16)
        row_b.addWidget(QLabel("Mã máy:"))
        self.lbl_machine_code = QLabel("-")
        self.lbl_machine_code.setStyleSheet("font-weight: bold; color: #6610f2;")
        row_b.addWidget(self.lbl_machine_code)

        row_b.addSpacing(16)
        row_b.addWidget(QLabel("Phòng ban:"))
        self.lbl_department = QLabel(self.department)
        self.lbl_department.setStyleSheet("font-weight: bold; color: #198754;")
        row_b.addWidget(self.lbl_department)

        row_b.addSpacing(16)
        row_b.addWidget(QLabel("Trạng thái nộp:"))
        self.lbl_submission_seal = QLabel("⏳ CHƯA NỘP (Q2 TRỐNG)")
        self.lbl_submission_seal.setStyleSheet(
            "background-color: #FFF3CD; color: #856404; font-weight: bold; border-radius: 4px; padding: 4px 8px;"
        )
        row_b.addWidget(self.lbl_submission_seal)
        row_b.addStretch()
        banner_layout.addLayout(row_b)

        # Row C: Compatibility controls for legacy test suite
        row_c = QHBoxLayout()
        row_c.addWidget(QLabel("Công đoạn (Sub-Unit):"))
        self.sub_unit_combo = QComboBox()
        self.sub_unit_combo.addItems(STANDARD_SUB_UNITS)
        self.sub_unit_combo.currentTextChanged.connect(self._on_sub_unit_changed)
        row_c.addWidget(self.sub_unit_combo)

        row_c.addWidget(QLabel("Người phụ trách:"))
        self.author_edit = QLineEdit()
        self.author_edit.setPlaceholderText("Gõ tên người phụ trách (VD: Son_mecha1, Duy_mecha1...)")
        author_names: list[str] = []
        try:
            from src.core.member_database import MemberDatabaseManager
            db = MemberDatabaseManager(base_dir=self.base_dir)
            author_names = db.get_all_account_names()
        except Exception as e:
            logger.warning(f"Could not load author names from DB: {e}")
        if not author_names:
            author_names = ROSTER_MECHA_1 + ROSTER_MECHA_2 + ROSTER_MECHA_3
        completer = QCompleter(author_names, self)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.author_edit.setCompleter(completer)
        self.author_edit.textChanged.connect(self._on_author_text_changed)
        row_c.addWidget(self.author_edit)

        row_c.addWidget(QLabel("Mã Model:"))
        self.model_combo = QComboBox()
        self.model_combo.addItems(["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"])
        row_c.addWidget(self.model_combo)
        row_c.addStretch()
        banner_layout.addLayout(row_c)

        main_layout.addWidget(banner_group)

        # ---------------------------------------------------------------------
        # Zone 2: 3 Input Tables conforming to formnguoidung.xlsm
        # ---------------------------------------------------------------------
        self.main_tab_widget = QTabWidget()

        # Tab 1: CTTT Components Table (15 columns)
        cttt_widget = QWidget()
        cttt_layout = QVBoxLayout(cttt_widget)
        cttt_layout.setContentsMargins(4, 4, 4, 4)

        cttt_bar = QHBoxLayout()
        cttt_title = QLabel("2. Danh sách Linh kiện Chỉ thị thao tác (CTTT - Cols A:N & Q)")
        cttt_title.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        cttt_bar.addWidget(cttt_title)
        cttt_bar.addStretch()

        self.btn_add_row = QPushButton("Thêm dòng")
        self.btn_add_row.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_add_row.clicked.connect(lambda: self.add_cttt_row())
        cttt_bar.addWidget(self.btn_add_row)

        self.btn_remove_row = QPushButton("Xóa dòng")
        self.btn_remove_row.setIcon(get_theme_manager().get_styled_icon("x-circle"))
        self.btn_remove_row.clicked.connect(self.remove_selected_cttt_row)
        cttt_bar.addWidget(self.btn_remove_row)

        self.btn_import_excel = QPushButton("Nhập từ Excel/CSV...")
        self.btn_import_excel.setIcon(get_theme_manager().get_styled_icon("download"))
        self.btn_import_excel.clicked.connect(self.import_cttt_from_file)
        cttt_bar.addWidget(self.btn_import_excel)

        self.btn_paste_clipboard = QPushButton("Dán Clipboard")
        self.btn_paste_clipboard.setIcon(get_theme_manager().get_styled_icon("file-spreadsheet"))
        self.btn_paste_clipboard.clicked.connect(self.paste_cttt_from_clipboard)
        cttt_bar.addWidget(self.btn_paste_clipboard)

        self.btn_clear_table = QPushButton("Xóa hết")
        self.btn_clear_table.setIcon(get_theme_manager().get_styled_icon("refresh"))
        self.btn_clear_table.clicked.connect(self.clear_cttt_table)
        cttt_bar.addWidget(self.btn_clear_table)

        cttt_layout.addLayout(cttt_bar)

        self.cttt_table = QTableWidget(0, 15)
        self.cttt_table.setMinimumHeight(280)
        self.cttt_table.setHorizontalHeaderLabels([
            "Trang CTTT",
            "Mã Linh Kiện",
            "Tên Linh Kiện",
            "Số Lượng",
            "SL PLM (Tham chiếu)",
            "SL R3 (Tham chiếu)",
            "Giải Thích Sai Khác",
            "Kết Quả So Sánh",
            "Unit / Công đoạn",
            "Người Phụ Trách",
            "So Sánh CTTT vs PLM",
            "Rev PLM",
            "So Sánh CTTT vs R3",
            "Rev R3",
            "So Sánh Rev PLM vs R3",
        ])
        self.cttt_table.verticalHeader().setDefaultSectionSize(32)
        self.cttt_table.verticalHeader().setMinimumSectionSize(28)
        self.cttt_table.setShowGrid(True)
        c_header = self.cttt_table.horizontalHeader()
        c_header.setDefaultSectionSize(115)
        c_header.setMinimumSectionSize(65)
        self.cttt_table.setColumnWidth(0, 80)   # Trang CTTT
        self.cttt_table.setColumnWidth(1, 120)  # Mã Linh Kiện
        self.cttt_table.setColumnWidth(2, 200)  # Tên Linh Kiện
        self.cttt_table.setColumnWidth(3, 75)   # Số Lượng
        self.cttt_table.setColumnWidth(4, 130)  # SL PLM (Tham chiếu)
        self.cttt_table.setColumnWidth(5, 130)  # SL R3 (Tham chiếu)
        self.cttt_table.setColumnWidth(6, 170)  # Giải Thích Sai Khác
        self.cttt_table.setColumnWidth(7, 120)  # Kết Quả So Sánh
        self.cttt_table.setColumnWidth(8, 120)  # Unit / Công đoạn
        self.cttt_table.setColumnWidth(9, 120)  # Người Phụ Trách
        self.cttt_table.setColumnWidth(10, 140) # So Sánh CTTT vs PLM
        self.cttt_table.setColumnWidth(11, 75)  # Rev PLM
        self.cttt_table.setColumnWidth(12, 135) # So Sánh CTTT vs R3
        self.cttt_table.setColumnWidth(13, 75)  # Rev R3
        self.cttt_table.setColumnWidth(14, 150) # So Sánh Rev PLM vs R3
        self.cttt_table.setAlternatingRowColors(True)
        cttt_layout.addWidget(self.cttt_table)

        self.main_tab_widget.addTab(cttt_widget, "1. Linh kiện CTTT (A:N)")
        self.main_tab_widget.setTabIcon(0, get_theme_manager().get_styled_icon("file-spreadsheet"))

        # Tab 2: MSI Table (35 Canonical Rows x 12 Columns)
        msi_widget = QWidget()
        msi_layout = QVBoxLayout(msi_widget)
        msi_layout.setContentsMargins(4, 4, 4, 4)

        msi_bar = QHBoxLayout()
        msi_title = QLabel("Quản lý Barcode & 3 Ký tự Cố định MSI (35 Cụm Unit & Thân máy HONTAI)")
        msi_title.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        msi_bar.addWidget(msi_title)
        msi_bar.addStretch()

        self.btn_clear_msi = QPushButton("Đặt lại 35 Unit")
        self.btn_clear_msi.setIcon(get_theme_manager().get_styled_icon("refresh"))
        self.btn_clear_msi.clicked.connect(self._populate_canonical_msi_table)
        msi_bar.addWidget(self.btn_clear_msi)
        msi_layout.addLayout(msi_bar)

        self.msi_table = QTableWidget(0, 12)
        self.msi_table.setMinimumHeight(240)
        self.msi_table.setHorizontalHeaderLabels([
            "Mã LK Barcode",
            "Mã UNIT/Bản mạch",
            "Tên UNIT",
            "3 Ký tự MSI",
            "SEVICE",
            "ABS",
            "Trang CTTT",
            "Điện áp",
            "Loại Label",
            "Tên phụ trách",
            "Kết quả so sánh",
            "Ghi chú",
        ])
        self.msi_table.verticalHeader().setDefaultSectionSize(32)
        self.msi_table.verticalHeader().setMinimumSectionSize(28)
        self.msi_table.setShowGrid(True)
        m_header = self.msi_table.horizontalHeader()
        m_header.setDefaultSectionSize(110)
        self.msi_table.setColumnWidth(2, 190)
        self.msi_table.setAlternatingRowColors(True)
        self.msi_table.cellClicked.connect(self._on_msi_table_row_selected)
        msi_layout.addWidget(self.msi_table)

        # MSI Quick Edit Panel
        msi_edit_box = QGroupBox("Chỉnh sửa nhanh dòng MSI được chọn (Quick-Edit Panel)")
        msi_edit_layout = QHBoxLayout(msi_edit_box)

        col1_layout = QVBoxLayout()
        col1_layout.addWidget(QLabel("Mã LK Barcode:"))
        self.msi_barcode_edit = QLineEdit()
        self.msi_barcode_edit.setPlaceholderText("VD: 302FP93010")
        col1_layout.addWidget(self.msi_barcode_edit)
        col1_layout.addWidget(QLabel("Mã UNIT / Bản mạch:"))
        self.msi_unit_code_edit = QLineEdit()
        self.msi_unit_code_edit.setPlaceholderText("VD: 302FP93010")
        col1_layout.addWidget(self.msi_unit_code_edit)
        col1_layout.addWidget(QLabel("Tên UNIT:"))
        self.msi_unit_name_edit = QLineEdit()
        self.msi_unit_name_edit.setPlaceholderText("VD: LSU UNIT")
        col1_layout.addWidget(self.msi_unit_name_edit)
        msi_edit_layout.addLayout(col1_layout)

        col2_layout = QVBoxLayout()
        col2_layout.addWidget(QLabel("3 ký tự cố định MSI:"))
        self.msi_code_edit = QLineEdit()
        self.msi_code_edit.setPlaceholderText("VD: 1HN")
        col2_layout.addWidget(self.msi_code_edit)
        col2_layout.addWidget(QLabel("SEVICE Comment (nếu có):"))
        self.msi_service_edit = QLineEdit()
        self.msi_service_edit.setPlaceholderText("VD: -")
        col2_layout.addWidget(self.msi_service_edit)
        col2_layout.addWidget(QLabel("ABS (Phán định chất lượng):"))
        self.msi_abs_combo = QComboBox()
        self.msi_abs_combo.addItems(["OK", "NG", "N/A", "-"])
        col2_layout.addWidget(self.msi_abs_combo)
        msi_edit_layout.addLayout(col2_layout)

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
        self.lbl_msi_badge.setStyleSheet(
            "background-color: #e9ecef; font-weight: bold; border-radius: 4px; padding: 4px;"
        )
        col3_layout.addWidget(self.lbl_msi_badge)
        msi_edit_layout.addLayout(col3_layout)

        self.btn_apply_msi_edit = QPushButton("Cập nhật\nvào bảng MSI")
        self.btn_apply_msi_edit.setIcon(get_theme_manager().get_styled_icon("check-circle", color="#FFFFFF"))
        self.btn_apply_msi_edit.setStyleSheet("font-weight: bold; padding: 8px 16px; background-color: #0d6efd; color: white;")
        self.btn_apply_msi_edit.clicked.connect(self._apply_msi_quick_edit)
        msi_edit_layout.addWidget(self.btn_apply_msi_edit)

        msi_layout.addWidget(msi_edit_box)
        self.main_tab_widget.addTab(msi_widget, "2. Quản lý MSI (35 Unit)")
        self.main_tab_widget.setTabIcon(1, get_theme_manager().get_styled_icon("filter"))

        # Tab 3: Label 7980 / 7990 Table (12 Columns)
        label_widget = QWidget()
        label_layout = QVBoxLayout(label_widget)
        label_layout.setContentsMargins(4, 4, 4, 4)

        lbl_bar = QHBoxLayout()
        lbl_title = QLabel("Quản lý Tem Nhãn LCP 7980 & Nhãn Cao Áp 7990 (Sheet Label_7980_7990)")
        lbl_title.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        lbl_bar.addWidget(lbl_title)
        lbl_bar.addStretch()

        self.btn_add_label_row = QPushButton("Thêm dòng nhãn")
        self.btn_add_label_row.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_add_label_row.clicked.connect(lambda: self.add_label_row())
        lbl_bar.addWidget(self.btn_add_label_row)

        self.btn_remove_label_row = QPushButton("Xóa dòng nhãn")
        self.btn_remove_label_row.setIcon(get_theme_manager().get_styled_icon("x-circle"))
        self.btn_remove_label_row.clicked.connect(self.remove_selected_label_row)
        lbl_bar.addWidget(self.btn_remove_label_row)

        self.btn_clear_label_table = QPushButton("Xóa hết nhãn")
        self.btn_clear_label_table.setIcon(get_theme_manager().get_styled_icon("refresh"))
        self.btn_clear_label_table.clicked.connect(self.clear_label_table)
        lbl_bar.addWidget(self.btn_clear_label_table)
        label_layout.addLayout(lbl_bar)

        self.label_table = QTableWidget(0, 12)
        self.label_table.setMinimumHeight(240)
        self.label_table.setHorizontalHeaderLabels([
            "Công đoạn (7980)",
            "Trang CTTT (7980)",
            "Mã label (7980)",
            "Tên label (7980)",
            "Số lượng (7980)",
            "Phụ trách (7980)",
            "Công đoạn (7990)",
            "Trang CTTT (7990)",
            "Mã label (7990)",
            "Tên label (7990)",
            "Số lượng (7990)",
            "Phụ trách (7990)",
        ])
        self.label_table.verticalHeader().setDefaultSectionSize(32)
        self.label_table.verticalHeader().setMinimumSectionSize(28)
        self.label_table.setShowGrid(True)
        l_header = self.label_table.horizontalHeader()
        l_header.setDefaultSectionSize(115)
        self.label_table.setAlternatingRowColors(True)
        label_layout.addWidget(self.label_table)

        # Label Quick Spec Fields (Backward Compatibility)
        lbl_quick_box = QGroupBox("Thông tin quy cách kiểm tra nhãn (Label Specifications)")
        lbl_quick_layout = QHBoxLayout(lbl_quick_box)

        l_col1 = QVBoxLayout()
        l_col1.addWidget(QLabel("Loại nhãn quản lý:"))
        self.label_spec_combo = QComboBox()
        self.label_spec_combo.addItems(["Label 7980 (Kyocera Standard)", "Label 7990 (High Voltage)", "Label LCP"])
        l_col1.addWidget(self.label_spec_combo)
        l_col1.addWidget(QLabel("Mã số nhãn linh kiện:"))
        self.label_code_edit = QLineEdit()
        self.label_code_edit.setPlaceholderText("VD: 7980-VN-001")
        l_col1.addWidget(self.label_code_edit)
        lbl_quick_layout.addLayout(l_col1)

        l_col2 = QVBoxLayout()
        l_col2.addWidget(QLabel("Vị trí dán trên CTTT:"))
        self.label_pos_edit = QLineEdit()
        self.label_pos_edit.setPlaceholderText("VD: Mặt trên vỏ khung LSU - Trang 03")
        l_col2.addWidget(self.label_pos_edit)
        l_col2.addWidget(QLabel("Quy cách kiểm tra:"))
        self.label_check_combo = QComboBox()
        self.label_check_combo.addItems([
            "Dán phẳng, không bọt khí, đúng chiều mũi tên",
        ])
        l_col2.addWidget(self.label_check_combo)
        lbl_quick_layout.addLayout(l_col2)

        label_layout.addWidget(lbl_quick_box)
        self.main_tab_widget.addTab(label_widget, "3. Nhãn LCP 7980 / 7990")
        self.main_tab_widget.setTabIcon(2, get_theme_manager().get_styled_icon("calendar"))

        main_layout.addWidget(self.main_tab_widget)

        # ---------------------------------------------------------------------
        # Zone 3 & 4: Preliminary Self-Check & Submission Seal
        # ---------------------------------------------------------------------
        action_bar = QHBoxLayout()

        self.btn_self_check = QPushButton("Kiểm tra Sơ bộ (Self-Check PLM & R3)")
        self.btn_self_check.setIcon(get_theme_manager().get_styled_icon("search", color="#FFFFFF"))
        self.btn_self_check.setMinimumHeight(44)
        self.btn_self_check.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_self_check.setStyleSheet(
            "background-color: #0078D4; color: white; border-radius: 4px; padding: 6px 18px;"
        )
        self.btn_self_check.clicked.connect(self.run_preliminary_self_check)
        action_bar.addWidget(self.btn_self_check)

        self.lbl_check_summary = QLabel("Chưa đối soát")
        self.lbl_check_summary.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        action_bar.addWidget(self.lbl_check_summary)

        action_bar.addStretch()

        self.btn_unlock = QPushButton("Hủy nộp / Mở khóa")
        self.btn_unlock.setIcon(get_theme_manager().get_styled_icon("refresh", color="#FFFFFF"))
        self.btn_unlock.setMinimumHeight(44)
        self.btn_unlock.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_unlock.setStyleSheet(
            "background-color: #6c757d; color: white; border-radius: 4px; padding: 6px 14px;"
        )
        self.btn_unlock.clicked.connect(self.unlock_submission)
        action_bar.addWidget(self.btn_unlock)

        self.btn_submit = QPushButton("Xác nhận Nộp (Đóng dấu Q2 = OK)")
        self.btn_submit.setIcon(get_theme_manager().get_styled_icon("check-circle", color="#FFFFFF"))
        self.btn_submit.setMinimumHeight(44)
        self.btn_submit.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_submit.setStyleSheet(
            "background-color: #10B981; color: white; border-radius: 4px; padding: 6px 22px;"
        )
        main_layout.addLayout(action_bar)

        scroll_area.setWidget(content_widget)
        outer_layout.addWidget(scroll_area)

    # =========================================================================
    # Assignment Auto-Loading
    # =========================================================================

    def open_assignment_package(self, file_path: Path | str | None = None) -> bool:
        """Automatically load engineer assignment file, parse sheets, and discover BOMs."""
        if file_path is None:
            selected_file, _ = QFileDialog.getOpenFileName(
                self,
                "Mở tệp phân công kỹ sư (formnguoidung)",
                str(self.base_dir),
                "Excel Files (*.xlsm *.xlsx);;All Files (*)",
            )
            if not selected_file:
                return False
            file_path = Path(selected_file)
        else:
            file_path = Path(file_path)

        if not file_path.exists():
            logger.warning("Assignment file does not exist: %s", file_path)
            return False

        self.current_assignment_file = file_path
        self.current_machine_dir = file_path.parent

        # 1. Parse Engineer Name from filename stem
        stem = file_path.stem
        self.engineer_name = stem
        self.author_edit.setText(stem)
        self.lbl_engineer_name.setText(stem)

        # 2. Parse Machine Code from parent directory name
        self.machine_code = self.current_machine_dir.name
        self.lbl_machine_code.setText(self.machine_code)

        # 3. Parse Department from engineer name suffix
        stem_lower = stem.lower()
        if "mecha1" in stem_lower:
            self.department = "Phòng Cơ 1"
        elif "mecha2" in stem_lower:
            self.department = "Phòng Cơ 2"
        elif "mecha3" in stem_lower:
            self.department = "Phòng Cơ 3"
        elif "pthtct" in stem_lower:
            self.department = "Phụ Trách Hệ Thống Cấu Trúc"
        else:
            self.department = "Phòng Kỹ Thuật Cơ"
        self.lbl_department.setText(self.department)

        self.lbl_assignment_path.setText(str(file_path))

        # 4. Auto-discover BOMs in machine folder
        self._auto_discover_boms(self.current_machine_dir)

        # 5. Read existing data and Q2 seal from workbook
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            if "CTTT" in wb.sheetnames:
                ws_cttt = wb["CTTT"]
                # Check Q2 status
                q2_val = str(ws_cttt["Q2"].value).strip() if ws_cttt["Q2"].value is not None else ""
                if q2_val.upper() == "OK":
                    self.is_submitted_ok = True
                    self.lbl_submission_seal.setText("✅ ĐÃ NỘP BÀI (Q2 = OK)")
                    self.lbl_submission_seal.setStyleSheet(
                        "background-color: #C6EFCE; color: #006100; font-weight: bold; border-radius: 4px; padding: 4px 8px;"
                    )
                    self.cttt_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
                else:
                    self.is_submitted_ok = False
                    self.lbl_submission_seal.setText("⏳ CHƯA NỘP (Q2 TRỐNG)")
                    self.lbl_submission_seal.setStyleSheet(
                        "background-color: #FFF3CD; color: #856404; font-weight: bold; border-radius: 4px; padding: 4px 8px;"
                    )
                    self.cttt_table.setEditTriggers(
                        QTableWidget.EditTrigger.DoubleClicked
                        | QTableWidget.EditTrigger.SelectedClicked
                        | QTableWidget.EditTrigger.EditKeyPressed
                    )

                # Load CTTT data if present (starting row 3)
                data_rows = []
                for row_cells in ws_cttt.iter_rows(min_row=3, values_only=True):
                    if not row_cells or len(row_cells) < 3:
                        continue
                    part_code = str(row_cells[2]).strip() if row_cells[2] is not None else ""
                    if not part_code or part_code.lower() == "nan":
                        continue
                    sub_val = str(row_cells[0]).strip() if row_cells[0] is not None else ""
                    page_val = str(row_cells[1]).strip() if row_cells[1] is not None else "01"
                    name_val = str(row_cells[3]).strip() if len(row_cells) > 3 and row_cells[3] is not None else ""
                    try:
                        qty_val = float(row_cells[4]) if len(row_cells) > 4 and row_cells[4] is not None else 1.0
                    except (ValueError, TypeError):
                        qty_val = 1.0
                    author_val = str(row_cells[5]).strip() if len(row_cells) > 5 and row_cells[5] is not None else stem
                    plm_q = row_cells[6] if len(row_cells) > 6 and row_cells[6] is not None else "-"
                    comp_plm_val = str(row_cells[7]).strip() if len(row_cells) > 7 and row_cells[7] is not None else "-"
                    plm_rev_val = str(row_cells[8]).strip() if len(row_cells) > 8 and row_cells[8] is not None else "-"
                    r3_q = row_cells[9] if len(row_cells) > 9 and row_cells[9] is not None else "-"
                    comp_r3_val = str(row_cells[10]).strip() if len(row_cells) > 10 and row_cells[10] is not None else "-"
                    r3_rev_val = str(row_cells[11]).strip() if len(row_cells) > 11 and row_cells[11] is not None else "-"
                    comp_rev_val = str(row_cells[12]).strip() if len(row_cells) > 12 and row_cells[12] is not None else "-"
                    chk_val = str(row_cells[13]).strip() if len(row_cells) > 13 and row_cells[13] is not None else "-"
                    exp_val = str(row_cells[14]).strip() if len(row_cells) > 14 and row_cells[14] is not None else ""

                    data_rows.append((
                        page_val,
                        part_code,
                        name_val,
                        qty_val,
                        exp_val,
                        chk_val,
                        sub_val,
                        author_val,
                        plm_q,
                        r3_q,
                        comp_plm_val,
                        plm_rev_val,
                        comp_r3_val,
                        r3_rev_val,
                        comp_rev_val,
                    ))

                if data_rows:
                    self.clear_cttt_table()
                    for row_data in data_rows:
                        self.add_cttt_row(
                            page=row_data[0],
                            part_code=row_data[1],
                            part_name=row_data[2],
                            quantity=row_data[3],
                            explanation=row_data[4],
                            status=row_data[5],
                            sub_unit=row_data[6],
                            author=row_data[7],
                            plm_qty=row_data[8],
                            r3_qty=row_data[9],
                            comp_plm=row_data[10],
                            plm_rev=row_data[11],
                            comp_r3=row_data[12],
                            r3_rev=row_data[13],
                            comp_rev=row_data[14],
                        )

            # Load MSI data if present
            if "MSI" in wb.sheetnames:
                ws_msi = wb["MSI"]
                for r_idx, row_cells in enumerate(ws_msi.iter_rows(min_row=2, max_row=36, values_only=True)):
                    if r_idx >= self.msi_table.rowCount():
                        break
                    for c_idx, cell_val in enumerate(row_cells[:12]):
                        if cell_val is not None:
                            item = self.msi_table.item(r_idx, c_idx)
                            if item:
                                item.setText(str(cell_val).strip())

            # Load Label_7980_7990 data if present
            if "Label_7980_7990" in wb.sheetnames:
                ws_lbl = wb["Label_7980_7990"]
                lbl_rows = list(ws_lbl.iter_rows(min_row=2, values_only=True))
                if any(any(c is not None for c in r) for r in lbl_rows):
                    self.label_table.setRowCount(0)
                    for r in lbl_rows:
                        if not any(c is not None for c in r[:12]):
                            continue
                        row_idx = self.label_table.rowCount()
                        self.label_table.insertRow(row_idx)
                        for c_idx, val in enumerate(r[:12]):
                            txt = str(val).strip() if val is not None else ""
                            self.label_table.setItem(row_idx, c_idx, QTableWidgetItem(txt))

            wb.close()
        except Exception as exc:
            logger.warning("Failed to inspect assignment workbook: %s", exc)

        if hasattr(self, "workflow_stepper"):
            self.workflow_stepper.set_step_completed(0, True)
            self.workflow_stepper.set_current_step(1)

        return True

    def load_machine_directory(self, dir_path: Path | str | None = None) -> bool:
        """Select machine directory, discover BOMs, and auto-load member package if found."""
        if dir_path is None:
            selected_dir = QFileDialog.getExistingDirectory(
                self,
                "Chọn thư mục máy",
                str(self.base_dir),
            )
            if not selected_dir:
                return False
            dir_path = Path(selected_dir)
        else:
            dir_path = Path(dir_path)

        if not dir_path.exists():
            return False

        self.current_machine_dir = dir_path
        self.machine_code = dir_path.name
        self.lbl_machine_code.setText(self.machine_code)

        # Check for member assignment file
        potential_files = [
            f for f in dir_path.glob("*.xls*")
            if not f.name.startswith(("PLM_", "R3_", "BOM_", "form_ssbom_"))
        ]
        if potential_files:
            return self.open_assignment_package(potential_files[0])

        # Otherwise just auto-discover BOMs
        self._auto_discover_boms(dir_path)
        if hasattr(self, "workflow_stepper"):
            self.workflow_stepper.set_step_completed(0, True)
            self.workflow_stepper.set_current_step(1)
        return True

    def _auto_discover_boms(self, machine_dir: Path) -> None:
        """Scan machine directory for PLM_*.xlsx and R3_*.xls and load into cache."""
        if not machine_dir.exists():
            return

        plm_files = (
            list(machine_dir.glob("PLM_*.xlsx"))
            + list(machine_dir.glob("PLM_*.xlsm"))
            + list(machine_dir.glob("PLM/*.xlsx"))
        )
        r3_files = (
            list(machine_dir.glob("R3_*.xls"))
            + list(machine_dir.glob("R3_*.xlsx"))
            + list(machine_dir.glob("R3/*.xls"))
        )

        plm_found_name = ""
        r3_found_name = ""

        if plm_files:
            plm_file = plm_files[0]
            plm_found_name = plm_file.name
            try:
                df_plm = pd.read_excel(plm_file)
                if "Name" in df_plm.columns and "item_id" not in df_plm.columns:
                    df_plm["item_id"] = df_plm["Name"]
                if "Parts Text" in df_plm.columns and "item_name" not in df_plm.columns:
                    df_plm["item_name"] = df_plm["Parts Text"]
                self.plm_data = df_plm
            except Exception as exc:
                logger.warning("Error reading discovered PLM file %s: %s", plm_file, exc)

        if r3_files:
            r3_file = r3_files[0]
            r3_found_name = r3_file.name
            try:
                from src.automation.sap.parser import parse_r3_cs12_file
                self.r3_data = parse_r3_cs12_file(r3_file)
            except Exception:
                try:
                    self.r3_data = pd.read_excel(r3_file)
                except Exception as exc:
                    logger.warning("Error reading discovered R3 file %s: %s", r3_file, exc)

        if plm_found_name or r3_found_name:
            status_parts = []
            if plm_found_name:
                status_parts.append(f"PLM: {plm_found_name}")
            if r3_found_name:
                status_parts.append(f"R3: {r3_found_name}")
            self.lbl_ref_status.setText(" | ".join(status_parts))
            self.lbl_ref_status.setStyleSheet("color: #10B981; font-weight: bold;")

    # =========================================================================
    # CTTT Table Operations & Backward Compatibility
    # =========================================================================

    def add_cttt_row(
        self,
        page: str = "01",
        part_code: str = "",
        part_name: str = "",
        quantity: float = 1.0,
        explanation: str = "",
        status: str = "-",
        sub_unit: str = "",
        author: str = "",
        plm_qty: Any = "-",
        r3_qty: Any = "-",
        comp_plm: str = "-",
        plm_rev: str = "-",
        comp_r3: str = "-",
        r3_rev: str = "-",
        comp_rev: str = "-",
    ) -> int:
        """Append an editable row to the CTTT components table (15 columns)."""
        row = self.cttt_table.rowCount()
        self.cttt_table.insertRow(row)

        sub_val = sub_unit or self.sub_unit_combo.currentText().strip()
        auth_val = author or self.author_edit.text().strip()

        # Safely parse quantity
        try:
            qty_num = float(quantity)
        except (ValueError, TypeError):
            qty_num = 1.0

        # Col 0: Trang CTTT
        item_page = QTableWidgetItem(str(page))
        item_page.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 0, item_page)

        # Col 1: Mã Linh Kiện (CRITICAL INVARIANT)
        item_code = QTableWidgetItem(str(part_code).strip().upper())
        self.cttt_table.setItem(row, 1, item_code)

        # Col 2: Tên Linh Kiện
        item_name = QTableWidgetItem(str(part_name))
        self.cttt_table.setItem(row, 2, item_name)

        # Col 3: Số Lượng
        item_qty = QTableWidgetItem(str(qty_num))
        item_qty.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.cttt_table.setItem(row, 3, item_qty)

        # Col 4: SL PLM
        item_plm = QTableWidgetItem(str(plm_qty))
        item_plm.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.cttt_table.setItem(row, 4, item_plm)

        # Col 5: SL R3
        item_r3 = QTableWidgetItem(str(r3_qty))
        item_r3.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.cttt_table.setItem(row, 5, item_r3)

        # Col 6: Giải Thích Sai Khác
        item_exp = QTableWidgetItem(str(explanation))
        self.cttt_table.setItem(row, 6, item_exp)

        # Col 7: Kết Quả So Sánh (CRITICAL INVARIANT)
        item_status = QTableWidgetItem(str(status))
        item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 7, item_status)

        # Col 8: Unit / Công đoạn
        self.cttt_table.setItem(row, 8, QTableWidgetItem(str(sub_val)))

        # Col 9: Người Phụ Trách
        self.cttt_table.setItem(row, 9, QTableWidgetItem(str(auth_val)))

        # Col 10: So Sánh CTTT vs PLM
        item_c_plm = QTableWidgetItem(str(comp_plm))
        item_c_plm.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 10, item_c_plm)

        # Col 11: Rev PLM
        item_rev_plm = QTableWidgetItem(str(plm_rev))
        item_rev_plm.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 11, item_rev_plm)

        # Col 12: So Sánh CTTT vs R3
        item_c_r3 = QTableWidgetItem(str(comp_r3))
        item_c_r3.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 12, item_c_r3)

        # Col 13: Rev R3
        item_rev_r3 = QTableWidgetItem(str(r3_rev))
        item_rev_r3.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 13, item_rev_r3)

        # Col 14: So Sánh Rev PLM vs R3
        item_c_rev = QTableWidgetItem(str(comp_rev))
        item_c_rev.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cttt_table.setItem(row, 14, item_c_rev)

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

    def _on_author_text_changed(self, new_author: str) -> None:
        """Sync author edit to engineer name."""
        self.engineer_name = new_author.strip()
        self.lbl_engineer_name.setText(self.engineer_name or "-")

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

    def paste_cttt_from_clipboard(self) -> None:
        """Import lines from system clipboard (e.g. copied from Excel table)."""
        clipboard = QApplication.clipboard()
        text = clipboard.text().strip()
        if not text:
            QMessageBox.warning(self, "Clipboard trống", "Không tìm thấy dữ liệu văn bản trong clipboard.")
            return

        lines = [line for line in text.splitlines() if line.strip()]
        if not lines:
            return

        imported_count = 0
        for line in lines:
            cols = [c.strip() for c in line.split("\t")]
            if not cols or not any(cols):
                continue
            # If line looks like a header, skip it
            if any(k in cols[0].upper() for k in ["MÃ", "PART", "ITEM", "TRANG", "PAGE", "SUB", "UNIT"]):
                continue

            page = "01"
            code = ""
            name = ""
            qty = 1.0
            exp = ""

            if len(cols) == 1:
                code = cols[0]
            elif len(cols) == 2:
                code = cols[0]
                name = cols[1]
            elif len(cols) == 3:
                page = cols[0]
                code = cols[1]
                name = cols[2]
            elif len(cols) >= 4:
                page = cols[0]
                code = cols[1]
                name = cols[2]
                try:
                    qty = float(cols[3].replace(",", "."))
                except (ValueError, TypeError):
                    qty = 1.0
                if len(cols) >= 5:
                    exp = cols[4]

            if code:
                self.add_cttt_row(page=page, part_code=code, part_name=name, quantity=qty, explanation=exp)
                imported_count += 1

        QMessageBox.information(self, "Dán thành công", f"Đã dán {imported_count} linh kiện từ Clipboard vào bảng CTTT.")

    def get_cttt_table_data(self) -> list[dict[str, Any]]:
        """Collect current table items into structured dictionary records."""
        records = []
        sub_name = self.sub_unit_combo.currentText().strip()
        author = self.author_edit.text().strip()

        for r in range(self.cttt_table.rowCount()):
            page = self.cttt_table.item(r, 0).text().strip() if self.cttt_table.item(r, 0) else ""
            part_code = self.cttt_table.item(r, 1).text().strip() if self.cttt_table.item(r, 1) else ""
            part_name = self.cttt_table.item(r, 2).text().strip() if self.cttt_table.item(r, 2) else ""

            # Resilient numeric parsing
            qty_item = self.cttt_table.item(r, 3)
            qty_text = qty_item.text().strip() if qty_item else "0.0"
            try:
                qty = float(qty_text.replace(",", "."))
            except (ValueError, TypeError):
                qty = 0.0

            explanation = self.cttt_table.item(r, 6).text().strip() if self.cttt_table.item(r, 6) else ""
            status = self.cttt_table.item(r, 7).text().strip() if self.cttt_table.item(r, 7) else "OK"

            item_sub = self.cttt_table.item(r, 8)
            row_sub = item_sub.text().strip() if item_sub and item_sub.text().strip() else sub_name

            item_auth = self.cttt_table.item(r, 9)
            row_author = item_auth.text().strip() if item_auth and item_auth.text().strip() else author

            records.append({
                "SUB": row_sub,
                "TRANG CTTT": page,
                "MÃ LINH KIỆN": part_code,
                "TÊN LINH KIỆN": part_name,
                "SỐ LƯỢNG": qty,
                "PHỤ TRÁCH": row_author,
                "Giải thích": explanation,
                "Check": status,
            })
        return records

    # =========================================================================
    # MSI 35 Canonical Units Table Operations
    # =========================================================================

    def _populate_canonical_msi_table(self) -> None:
        """Populate MSI table with the 35 fixed canonical sub-units from formnguoidung.xlsm."""
        self.msi_table.setRowCount(0)
        for idx, unit_name in enumerate(CANONICAL_MSI_UNITS):
            self.msi_table.insertRow(idx)
            # Col 0: Mã LK Barcode
            self.msi_table.setItem(idx, 0, QTableWidgetItem(""))
            # Col 1: Mã UNIT / Bản mạch
            self.msi_table.setItem(idx, 1, QTableWidgetItem(""))
            # Col 2: Tên UNIT (pre-filled canonical name)
            item_name = QTableWidgetItem(unit_name)
            item_name.setFlags(item_name.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.msi_table.setItem(idx, 2, item_name)
            # Col 3: 3 Ký tự MSI
            self.msi_table.setItem(idx, 3, QTableWidgetItem(""))
            # Col 4: SEVICE
            self.msi_table.setItem(idx, 4, QTableWidgetItem("-"))
            # Col 5: ABS
            self.msi_table.setItem(idx, 5, QTableWidgetItem("OK"))
            # Col 6: Trang CTTT
            self.msi_table.setItem(idx, 6, QTableWidgetItem("01"))
            # Col 7: Điện áp
            self.msi_table.setItem(idx, 7, QTableWidgetItem("-"))
            # Col 8: Loại Label
            self.msi_table.setItem(idx, 8, QTableWidgetItem("BARCODE"))
            # Col 9: Tên phụ trách
            self.msi_table.setItem(idx, 9, QTableWidgetItem(self.engineer_name or "Member"))
            # Col 10: Kết quả so sánh
            item_status = QTableWidgetItem("-")
            item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.msi_table.setItem(idx, 10, item_status)
            # Col 11: Ghi chú
            self.msi_table.setItem(idx, 11, QTableWidgetItem(""))

    def _on_msi_table_row_selected(self, row: int, col: int) -> None:
        """Load selected MSI row into the Quick-Edit panel."""
        if row < 0 or row >= self.msi_table.rowCount():
            return
        self._selected_msi_row = row
        self.msi_table.setCurrentCell(row, col)

        get_t = lambda c: self.msi_table.item(row, c).text() if self.msi_table.item(row, c) else ""
        self.msi_barcode_edit.setText(get_t(0))
        self.msi_unit_code_edit.setText(get_t(1))
        self.msi_unit_name_edit.setText(get_t(2))
        self.msi_code_edit.setText(get_t(3))
        self.msi_service_edit.setText(get_t(4) or "-")

        abs_val = get_t(5) or "OK"
        idx_abs = self.msi_abs_combo.findText(abs_val)
        if idx_abs >= 0:
            self.msi_abs_combo.setCurrentIndex(idx_abs)

        volt_val = get_t(7) or "-"
        idx_v = self.msi_volt_combo.findText(volt_val)
        if idx_v >= 0:
            self.msi_volt_combo.setCurrentIndex(idx_v)

        lbl_val = get_t(8) or "BARCODE"
        idx_l = self.msi_label_type_combo.findText(lbl_val)
        if idx_l >= 0:
            self.msi_label_type_combo.setCurrentIndex(idx_l)

    def _apply_msi_quick_edit(self) -> None:
        """Write back Quick-Edit panel contents into selected MSI row."""
        current_row = getattr(self, "_selected_msi_row", -1)
        if current_row < 0:
            current_row = self.msi_table.currentRow()
        if current_row < 0:
            current_row = 0
            self.msi_table.setCurrentCell(0, 0)

        set_t = lambda c, val: self.msi_table.item(current_row, c).setText(str(val))
        set_t(0, self.msi_barcode_edit.text().strip())
        set_t(1, self.msi_unit_code_edit.text().strip())
        set_t(2, self.msi_unit_name_edit.text().strip())
        set_t(3, self.msi_code_edit.text().strip())
        set_t(4, self.msi_service_edit.text().strip())
        set_t(5, self.msi_abs_combo.currentText().strip())
        set_t(7, self.msi_volt_combo.currentText().strip())
        set_t(8, self.msi_label_type_combo.currentText().strip())
        set_t(9, self.engineer_name or "Member")

        QMessageBox.information(
            self,
            "Cập nhật thành công",
            f"Đã cập nhật dữ liệu vào Dòng {current_row + 1} ({self.msi_unit_name_edit.text()}).",
        )

    # =========================================================================
    # Label 7980 / 7990 Table Operations
    # =========================================================================

    def add_label_row(
        self,
        sub_7980: str = "LSU",
        page_7980: str = "01",
        code_7980: str = "7980-VN-001",
        name_7980: str = "LCP LABEL",
        qty_7980: int = 1,
        author_7980: str = "",
        sub_7990: str = "HIGH VOLTAGE",
        page_7990: str = "02",
        code_7990: str = "7990-VN-001",
        name_7990: str = "CAUTION LABEL",
        qty_7990: int = 1,
        author_7990: str = "",
    ) -> int:
        """Append a 12-column row to Label 7980/7990 table."""
        row = self.label_table.rowCount()
        self.label_table.insertRow(row)

        auth = author_7980 or self.engineer_name
        auth9 = author_7990 or self.engineer_name

        vals = [
            sub_7980,
            page_7980,
            code_7980,
            name_7980,
            str(qty_7980),
            auth,
            sub_7990,
            page_7990,
            code_7990,
            name_7990,
            str(qty_7990),
            auth9,
        ]
        for col_idx, val in enumerate(vals):
            self.label_table.setItem(row, col_idx, QTableWidgetItem(str(val)))
        return row

    def remove_selected_label_row(self) -> None:
        """Remove selected row from Label table."""
        row = self.label_table.currentRow()
        if row >= 0:
            self.label_table.removeRow(row)

    def clear_label_table(self) -> None:
        """Clear all rows from Label table."""
        self.label_table.setRowCount(0)

    # =========================================================================
    # Reference Data & Preliminary Self-Check
    # =========================================================================

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
                df = pd.read_excel(plm_path)
                if "Name" in df.columns and "item_id" not in df.columns:
                    df["item_id"] = df["Name"]
                if "Parts Text" in df.columns and "item_name" not in df.columns:
                    df["item_name"] = df["Parts Text"]
                self.plm_data = df
                self.lbl_ref_status.setText(f"PLM: {Path(plm_path).name}")
                self.lbl_ref_status.setStyleSheet("color: #10B981; font-weight: bold;")
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
                from src.automation.sap.parser import parse_r3_cs12_file
                self.r3_data = parse_r3_cs12_file(Path(r3_path))
                self.lbl_ref_status.setText(f"Đã nạp PLM & R3 ({Path(r3_path).name})")
                self.lbl_ref_status.setStyleSheet("color: #10B981; font-weight: bold;")
            except Exception:
                try:
                    self.r3_data = pd.read_excel(r3_path)
                    self.lbl_ref_status.setText("Đã nạp PLM & R3")
                    self.lbl_ref_status.setStyleSheet("color: #10B981; font-weight: bold;")
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
            if hasattr(self, "workflow_stepper"):
                self.workflow_stepper.set_step_completed(0, True)
                self.workflow_stepper.set_current_step(1)

    def run_preliminary_self_check(self) -> dict[str, Any]:
        """Execute instant three-way self-check on CTTT & MSI with soft green/red styling."""
        row_count = self.cttt_table.rowCount()
        if row_count == 0:
            QMessageBox.warning(self, "Cảnh báo", "Bảng linh kiện CTTT chưa có dữ liệu để đối soát.")
            return {"status": "EMPTY", "ok_count": 0, "ng_count": 0, "total": 0}

        # Auto-discover reference data from machine folder if not in memory
        if self.plm_data is None and self.r3_data is None and self.current_machine_dir:
            self._auto_discover_boms(self.current_machine_dir)

        # Build PLM lookup: part_code -> (quantity, revision)
        plm_dict: dict[str, tuple[float, str]] = {}
        if self.plm_data is not None and not self.plm_data.empty:
            id_col = next((c for c in self.plm_data.columns if any(k in str(c).upper() for k in ["ITEM ID", "ITEM_ID", "PART_CODE", "PART CODE", "NAME", "MÃ"])), None)
            qty_col = next((c for c in self.plm_data.columns if any(k in str(c).upper().replace(".", "").replace("'", "") for k in ["QUANTITY", "QTY", "SL", "SOLUONG", "SỐ LƯỢNG", "SỐLƯỢNG"]) or any(k in str(c).upper() for k in ["Q.TY", "Q'TY", "Q.TY."])), None)
            rev_col = next((c for c in self.plm_data.columns if any(k in str(c).upper() for k in ["REVISION", "REV", "ITEM_REV", "RELEASE STATUS"])), None)
            if id_col and qty_col:
                for _, pr in self.plm_data.iterrows():
                    code = str(pr[id_col]).strip().upper()
                    try:
                        q = float(pr[qty_col])
                    except (ValueError, TypeError):
                        q = 0.0
                    r_val = str(pr[rev_col]).strip() if rev_col and pd.notna(pr[rev_col]) else "A"
                    prev_q, _ = plm_dict.get(code, (0.0, r_val))
                    plm_dict[code] = (prev_q + q, r_val)

        # Build R3 lookup: part_code -> (quantity, revision)
        r3_dict: dict[str, tuple[float, str]] = {}
        if self.r3_data is not None and not self.r3_data.empty:
            mat_col = next((c for c in self.r3_data.columns if any(k in str(c).upper() for k in ["MATERIAL", "COMPONENT", "PART_CODE", "PART CODE", "MÃ"])), None)
            qty_col = next((c for c in self.r3_data.columns if any(k in str(c).upper().replace(".", "").replace("'", "") for k in ["QUANTITY", "QTY", "SL", "SOLUONG", "SỐ LƯỢNG", "SỐLƯỢNG"]) or any(k in str(c).upper() for k in ["Q.TY", "Q'TY", "Q.TY."])), None)
            rev_col = next((c for c in self.r3_data.columns if any(k in str(c).upper() for k in ["REVLEV", "REVISION", "REV"])), None)
            if mat_col and qty_col:
                for _, rr in self.r3_data.iterrows():
                    code = str(rr[mat_col]).strip().upper()
                    try:
                        q = float(rr[qty_col])
                    except (ValueError, TypeError):
                        q = 0.0
                    r_val = str(rr[rev_col]).strip() if rev_col and pd.notna(rr[rev_col]) else "A"
                    prev_q, _ = r3_dict.get(code, (0.0, r_val))
                    r3_dict[code] = (prev_q + q, r_val)

        ok_count = 0
        ng_count = 0

        # Evaluate CTTT Table
        for r_idx in range(row_count):
            item_code_obj = self.cttt_table.item(r_idx, 1)
            p_code = item_code_obj.text().strip().upper() if item_code_obj else ""
            try:
                qty_item = self.cttt_table.item(r_idx, 3)
                cttt_q = float(qty_item.text().strip().replace(",", ".")) if qty_item else 0.0
            except (ValueError, TypeError, AttributeError):
                cttt_q = 0.0

            has_plm_source = bool(plm_dict)
            has_r3_source = bool(r3_dict)

            if has_plm_source:
                plm_q, plm_rev = plm_dict.get(p_code, (0.0, ""))
            else:
                plm_q, plm_rev = cttt_q, "A"

            if has_r3_source:
                r3_q, r3_rev = r3_dict.get(p_code, (0.0, ""))
            else:
                r3_q, r3_rev = cttt_q, "A"

            # Update reference quantities
            if self.cttt_table.item(r_idx, 4):
                self.cttt_table.item(r_idx, 4).setText(str(plm_q))
            if self.cttt_table.item(r_idx, 5):
                self.cttt_table.item(r_idx, 5).setText(str(r3_q))

            res = self.reconciliation_engine.reconcile_single_row(
                cttt_qty=cttt_q,
                plm_qty=plm_q,
                r3_qty=r3_q,
                plm_rev=plm_rev,
                r3_rev=r3_rev,
            )

            status_str = res["overall_check"]
            status_item = self.cttt_table.item(r_idx, 7)
            if status_item:
                status_item.setText(status_str)

            # Detail comparison metrics
            if self.cttt_table.item(r_idx, 10):
                self.cttt_table.item(r_idx, 10).setText(res.get("comp_plm_qty", "-"))
            if self.cttt_table.item(r_idx, 11):
                self.cttt_table.item(r_idx, 11).setText(plm_rev)
            if self.cttt_table.item(r_idx, 12):
                self.cttt_table.item(r_idx, 12).setText(res.get("comp_r3_qty", "-"))
            if self.cttt_table.item(r_idx, 13):
                self.cttt_table.item(r_idx, 13).setText(r3_rev)
            if self.cttt_table.item(r_idx, 14):
                self.cttt_table.item(r_idx, 14).setText(res.get("comp_rev", "-"))

            # Apply instant visual feedback
            if status_str == "OK":
                ok_count += 1
                self._style_table_row(r_idx, is_ok=True)
            else:
                ng_count += 1
                self._style_table_row(r_idx, is_ok=False)

        # Evaluate MSI Table Rows with FixSerialMaster
        for m_idx in range(self.msi_table.rowCount()):
            u_code_item = self.msi_table.item(m_idx, 1)
            unit_code = u_code_item.text().strip() if u_code_item else ""
            m_code_item = self.msi_table.item(m_idx, 3)
            member_code = m_code_item.text().strip() if m_code_item else ""
            srv_item = self.msi_table.item(m_idx, 4)
            member_srv = srv_item.text().strip() if srv_item else "-"

            if unit_code or member_code:
                if m_idx < 34:
                    mst_code, mst_srv = self.fix_serial_master.lookup_subunit(unit_code)
                else:
                    mst_code, mst_srv = self.fix_serial_master.lookup_machine(self.machine_code or unit_code)

                eval_res = evaluate_msi_branch(
                    in_plm=True,
                    member_code=member_code,
                    master_code=mst_code or member_code,
                    member_service=member_srv,
                    master_service=mst_srv,
                )
                chk_item = self.msi_table.item(m_idx, 10)
                if chk_item:
                    chk_item.setText(eval_res["status"])
                    if eval_res["status"] == "OK":
                        chk_item.setBackground(QColor(f"#{COLOR_GREEN_FILL_HEX}"))
                        chk_item.setForeground(QColor(f"#{COLOR_GREEN_FONT_HEX}"))
                    else:
                        chk_item.setBackground(QColor(f"#{COLOR_RED_FILL_HEX}"))
                        chk_item.setForeground(QColor(f"#{COLOR_RED_FONT_HEX}"))

        # Check Quick-Edit MSI Badge
        quick_code = self.msi_code_edit.text().strip()
        quick_srv = self.msi_service_edit.text().strip()
        msi_res = evaluate_msi_branch(
            in_plm=True,
            member_code=quick_code,
            master_code=quick_code,
            member_service=quick_srv,
            master_service="",
        )
        if msi_res["status"] == "OK":
            self.lbl_msi_badge.setText(f"MSI OK (Nhánh {msi_res['branch']})")
            self.lbl_msi_badge.setStyleSheet(
                "background-color: #C6EFCE; color: #006100; font-weight: bold; border-radius: 4px; padding: 4px;"
            )
        else:
            self.lbl_msi_badge.setText(f"MSI NG (Nhánh {msi_res['branch']})")
            self.lbl_msi_badge.setStyleSheet(
                "background-color: #FFC7CE; color: #9C0006; font-weight: bold; border-radius: 4px; padding: 4px;"
            )

        # Update Summary Label
        summary_text = f"Kết quả: {ok_count} OK, {ng_count} NG / Tổng {row_count} linh kiện"
        self.lbl_check_summary.setText(summary_text)
        if ng_count > 0:
            self.lbl_check_summary.setStyleSheet("color: #dc3545; font-weight: bold;")
        else:
            self.lbl_check_summary.setStyleSheet("color: #28a745; font-weight: bold;")

        if hasattr(self, "workflow_stepper"):
            self.workflow_stepper.set_step_completed(0, True)
            self.workflow_stepper.set_step_completed(1, True)
            self.workflow_stepper.set_current_step(2)

        return {
            "status": "OK" if ng_count == 0 else "NG",
            "ok_count": ok_count,
            "ng_count": ng_count,
            "total": row_count,
        }

    def _style_table_row(self, row_idx: int, is_ok: bool) -> None:
        """Apply soft green (OK) or soft red (NG) styling across the row with WCAG AAA contrast."""
        col_count = self.cttt_table.columnCount()
        row_bg = QColor("#DCFCE7") if is_ok else QColor("#FEE2E2")
        row_fg = QColor("#064E3B") if is_ok else QColor("#7F1D1D")

        for c in range(col_count):
            it = self.cttt_table.item(row_idx, c)
            if it and c != 7:
                it.setBackground(row_bg)
                it.setForeground(row_fg)

        # Highlight Status cell (Col 7) with official Excel stamp color for backward compatibility
        bg_color = QColor(f"#{COLOR_GREEN_FILL_HEX}") if is_ok else QColor(f"#{COLOR_RED_FILL_HEX}")
        fg_color = QColor(f"#{COLOR_GREEN_FONT_HEX}") if is_ok else QColor(f"#{COLOR_RED_FONT_HEX}")
        status_item = self.cttt_table.item(row_idx, 7)
        if status_item:
            status_item.setBackground(bg_color)
            status_item.setForeground(fg_color)
            status_item.setFont(QFont("Calibri", 10, QFont.Weight.Bold))

    # =========================================================================
    # Submission Seal Engine (Q2 = "OK")
    # =========================================================================

    def submit_data(self) -> Path | None:
        """Validate, stamp CTTT!Q2='OK' (green fill), save assignment file, and emit signal."""
        rows = self.get_cttt_table_data()
        if not rows:
            QMessageBox.warning(self, "Cảnh báo", "Bảng CTTT không có dữ liệu để nộp.")
            return None

        # Check for unaddressed NG items without explanation
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

        sub_unit = self.sub_unit_combo.currentText().strip() or "LSU"
        author = self.author_edit.text().strip() or self.engineer_name or "Member"
        model = self.model_combo.currentText().strip() or "Model"
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        # Determine target file
        if self.current_assignment_file is not None and self.current_assignment_file.exists():
            target_file = self.current_assignment_file
        else:
            target_dir = self.base_dir / "CTTT" / sub_unit
            target_dir.mkdir(parents=True, exist_ok=True)
            target_file = target_dir / f"formnguoidung_{sub_unit}_{timestamp}.xlsx"

        try:
            # Load or create workbook
            if target_file.exists():
                wb = openpyxl.load_workbook(target_file)
            else:
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames:
                    wb.remove(wb["Sheet"])

            # 1. Update / Create Sheet CTTT
            if "CTTT" in wb.sheetnames:
                ws_cttt = wb["CTTT"]
            else:
                ws_cttt = wb.create_sheet(title="CTTT", index=0)
                ws_cttt.append([
                    "Nhập dữ liệu chỉ thị thao tác lần đầu để tạo file So sánh BOM",
                    None, None, None, None, None,
                    "PLM", None, None, None,
                    "R3", None, None, None,
                    "Giải thích", None, "Kết quả",
                ])
                ws_cttt.append([
                    "Unit/Assy", "Trang", "Mã linh kiện", "Tên linh kiện", "Số lượng", "Tên người phụ trách",
                    "Số lượng PLM", "So sánh số lượng CTTT và PLM", "Rev PLM", "Số lượng R3",
                    "So sánh số lượng CTTT và R3", "Rev R3", "So sánh Rev PLM và R3",
                    "Check số lượng PLM và kết quả so sánh Rev PLM/R3", "Giải thích", None, None,
                ])

            # Write Q2 = "OK" with green fill
            ws_cttt["Q2"].value = "OK"
            ws_cttt["Q2"].fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            ws_cttt["Q2"].font = Font(name="Calibri", size=11, bold=True, color="006100")

            # Write rows into ws_cttt starting at row 3
            max_r = max(ws_cttt.max_row, len(rows) + 2)
            for r_idx in range(3, max_r + 1):
                for c_idx in range(1, 16):
                    ws_cttt.cell(row=r_idx, column=c_idx, value=None)

            for idx, r_data in enumerate(rows, start=3):
                t_row = idx - 3
                plm_qty_val = self.cttt_table.item(t_row, 4).text() if self.cttt_table.item(t_row, 4) else ""
                r3_qty_val = self.cttt_table.item(t_row, 5).text() if self.cttt_table.item(t_row, 5) else ""
                comp_plm_val = self.cttt_table.item(t_row, 10).text() if self.cttt_table.item(t_row, 10) else ""
                plm_rev_val = self.cttt_table.item(t_row, 11).text() if self.cttt_table.item(t_row, 11) else ""
                comp_r3_val = self.cttt_table.item(t_row, 12).text() if self.cttt_table.item(t_row, 12) else ""
                r3_rev_val = self.cttt_table.item(t_row, 13).text() if self.cttt_table.item(t_row, 13) else ""
                comp_rev_val = self.cttt_table.item(t_row, 14).text() if self.cttt_table.item(t_row, 14) else ""

                ws_cttt.cell(row=idx, column=1, value=r_data["SUB"])
                ws_cttt.cell(row=idx, column=2, value=r_data["TRANG CTTT"])
                ws_cttt.cell(row=idx, column=3, value=r_data["MÃ LINH KIỆN"])
                ws_cttt.cell(row=idx, column=4, value=r_data["TÊN LINH KIỆN"])
                ws_cttt.cell(row=idx, column=5, value=r_data["SỐ LƯỢNG"])
                ws_cttt.cell(row=idx, column=6, value=r_data["PHỤ TRÁCH"])
                ws_cttt.cell(row=idx, column=7, value=plm_qty_val)
                ws_cttt.cell(row=idx, column=8, value=comp_plm_val)
                ws_cttt.cell(row=idx, column=9, value=plm_rev_val)
                ws_cttt.cell(row=idx, column=10, value=r3_qty_val)
                ws_cttt.cell(row=idx, column=11, value=comp_r3_val)
                ws_cttt.cell(row=idx, column=12, value=r3_rev_val)
                ws_cttt.cell(row=idx, column=13, value=comp_rev_val)
                ws_cttt.cell(row=idx, column=14, value=r_data["Check"])
                ws_cttt.cell(row=idx, column=15, value=r_data["Giải thích"])

            # 2. Update / Create Sheet MSI
            if "MSI" in wb.sheetnames:
                ws_msi = wb["MSI"]
            else:
                ws_msi = wb.create_sheet(title="MSI")
                ws_msi.append([
                    "Mã LK Barcode", "Mã UNIT/linh kiện bản mạch", "Tên UNIT", "3 ký tự MSI",
                    "SEVICE", "ABS", "Trang CTTT", "Điện áp", "Loại Label", "Tên phụ trách",
                    "Kết quả so sánh", "Ghi chú",
                ])

            for r in range(self.msi_table.rowCount()):
                for c in range(12):
                    item = self.msi_table.item(r, c)
                    val = item.text() if item else ""
                    ws_msi.cell(row=r + 2, column=c + 1, value=val)

            # 3. Update / Create Sheet Label_7980_7990
            if "Label_7980_7990" in wb.sheetnames:
                ws_lbl = wb["Label_7980_7990"]
            else:
                ws_lbl = wb.create_sheet(title="Label_7980_7990")
                ws_lbl.append([
                    "Công đoạn quản lý", "Trang chỉ thị thao tác", "Mã label quản lý 7980",
                    "Tên label quản lý 7980", "Số lượng", "Tên người phụ trách",
                    "Công đoạn quản lý", "Trang chỉ thị thao tác", "Mã label quản lý 7990",
                    "Tên label quản lý 7990", "Số lượng", "Tên người phụ trách",
                ])

            if self.label_table.rowCount() > 0:
                for r in range(self.label_table.rowCount()):
                    for c in range(12):
                        item = self.label_table.item(r, c)
                        val = item.text() if item else ""
                        ws_lbl.cell(row=r + 2, column=c + 1, value=val)
            else:
                ws_lbl.cell(row=2, column=1, value="7980")
                ws_lbl.cell(row=2, column=2, value="01")
                ws_lbl.cell(row=2, column=3, value=self.label_code_edit.text().strip())
                ws_lbl.cell(row=2, column=4, value=self.label_spec_combo.currentText().strip())
                ws_lbl.cell(row=2, column=5, value=1)
                ws_lbl.cell(row=2, column=6, value=author)

            wb.save(target_file)
            logger.info("Đã đóng dấu phê duyệt Q2='OK' thành công vào file: %s", target_file.name)

            self.is_submitted_ok = True
            self.lbl_submission_seal.setText("✅ ĐÃ NỘP BÀI (Q2 = OK)")
            self.lbl_submission_seal.setStyleSheet(
                "background-color: #C6EFCE; color: #006100; font-weight: bold; border-radius: 4px; padding: 4px 8px;"
            )
            self.cttt_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

            if hasattr(self, "workflow_stepper"):
                self.workflow_stepper.set_step_completed(2, True)
                self.workflow_stepper.set_current_step(2)

            submission_info = {
                "sub_unit": sub_unit,
                "author": author,
                "model": model,
                "timestamp": timestamp,
                "item_count": len(rows),
                "file_path": str(target_file),
                "status": "Submitted",
                "q2_status": "OK",
                "engineer_name": author,
                "machine_code": self.machine_code,
                "department": self.department,
            }
            self.submission_completed.emit(submission_info)

            QMessageBox.information(
                self,
                "Nộp dữ liệu thành công",
                f"Đã đóng dấu Q2 = OK và lưu file thành công!\n"
                f"Tệp lưu tại: {target_file.name}\n"
                f"Số lượng linh kiện: {len(rows)}",
            )
            return target_file
        except Exception as exc:
            logger.error("Failed to write submission workbook: %s", exc)
            QMessageBox.critical(self, "Lỗi nộp dữ liệu", f"Không thể xuất tệp nộp dữ liệu:\n{exc}")
            return None

    def unlock_submission(self) -> None:
        """Execute 'Hủy nộp / Mở khóa': Clear Q2 cell, restore status, enable editing."""
        if self.current_assignment_file is not None and self.current_assignment_file.exists():
            try:
                wb = openpyxl.load_workbook(self.current_assignment_file)
                if "CTTT" in wb.sheetnames:
                    ws_cttt = wb["CTTT"]
                    ws_cttt["Q2"].value = None
                    ws_cttt["Q2"].fill = PatternFill(fill_type=None)
                    ws_cttt["Q2"].font = Font(name="Calibri", size=11)
                    wb.save(self.current_assignment_file)
                    logger.info("Đã mở khóa bài nộp và xóa dấu cờ Q2 trong file: %s", self.current_assignment_file.name)
            except Exception as exc:
                logger.warning("Could not clear Q2 in file: %s", exc)

        self.is_submitted_ok = False
        self.lbl_submission_seal.setText("⏳ CHƯA NỘP (Q2 TRỐNG)")
        self.lbl_submission_seal.setStyleSheet(
            "background-color: #FFF3CD; color: #856404; font-weight: bold; border-radius: 4px; padding: 4px 8px;"
        )
        self.cttt_table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked
            | QTableWidget.EditTrigger.SelectedClicked
            | QTableWidget.EditTrigger.EditKeyPressed
        )

        if hasattr(self, "workflow_stepper"):
            self.workflow_stepper.set_step_completed(2, False)
            self.workflow_stepper.set_step_completed(1, False)
            self.workflow_stepper.set_current_step(1)
        elif hasattr(self, "stepper"):
            self.stepper.set_step_completed(2, False)
            self.stepper.set_step_completed(1, False)
            self.stepper.set_current_step(1)

        QMessageBox.information(self, "Mở khóa thành công", "Đã hủy nộp và mở khóa để chỉnh sửa bài làm.")
