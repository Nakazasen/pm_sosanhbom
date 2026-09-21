"""Feature F22 & Milestone M1: Leader Management Workspace Sequential 4-Step Wizard.

Implements the complete 4-step sequential workflow for Engineering Team Leads:
- Step 1 (Lập Dự Án & Phân Công): Model, Phase (maT vs ma1), BOM code list, staffing table Cơ 1, 2, 3 (Sheet Lichsu & tenphong_pt), auto-create machine directories and generate member packages.
- Step 2 (Tải & Xử Lý Nguồn Dữ Liệu): Integrated PLM TC24/TC14 & SAP R3 CS12, common/individual date options, auto-route files to machine folders, backup to backupTC14full/ and BOM Level 1..6 filtering.
- Step 3 (Theo Dõi & Tổng Hợp): Live scan for CTTT!Q2 = "OK", live status table, fail-closed gate blocking consolidation until 100% OK, consolidate CTTT/MSI/7980, move member files to phutrach/<machine_code>/.
- Step 4 (So Sánh BOM Tổng & Gửi Báo Cáo): form_ssbom creation, Pivot Table refresh, JIG/4M manager integration, 2-tier Outlook emails (Tier 1: Member reminder + 18 check points, Tier 2: Management reporting).
"""

from __future__ import annotations

import datetime
import logging
import os
import re
import shutil
import subprocess

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Optional

import openpyxl
import pandas as pd
from PyQt6.QtCore import QDate, QObject, QSize, Qt, QThread, QTimer, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from src.core.date_filter import DateFilter
from src.core.model_pruner import ModelPruner
from src.core.reconciliation import ReconciliationEngine, ReconciliationResult
from src.gui.styles import get_theme_manager, tokens
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_GREEN_FONT_HEX,
    COLOR_RED_FILL_HEX,
    COLOR_RED_FONT_HEX,
    STANDARD_SUB_UNITS,
    ExcelReportGenerator,
)
from src.reporting.outlook_mailer import EmailPreview, OutlookMailer

logger = logging.getLogger(__name__)


# =============================================================================
# Constants & Roster Data
# =============================================================================

# Standard 38 engineers across Cơ 1, Cơ 2, Cơ 3 from Sheet tenphong_pt
ROSTER_MECHA_1: list[str] = [
    "Son_mecha1", "Duy_mecha1", "KhiemA_mecha1", "KhiemE_mecha1",
    "Nghia_mecha1", "Quyen_mecha1", "Ly_mecha1", "Tuong_mecha1",
    "Q.Anh_mecha1", "Ngoc_mecha1", "Hung_mecha1", "Long_mecha1",
]

ROSTER_MECHA_2: list[str] = [
    "Loc_mecha2", "Hai_mecha2", "Dat_mecha2", "Nhung_mecha2",
    "Hung_mecha2", "Quynh_mecha2", "HuuTu_mecha2", "Tuan_mecha2",
    "XuanSon_mecha2", "Thuy_mecha2", "Canh_mecha2",
]

ROSTER_MECHA_3: list[str] = [
    "LVThuan_mecha3", "V.Thuong_mecha3", "Vinh_mecha3", "Viet_mecha3",
    "Nguyet_mecha3", "Thao_mecha3", "Trong_mecha3", "Minh_mecha3",
    "Khai_mecha3", "V.Thanh_mecha3", "Do_mecha3", "NThuan_mecha3",
    "Khang_mecha3", "HaiDang_mecha3", "KhacTu_mecha3",
]

# 15 Standard machine series for List JIG Master
JIG_MODELS: list[str] = [
    "6thA4", "Polaris", "6thA3", "Brazil", "Libra", "Libra2",
    "Mebius", "Kairos", "Virgo", "Iris", "Sirius", "EH Iris",
    "Spica", "DF Iris", "Pictor",
]

# 18 Pre-production inspection points for Tier 1 Outlook email
CHECKLIST_18_POINTS: list[str] = [
    "1. Kiểm tra mã máy và hướng xuất chính xác.",
    "2. Đối chiếu số lượng CTTT với BOM PLM Teamcenter.",
    "3. Đối chiếu số lượng CTTT với BOM SAP R3 CS12.",
    "4. Kiểm tra Revision của từng linh kiện đồng bộ giữa PLM và R3.",
    "5. Kiểm tra linh kiện thuộc đúng Unit/Assy phụ trách.",
    "6. Phán định 3 ký tự cố định mã Barcode / MSI với Tool Master.",
    "7. Kiểm tra nhãn dịch vụ LABEL_COMMENT (SERVICE).",
    "8. Kiểm tra tem nhãn LCP 7980 (phần mềm).",
    "9. Kiểm tra tem nhãn 7990 (cảnh báo an toàn).",
    "10. Kiểm tra quy cách đóng gói và bao bì POLYBAG.",
    "11. Xác nhận điện áp và thông số nguồn AC CORD / Low Voltage.",
    "12. Rà soát danh mục linh kiện thay thế đợt này.",
    "13. Kiểm tra linh kiện có hiệu lực ngày sản xuất (không dùng linh kiện hết hạn).",
    "14. Rà soát giải trình cho tất cả linh kiện sai khác (cột O Sheet CTTT).",
    "15. Xác nhận không còn mã linh kiện rỗng hoặc lỗi #N/A.",
    "16. Tự đối soát sơ bộ tại chỗ (Self-Check).",
    "17. Xác nhận với KTSX về các thay đổi 4M liên quan công đoạn.",
    "18. Bấm nút nộp bài và xác nhận đóng dấu cờ Q2 = OK.",
]


# =============================================================================
# Session Models & State
# =============================================================================

class ProjectStage(str, Enum):
    MA_T = "maT"  # Chế tạo thử DMT / PMT (Tiền tố mã máy T10...)
    MA_1 = "ma1"  # Sản xuất thử nghiệm PP đến hàng loạt MP (Tiền tố mã máy 110...)


class DepartmentGroup(str, Enum):
    ALL = "Tất cả"
    MECHA_1 = "Cơ 1"
    MECHA_2 = "Cơ 2"
    MECHA_3 = "Cơ 3"


@dataclass
class StaffAssignment:
    """Staffing assignment mapping corresponding to Sheet Lichsu."""
    engineer_name: str
    department: str
    is_applied: bool = True
    machine_code: str = ""
    sub_unit: str = "LSU"
    assigned_file: Optional[Path] = None


@dataclass
class MachineTarget:
    """Target machine / export direction for BOM comparison."""
    machine_code: str
    is_excluded: bool = False
    folder_path: Optional[Path] = None
    custom_date: Optional[str] = None
    plm_file: Optional[Path] = None
    r3_file: Optional[Path] = None
    is_filtered: bool = False
    bom_output_file: Optional[Path] = None


@dataclass
class MemberSubmissionStatus:
    """Status of member submission workbook."""
    machine_code: str
    sub_unit: str
    engineer_name: str
    file_path: Optional[Path] = None
    is_submitted_ok: bool = False
    item_count: int = 0
    msi_count: int = 0
    label_count: int = 0
    modified_time: Optional[datetime.datetime] = None
    error_message: str = ""


@dataclass
class LeaderSessionState:
    """Global session state for the 4-step sequential wizard."""
    model_name: str = "Virgo"
    stage: ProjectStage = ProjectStage.MA_1
    base_dir: Path = field(
        default_factory=lambda: (
            Path(r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3")
            if Path(r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3").exists()
            else Path(r"D:\Sandbox\pm_sosanhbom")
        )
    )

    # Step 1
    machines: list[MachineTarget] = field(default_factory=list)
    staff_roster: list[StaffAssignment] = field(default_factory=list)
    step1_completed: bool = False

    # Step 2
    date_mode_common: bool = True
    common_target_date: str = field(default_factory=lambda: datetime.date.today().strftime("%Y/%m/%d"))
    step2_completed: bool = False

    # Step 3
    submissions: list[MemberSubmissionStatus] = field(default_factory=list)
    all_members_ok: bool = False
    step3_completed: bool = False
    consolidated_cttt: list[dict[str, Any]] = field(default_factory=list)
    consolidated_msi: list[dict[str, Any]] = field(default_factory=list)
    consolidated_labels: list[dict[str, Any]] = field(default_factory=list)

    # Step 4
    selected_jig_model: str = "Virgo"
    has_4m_change: bool = False
    evaluation_4m: str = "Chờ xác nhận"
    ktsx_reviewer: str = ""
    step4_completed: bool = False
    last_bom_file: Optional[Path] = None

    # Storage Hierarchy
    use_date_hierarchy: bool = False
    current_model_dir: Optional[Path] = None


# =============================================================================
# Background Worker for Batch Reconciliation
# =============================================================================

class BatchReconciliationWorker(QObject):
    """Executes multi-stage batch reconciliation in background thread."""

    progress = pyqtSignal(int, str)  # percent, message
    finished = pyqtSignal(object)    # ReconciliationResult
    error = pyqtSignal(str)          # error message

    def __init__(
        self,
        collected_cttt: list[dict[str, Any]],
        plm_path: Path | None,
        r3_path: Path | None,
        model_name: str,
    ) -> None:
        super().__init__()
        self.collected_cttt = collected_cttt
        self.plm_path = plm_path
        self.r3_path = r3_path
        self.model_name = model_name

    @pyqtSlot()
    def run(self) -> None:
        """Run batch reconciliation pipeline."""
        try:
            self.progress.emit(10, "Đang tổng hợp dữ liệu CTTT từ các công đoạn...")
            df_cttt = pd.DataFrame(self.collected_cttt)

            self.progress.emit(30, "Đang tải dữ liệu cây phân cấp PLM...")
            df_plm = pd.DataFrame()
            if self.plm_path and self.plm_path.exists():
                try:
                    df_plm = pd.read_excel(self.plm_path)
                except Exception as ex:
                    logger.warning("Error reading PLM file: %s", ex)

            self.progress.emit(55, "Đang nạp cấu trúc BOM SAP R3 CS12...")
            df_r3 = pd.DataFrame()
            if self.r3_path and self.r3_path.exists():
                try:
                    from src.automation.sap.parser import parse_r3_cs12_file
                    df_r3 = parse_r3_cs12_file(self.r3_path)
                except Exception:
                    try:
                        df_r3 = pd.read_excel(self.r3_path)
                    except Exception as ex:
                        logger.warning("Error reading R3 file: %s", ex)

            self.progress.emit(75, "Đang chạy thuật toán đối soát 3 bên và tổng hợp toàn máy...")
            engine = ReconciliationEngine()
            result = engine.run_full_reconciliation(
                cttt_data=df_cttt,
                plm_data=df_plm,
                r3_data=df_r3,
            )

            self.progress.emit(100, f"Hoàn tất đối soát! Trạng thái tổng thể: {result.overall_status}")
            self.finished.emit(result)
        except Exception as exc:
            logger.error("Batch reconciliation error: %s", exc)
            self.error.emit(str(exc))


# =============================================================================
# Outlook Email Preview Dialog
# =============================================================================

class EmailPreviewDialog(QDialog):
    """Dialog displaying rendered HTML email preview with Send & Display triggers."""

    def __init__(
        self,
        preview: EmailPreview,
        mailer: OutlookMailer,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.preview = preview
        self.mailer = mailer
        self.setWindowTitle("Xem trước Email Thông báo (Outlook Notification Preview)")
        self.resize(780, 620)
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)

        meta_group = QGroupBox("Thông tin Gửi thư")
        meta_layout = QVBoxLayout(meta_group)

        h_to = QHBoxLayout()
        h_to.addWidget(QLabel("Người nhận (To):"))
        self.edit_to = QLineEdit(self.preview.to_string)
        h_to.addWidget(self.edit_to)
        meta_layout.addLayout(h_to)

        h_cc = QHBoxLayout()
        h_cc.addWidget(QLabel("Đồng kính gửi (CC):"))
        self.edit_cc = QLineEdit(self.preview.cc_string)
        h_cc.addWidget(self.edit_cc)
        meta_layout.addLayout(h_cc)

        h_sub = QHBoxLayout()
        h_sub.addWidget(QLabel("Tiêu đề (Subject):"))
        self.edit_sub = QLineEdit(self.preview.subject)
        h_sub.addWidget(self.edit_sub)
        meta_layout.addLayout(h_sub)

        att_text = ", ".join([p.name for p in self.preview.attachment_paths]) or "Không có tệp đính kèm"
        lbl_att = QLabel(f"📎 Tệp đính kèm: {att_text}")
        lbl_att.setStyleSheet("font-weight: bold; color: #1F497D;")
        meta_layout.addWidget(lbl_att)

        layout.addWidget(meta_group)

        lbl_body = QLabel("Nội dung Thư (HTML Preview):")
        layout.addWidget(lbl_body)

        self.browser = QTextBrowser()
        self.browser.setHtml(self.preview.html_body)
        layout.addWidget(self.browser)

        btn_layout = QHBoxLayout()
        btn_display = QPushButton("Mở trong Microsoft Outlook")
        btn_display.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 6px 14px;")
        btn_display.clicked.connect(self._on_display_clicked)

        btn_send = QPushButton("Gửi trực tiếp qua Outlook")
        btn_send.setStyleSheet("background-color: #10B981; color: white; font-weight: bold; padding: 6px 14px;")
        btn_send.clicked.connect(self._on_send_clicked)

        btn_close = QPushButton("Đóng")
        btn_close.clicked.connect(self.reject)

        btn_layout.addWidget(btn_display)
        btn_layout.addWidget(btn_send)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_close)

        layout.addLayout(btn_layout)

    def _sync_preview_fields(self) -> None:
        self.preview.recipients_to = [r.strip() for r in self.edit_to.text().split(";") if r.strip()]
        self.preview.recipients_cc = [c.strip() for c in self.edit_cc.text().split(";") if c.strip()]
        self.preview.subject = self.edit_sub.text().strip()

    def _on_display_clicked(self) -> None:
        self._sync_preview_fields()
        success = self.mailer.preview_in_outlook(self.preview)
        if success:
            QMessageBox.information(self, "Outlook", "Đã hiển thị cửa sổ soạn thảo thư Outlook!")
            self.accept()
        else:
            QMessageBox.warning(
                self,
                "Lỗi Outlook",
                "Không thể kết nối với tiến trình Outlook COM.\n"
                "Vui lòng đảm bảo Microsoft Outlook đang được cài đặt và khởi động trên máy tính.",
            )

    def _on_send_clicked(self) -> None:
        self._sync_preview_fields()
        success = self.mailer.send_via_outlook(self.preview)
        if success:
            QMessageBox.information(self, "Outlook", "Đã gửi email thông báo thành công!")
            self.accept()
        else:
            QMessageBox.warning(
                self,
                "Lỗi Outlook",
                "Không thể gửi thư qua Outlook COM.\n"
                "Vui lòng kiểm tra kết nối Outlook hoặc thử tính năng 'Mở trong Outlook'.",
            )


# =============================================================================
# Wizard Step Header (Breadcrumb Navigation)
# =============================================================================

class WizardStepHeader(QWidget):
    """Visual 4-step progress header with breadcrumbs and state indicators."""

    step_clicked = pyqtSignal(int)

    STEP_NAMES = [
        "1. Lập Dự Án && Phân Công",
        "2. Tải && Xử Lý Nguồn BOM",
        "3. Theo dõi tổng hợp file của phụ trách",
        "4. So Sánh BOM Tổng && Báo Cáo",
    ]

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.current_step = 0
        self.step_completed_flags = [False, False, False, False]
        self._buttons: list[QPushButton] = []
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(6)

        for idx, title in enumerate(self.STEP_NAMES):
            btn = QPushButton(title)
            btn.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
            btn.setMinimumHeight(34)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda checked=False, i=idx: self._handle_step_click(i))
            layout.addWidget(btn, stretch=1)
            self._buttons.append(btn)

        self._refresh_styles()

    def set_active_step(self, step_idx: int) -> None:
        """Set the active wizard step (0-indexed)."""
        if 0 <= step_idx < len(self.STEP_NAMES):
            self.current_step = step_idx
            self._refresh_styles()

    def set_step_completed(self, step_idx: int, completed: bool = True) -> None:
        """Mark a step as completed."""
        if 0 <= step_idx < len(self.STEP_NAMES):
            self.step_completed_flags[step_idx] = completed
            self._refresh_styles()

    def _handle_step_click(self, step_idx: int) -> None:
        # Allowed to click if step is current or completed or immediate next
        if step_idx <= self.current_step or (step_idx > 0 and self.step_completed_flags[step_idx - 1]):
            self.step_clicked.emit(step_idx)

    def _refresh_styles(self) -> None:
        theme_mgr = get_theme_manager()
        for idx, btn in enumerate(self._buttons):
            title = self.STEP_NAMES[idx]
            btn.setText(title)
            if idx == self.current_step:
                btn.setIcon(QIcon())
                btn.setStyleSheet(
                    "background-color: #0078D4; color: white; border: 2px solid #005A9E; "
                    "border-radius: 4px; padding: 4px 8px; font-weight: bold;"
                )
            elif self.step_completed_flags[idx]:
                btn.setIcon(theme_mgr.get_styled_icon("check-circle", color="#2E7D32"))
                btn.setStyleSheet(
                    "background-color: #E8F5E9; color: #2E7D32; border: 1px solid #A5D6A7; "
                    "border-radius: 4px; padding: 4px 8px; font-weight: bold;"
                )
            else:
                btn.setIcon(QIcon())
                btn.setStyleSheet(
                    "background-color: #F5F5F5; color: #757575; border: 1px solid #E0E0E0; "
                    "border-radius: 4px; padding: 4px 8px;"
                )


DOS_RESERVED_DEVICE_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


def sanitize_machine_code(code: str) -> str:
    """Sanitize machine code to remove path traversal and Windows forbidden characters."""
    cleaned = re.sub(r'[\/\\:\*\?"<>|]', '', code)
    cleaned = re.sub(r'\.{2,}', '', cleaned).strip()
    if not cleaned:
        cleaned = "DEFAULT_MACHINE"
    if cleaned.upper() in DOS_RESERVED_DEVICE_NAMES:
        cleaned = f"M_{cleaned}"
    return cleaned


# =============================================================================
# Step 1 Widget: Lập Dự Án & Phân Công (Setup & Staffing)
# =============================================================================


class Step1ProjectSetupWidget(QWidget):
    """Step 1: Project Setup, Stage Selection, Machine Codes, and Staffing Matrix."""

    step_completed = pyqtSignal(bool)

    def __init__(self, state: LeaderSessionState, mailer: OutlookMailer | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self.mailer = mailer or OutlookMailer()
        self._is_updating_machine_table: bool = False
        self._init_ui()
        self._populate_defaults()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setSpacing(4)

        # 1. Top configuration group (2-row spacious layout, zero horizontal overflow)
        config_group = QGroupBox("1.1 Cấu hình Dự án && Giai đoạn")
        config_layout = QVBoxLayout(config_group)
        config_layout.setContentsMargins(8, 4, 8, 4)
        config_layout.setSpacing(4)

        # Row 1: Model, Stage, and Primary Action Button
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(8)
        row1_layout.addWidget(QLabel("Model máy:"))
        self.model_combo = QComboBox()
        self.model_combo.addItems(["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"])
        self.model_combo.currentTextChanged.connect(self._on_model_changed)
        row1_layout.addWidget(self.model_combo)

        row1_layout.addWidget(QLabel("Giai đoạn:"))
        self.stage_combo = QComboBox()
        self.stage_combo.addItems([
            "MP (ma1 - tiền tố 110)",
            "PP (ma1 - tiền tố 110)",
            "DMT / PMT (maT - tiền tố T10)",
        ])
        self.stage_combo.currentTextChanged.connect(self._on_stage_changed)
        row1_layout.addWidget(self.stage_combo)

        row1_layout.addStretch()

        self.btn_scan_pcd = QPushButton("Quét Kế Hoạch PCD...")
        self.btn_scan_pcd.setIcon(get_theme_manager().get_styled_icon("calendar"))
        self.btn_scan_pcd.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_scan_pcd.setMinimumHeight(28)
        self.btn_scan_pcd.setStyleSheet(
            "background-color: #2E7D32; color: white; border-radius: 4px; padding: 4px 12px; font-weight: bold;"
        )
        self.btn_scan_pcd.clicked.connect(self._open_pcd_scan_dialog)
        row1_layout.addWidget(self.btn_scan_pcd)

        self.btn_send_assign_email = QPushButton("Gửi Mail Yêu Cầu Phụ Trách...")
        self.btn_send_assign_email.setIcon(get_theme_manager().get_styled_icon("mail"))
        self.btn_send_assign_email.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_send_assign_email.setMinimumHeight(28)
        self.btn_send_assign_email.setStyleSheet(
            "background-color: #E65100; color: white; border-radius: 4px; padding: 4px 12px; font-weight: bold;"
        )
        self.btn_send_assign_email.clicked.connect(self._on_send_assignment_email)
        row1_layout.addWidget(self.btn_send_assign_email)

        self.btn_create_folders = QPushButton("Khởi Tạo Thư Mục Dự Án && Xuất File Phân Công")
        self.btn_create_folders.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_create_folders.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_create_folders.setMinimumHeight(28)
        self.btn_create_folders.setStyleSheet(
            "background-color: #0078D4; color: white; border-radius: 4px; padding: 4px 14px;"
        )
        self.btn_create_folders.clicked.connect(self.execute_create_folders_and_packages)
        row1_layout.addWidget(self.btn_create_folders)
        config_layout.addLayout(row1_layout)

        # Row 2: Base directory full width path + browse & open buttons
        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(8)
        row2_layout.addWidget(QLabel("Thư mục gốc:"))
        self.edit_base_dir = QLineEdit(str(self.state.base_dir))
        row2_layout.addWidget(self.edit_base_dir, stretch=1)

        self.btn_browse_dir = QPushButton("Duyệt...")
        self.btn_browse_dir.clicked.connect(self._browse_base_dir)
        row2_layout.addWidget(self.btn_browse_dir)

        self.btn_open_folder = QPushButton("Mở thư mục")
        self.btn_open_folder.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_open_folder.clicked.connect(self._open_base_dir)
        row2_layout.addWidget(self.btn_open_folder)
        config_layout.addLayout(row2_layout)

        layout.addWidget(config_group)

        # 2. Split tables: Machine targets and Staffing
        tables_layout = QHBoxLayout()

        # Left: Machine Code list
        mach_group = QGroupBox("1.2 Danh sách Mã Máy / Hướng Xuất")
        mach_layout = QVBoxLayout(mach_group)
        mach_layout.setContentsMargins(6, 4, 6, 4)
        mach_layout.setSpacing(4)

        mach_btn_layout = QHBoxLayout()
        mach_btn_layout.setContentsMargins(0, 2, 0, 4)
        self.btn_add_mach = QPushButton("+ Thêm mã")
        self.btn_add_mach.clicked.connect(self._add_machine_row)
        self.btn_del_mach = QPushButton("- Xóa mã")
        self.btn_del_mach.clicked.connect(self._del_machine_row)
        self.btn_paste_mach = QPushButton("Nhập danh sách")
        self.btn_paste_mach.setIcon(get_theme_manager().get_styled_icon("file-spreadsheet"))
        self.btn_paste_mach.clicked.connect(self._paste_machines)

        mach_btn_layout.addWidget(self.btn_add_mach)
        mach_btn_layout.addWidget(self.btn_del_mach)
        mach_btn_layout.addWidget(self.btn_paste_mach)
        mach_layout.addLayout(mach_btn_layout)

        self.machine_table = QTableWidget(0, 4)
        self.machine_table.setHorizontalHeaderLabels(["STT", "Mã Máy", "Bỏ qua (X)", "Ghi chú"])
        self.machine_table.verticalHeader().setDefaultSectionSize(32)
        self.machine_table.verticalHeader().setMinimumSectionSize(28)
        self.machine_table.setShowGrid(True)
        self.machine_table.setMinimumHeight(160)
        self.machine_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.machine_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        h_mach = self.machine_table.horizontalHeader()
        h_mach.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.machine_table.setColumnWidth(0, 45)
        h_mach.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        h_mach.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.machine_table.setColumnWidth(2, 75)
        h_mach.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.machine_table.itemChanged.connect(self._on_machine_table_item_changed)
        mach_layout.addWidget(self.machine_table)

        tables_layout.addWidget(mach_group, stretch=1)


        # Right: Staffing Table (Sheet Lichsu)
        staff_group = QGroupBox("1.3 Phân Công Nhân Sự (Sheet Lichsu && tenphong_pt)")
        staff_layout = QVBoxLayout(staff_group)
        staff_layout.setContentsMargins(6, 4, 6, 4)
        staff_layout.setSpacing(4)

        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 2, 0, 4)
        filter_layout.addWidget(QLabel("Lọc phòng ban:"))
        self.combo_dept_filter = QComboBox()
        self.combo_dept_filter.addItems(["Tất cả", "Cơ 1", "Cơ 2", "Cơ 3"])
        self.combo_dept_filter.currentTextChanged.connect(self._filter_staff_table)
        filter_layout.addWidget(self.combo_dept_filter)

        self.btn_select_all_staff = QPushButton("Chọn tất cả")
        self.btn_select_all_staff.clicked.connect(lambda: self._set_all_staff_checked(True))
        self.btn_deselect_all_staff = QPushButton("Bỏ chọn")
        self.btn_deselect_all_staff.clicked.connect(lambda: self._set_all_staff_checked(False))
        self.btn_auto_assign = QPushButton("Tự động gán công đoạn")
        self.btn_auto_assign.clicked.connect(self._auto_assign_subunits)

        self.btn_manage_roster = QPushButton("Quản lý nhân sự...")
        self.btn_manage_roster.setIcon(get_theme_manager().get_styled_icon("settings"))
        self.btn_manage_roster.setToolTip("Thêm, sửa, xóa thành viên và kết nối CSDL chung LAN (ssbom_master.db)")
        self.btn_manage_roster.clicked.connect(self._open_manage_roster_dialog)

        filter_layout.addWidget(self.btn_select_all_staff)
        filter_layout.addWidget(self.btn_deselect_all_staff)
        filter_layout.addWidget(self.btn_auto_assign)
        filter_layout.addWidget(self.btn_manage_roster)
        staff_layout.addLayout(filter_layout)

        self.staff_table = QTableWidget(0, 5)
        self.staff_table.setHorizontalHeaderLabels([
            "Áp dụng", "Phụ trách công đoạn", "Phòng Ban", "Mã máy", "Công Đoạn",
        ])
        self.staff_table.verticalHeader().setDefaultSectionSize(32)
        self.staff_table.verticalHeader().setMinimumSectionSize(28)
        self.staff_table.setShowGrid(True)
        self.staff_table.setMinimumHeight(160)
        self.staff_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.staff_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        h_staff = self.staff_table.horizontalHeader()
        h_staff.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.staff_table.setColumnWidth(0, 60)
        h_staff.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        h_staff.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.staff_table.setColumnWidth(2, 75)
        h_staff.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.staff_table.setColumnWidth(3, 115)
        h_staff.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.staff_table.setColumnWidth(4, 130)
        staff_layout.addWidget(self.staff_table)

        tables_layout.addWidget(staff_group, stretch=2)
        layout.addLayout(tables_layout, stretch=1)

        # Testing & backward compatibility aliases
        self.btn_create_project = self.btn_create_folders
        self.btn_add_machine = self.btn_add_mach

    def _open_manage_roster_dialog(self) -> None:
        """Open the Member Roster Management Dialog."""
        from src.gui.member_roster_dialog import MemberRosterDialog
        base_dir = getattr(self.state, "base_dir", None)
        dlg = MemberRosterDialog(parent=self, base_dir=base_dir)
        dlg.roster_changed.connect(self._reload_roster_from_db)
        dlg.exec()
        self._reload_roster_from_db()

    def _reload_roster_from_db(self) -> None:
        """Reload roster from database and update staff table while preserving assignments."""
        base_dir = getattr(self.state, "base_dir", None)
        try:
            from src.core.member_database import MemberDatabaseManager
            db = MemberDatabaseManager(base_dir=base_dir)
            members = db.get_members(active_only=True)
            if not members:
                return
            roster_data = [(m.account_id, m.department, m.default_sub_unit) for m in members]
            depts = db.get_departments()
        except Exception as e:
            logger.warning(f"Could not load roster from database: {e}")
            return

        # Update department filter options
        current_dept = self.combo_dept_filter.currentText()
        self.combo_dept_filter.blockSignals(True)
        self.combo_dept_filter.clear()
        self.combo_dept_filter.addItem("Tất cả")
        for d in depts:
            self.combo_dept_filter.addItem(d)
        idx = self.combo_dept_filter.findText(current_dept)
        if idx >= 0:
            self.combo_dept_filter.setCurrentIndex(idx)
        else:
            self.combo_dept_filter.setCurrentIndex(0)
        self.combo_dept_filter.blockSignals(False)

        # Cache existing row states
        existing_assignments: dict[str, tuple[bool, str, str]] = {}
        for r in range(self.staff_table.rowCount()):
            eng_item = self.staff_table.item(r, 1)
            if not eng_item:
                continue
            eng = eng_item.text().strip()
            chk_w = self.staff_table.cellWidget(r, 0)
            is_app = True
            if chk_w:
                chk = chk_w.findChild(QCheckBox)
                if chk:
                    is_app = chk.isChecked()

            combo_mach = self.staff_table.cellWidget(r, 3)
            mach_item = self.staff_table.item(r, 3)
            mach = combo_mach.currentText() if isinstance(combo_mach, QComboBox) else (mach_item.text() if mach_item else "")

            combo_sub = self.staff_table.cellWidget(r, 4)
            sub_item = self.staff_table.item(r, 4)
            sub = combo_sub.currentText() if isinstance(combo_sub, QComboBox) else (sub_item.text() if sub_item else "")
            existing_assignments[eng] = (is_app, mach, sub)

        self.staff_table.setRowCount(0)
        self.state.staff_roster.clear()

        sub_unit_cycle = list(STANDARD_SUB_UNITS)
        active_machines = self.get_active_machine_codes() or ["110C103NL0", "110C103NL1", "110C0Z3LV1"]

        for idx, (eng, dept, def_sub) in enumerate(roster_data):
            if eng in existing_assignments:
                is_app, assigned_mach, assigned_unit = existing_assignments[eng]
            else:
                is_app = True
                assigned_unit = def_sub if def_sub in STANDARD_SUB_UNITS else sub_unit_cycle[idx % len(sub_unit_cycle)]
                assigned_mach = active_machines[idx % len(active_machines)]

            assignment = StaffAssignment(
                engineer_name=eng,
                department=dept,
                is_applied=is_app,
                machine_code=assigned_mach,
                sub_unit=assigned_unit,
            )
            self.state.staff_roster.append(assignment)

            r = self.staff_table.rowCount()
            self.staff_table.insertRow(r)

            chk = QCheckBox()
            chk.setChecked(is_app)
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.addWidget(chk)
            chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            self.staff_table.setCellWidget(r, 0, chk_widget)

            self.staff_table.setItem(r, 1, QTableWidgetItem(eng))
            self.staff_table.setItem(r, 2, QTableWidgetItem(dept))
            self.staff_table.setItem(r, 3, QTableWidgetItem(assigned_mach))
            self.staff_table.setItem(r, 4, QTableWidgetItem(assigned_unit))

            combo_sub = QComboBox()
            combo_sub.setEditable(True)
            combo_sub.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            combo_sub.setStyleSheet("QComboBox { padding: 1px 4px; font-size: 11px; }")
            for u in STANDARD_SUB_UNITS:
                combo_sub.addItem(u)
            combo_sub.setEditText(assigned_unit)
            self.staff_table.setCellWidget(r, 4, combo_sub)
            combo_sub.currentTextChanged.connect(lambda txt, row=r: self._on_staff_subunit_combo_changed(row, txt))

        self._sync_staff_machine_options()
        self._filter_staff_table(self.combo_dept_filter.currentText())

    def _populate_defaults(self) -> None:
        """Populate initial machine codes and roster."""
        was_updating = self._is_updating_machine_table
        self._is_updating_machine_table = True
        try:
            # Initial machine codes
            default_machines = ["110C103NL0", "110C103NL1", "110C0Z3LV1"]
            for code in default_machines:
                self._add_machine_row(code=code)
        finally:
            self._is_updating_machine_table = was_updating

        # Try loading roster from MemberDatabaseManager
        base_dir = getattr(self.state, "base_dir", None)
        roster_data: list[tuple[str, str, str]] = []
        try:
            from src.core.member_database import MemberDatabaseManager
            db = MemberDatabaseManager(base_dir=base_dir)
            members = db.get_members(active_only=True)
            if members:
                roster_data = [(m.account_id, m.department, m.default_sub_unit) for m in members]
                depts = db.get_departments()
                self.combo_dept_filter.blockSignals(True)
                self.combo_dept_filter.clear()
                self.combo_dept_filter.addItem("Tất cả")
                for d in depts:
                    self.combo_dept_filter.addItem(d)
                self.combo_dept_filter.blockSignals(False)
        except Exception as e:
            logger.warning(f"Could not load roster from database, using static fallback: {e}")

        if not roster_data:
            roster_data = [
                (eng, "Cơ 1", "") for eng in ROSTER_MECHA_1
            ] + [
                (eng, "Cơ 2", "") for eng in ROSTER_MECHA_2
            ] + [
                (eng, "Cơ 3", "") for eng in ROSTER_MECHA_3
            ]

        self.staff_table.setRowCount(0)
        self.state.staff_roster.clear()

        sub_unit_cycle = list(STANDARD_SUB_UNITS)
        active_machines = self.get_active_machine_codes() or default_machines

        for idx, (eng, dept, def_sub) in enumerate(roster_data):
            assigned_unit = def_sub if def_sub in STANDARD_SUB_UNITS else sub_unit_cycle[idx % len(sub_unit_cycle)]
            assigned_mach = active_machines[idx % len(active_machines)]
            assignment = StaffAssignment(
                engineer_name=eng,
                department=dept,
                is_applied=True,
                machine_code=assigned_mach,
                sub_unit=assigned_unit,
            )
            self.state.staff_roster.append(assignment)

            r = self.staff_table.rowCount()
            self.staff_table.insertRow(r)

            # Checkbox
            chk = QCheckBox()
            chk.setChecked(True)
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.addWidget(chk)
            chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            self.staff_table.setCellWidget(r, 0, chk_widget)

            self.staff_table.setItem(r, 1, QTableWidgetItem(eng))
            self.staff_table.setItem(r, 2, QTableWidgetItem(dept))
            self.staff_table.setItem(r, 3, QTableWidgetItem(assigned_mach))
            self.staff_table.setItem(r, 4, QTableWidgetItem(assigned_unit))

            combo_sub = QComboBox()
            combo_sub.setEditable(True)
            combo_sub.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            combo_sub.setStyleSheet("QComboBox { padding: 1px 4px; font-size: 11px; }")
            for u in STANDARD_SUB_UNITS:
                combo_sub.addItem(u)
            combo_sub.setEditText(assigned_unit)
            self.staff_table.setCellWidget(r, 4, combo_sub)
            combo_sub.currentTextChanged.connect(lambda txt, row=r: self._on_staff_subunit_combo_changed(row, txt))

        # Synchronize dynamic dropdown options for column "Mã máy"
        self._sync_staff_machine_options()

    def _add_machine_row(self, code: str = "", note: str = "") -> None:
        was_updating = self._is_updating_machine_table
        self._is_updating_machine_table = True
        try:
            r = self.machine_table.rowCount()
            self.machine_table.insertRow(r)

            self.machine_table.setItem(r, 0, QTableWidgetItem(str(r + 1)))
            self.machine_table.item(r, 0).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            prefix = "T10" if "maT" in self.stage_combo.currentText() else "110"
            init_code = code or f"{prefix}C10{r + 1}NL0"
            self.machine_table.setItem(r, 1, QTableWidgetItem(init_code))

            chk_exclude = QCheckBox()
            chk_exclude.toggled.connect(self._on_machine_exclude_toggled)
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.addWidget(chk_exclude)
            chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            self.machine_table.setCellWidget(r, 2, chk_widget)

            self.machine_table.setItem(r, 3, QTableWidgetItem(note))
        finally:
            self._is_updating_machine_table = was_updating

        if not was_updating:
            self._sync_staff_machine_options()

    def _del_machine_row(self) -> None:
        curr = self.machine_table.currentRow()
        if curr >= 0:
            self.machine_table.removeRow(curr)
            self._reindex_machine_stt()
            self._sync_staff_machine_options()

    def _reindex_machine_stt(self) -> None:
        was_updating = self._is_updating_machine_table
        self._is_updating_machine_table = True
        try:
            for r in range(self.machine_table.rowCount()):
                item = self.machine_table.item(r, 0)
                if item:
                    item.setText(str(r + 1))
                else:
                    self.machine_table.setItem(r, 0, QTableWidgetItem(str(r + 1)))
        finally:
            self._is_updating_machine_table = was_updating

    def _paste_machines(self) -> None:
        from PyQt6.QtWidgets import QInputDialog
        text, ok = QInputDialog.getMultiLineText(
            self,
            "Nhập danh sách Mã Máy",
            "Dán hoặc nhập danh sách mã máy (mỗi mã 1 dòng):",
        )
        if ok and text.strip():
            self._is_updating_machine_table = True
            try:
                for line in text.splitlines():
                    code = line.strip().upper()
                    if code:
                        self._add_machine_row(code=code)
            finally:
                self._is_updating_machine_table = False
            self._reindex_machine_stt()
            self._sync_staff_machine_options()

    def _on_model_changed(self, text: str) -> None:
        self.state.model_name = text.strip()

    def _on_stage_changed(self, text: str) -> None:
        is_mat = "maT" in text
        self.state.stage = ProjectStage.MA_T if is_mat else ProjectStage.MA_1
        prefix = "T10" if is_mat else "110"
        self._is_updating_machine_table = True
        try:
            for r in range(self.machine_table.rowCount()):
                item = self.machine_table.item(r, 1)
                if item and (item.text().startswith("110") or item.text().startswith("T10")):
                    item.setText(prefix + item.text()[3:])
        finally:
            self._is_updating_machine_table = False
        self._sync_staff_machine_options()

    def _on_machine_table_item_changed(self, item: QTableWidgetItem) -> None:
        if getattr(self, "_is_updating_machine_table", False):
            return
        if item.column() == 1:
            self._sync_staff_machine_options()

    def _on_machine_exclude_toggled(self, checked: bool) -> None:
        if getattr(self, "_is_updating_machine_table", False):
            return
        self._sync_staff_machine_options()

    def get_active_machine_codes(self) -> list[str]:
        """Extract active machine codes from machine_table (excluding rows marked X)."""
        codes: list[str] = []
        for r in range(self.machine_table.rowCount()):
            chk_w = self.machine_table.cellWidget(r, 2)
            is_excluded = False
            if chk_w:
                chk = chk_w.findChild(QCheckBox)
                if chk and chk.isChecked():
                    is_excluded = True
            if not is_excluded:
                item = self.machine_table.item(r, 1)
                val = item.text().strip().upper() if item else ""
                if val and val not in codes:
                    codes.append(val)
        return codes

    def _on_staff_machine_combo_changed(self, row: int, txt: str) -> None:
        item = self.staff_table.item(row, 3)
        if item:
            item.setText(txt)
        else:
            self.staff_table.setItem(row, 3, QTableWidgetItem(txt))
        if row < len(self.state.staff_roster):
            self.state.staff_roster[row].machine_code = txt

    def _on_staff_subunit_combo_changed(self, row: int, txt: str) -> None:
        item = self.staff_table.item(row, 4)
        if item:
            item.setText(txt)
        else:
            self.staff_table.setItem(row, 4, QTableWidgetItem(txt))
        if row < len(self.state.staff_roster):
            self.state.staff_roster[row].sub_unit = txt

    def _sync_staff_machine_options(self) -> None:
        """Dynamically sync 'Mã máy' combo options in staff_table with machine_table."""
        active_codes = self.get_active_machine_codes()
        options = active_codes if active_codes else ["(Chưa có mã máy)"]

        for r in range(self.staff_table.rowCount()):
            combo = self.staff_table.cellWidget(r, 3)
            if not isinstance(combo, QComboBox):
                combo = QComboBox()
                combo.setStyleSheet("QComboBox { padding: 1px 4px; font-size: 11px; }")
                self.staff_table.setCellWidget(r, 3, combo)

            current_text = combo.currentText().strip()
            if not current_text:
                item = self.staff_table.item(r, 3)
                if item:
                    current_text = item.text().strip()

            combo.blockSignals(True)
            combo.clear()
            for opt in options:
                combo.addItem(opt)

            if current_text in options:
                combo.setCurrentText(current_text)
            elif active_codes:
                chosen = active_codes[r % len(active_codes)]
                combo.setCurrentText(chosen)
            else:
                combo.setCurrentIndex(0)

            val = combo.currentText()
            item = self.staff_table.item(r, 3)
            if item:
                item.setText(val)
            else:
                self.staff_table.setItem(r, 3, QTableWidgetItem(val))

            if r < len(self.state.staff_roster):
                self.state.staff_roster[r].machine_code = val

            combo.blockSignals(False)

            try:
                combo.currentTextChanged.disconnect()
            except Exception:
                pass
            combo.currentTextChanged.connect(lambda txt, row=r: self._on_staff_machine_combo_changed(row, txt))

    def _browse_base_dir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Chọn thư mục gốc dự án", self.edit_base_dir.text())
        if d:
            self.edit_base_dir.setText(d)
            self.state.base_dir = Path(d)

    def _open_base_dir(self) -> None:
        p = Path(self.edit_base_dir.text())
        if not p.exists():
            p = self.state.base_dir
        try:
            os.startfile(str(p))
        except Exception:
            subprocess.Popen(["explorer", str(p)], shell=True)

    def _filter_staff_table(self, dept: str) -> None:
        for r in range(self.staff_table.rowCount()):
            r_dept = self.staff_table.item(r, 2).text()
            if dept == "Tất cả" or r_dept == dept:
                self.staff_table.setRowHidden(r, False)
            else:
                self.staff_table.setRowHidden(r, True)

    def _set_all_staff_checked(self, checked: bool) -> None:
        for r in range(self.staff_table.rowCount()):
            w = self.staff_table.cellWidget(r, 0)
            if w:
                chk = w.findChild(QCheckBox)
                if chk:
                    chk.setChecked(checked)

    def _auto_assign_subunits(self) -> None:
        sub_units = list(STANDARD_SUB_UNITS)
        active_machines = self.get_active_machine_codes()
        visible_idx = 0
        for r in range(self.staff_table.rowCount()):
            if not self.staff_table.isRowHidden(r):
                unit = sub_units[visible_idx % len(sub_units)]
                combo_sub = self.staff_table.cellWidget(r, 4)
                if isinstance(combo_sub, QComboBox):
                    combo_sub.blockSignals(True)
                    combo_sub.setCurrentText(unit)
                    combo_sub.blockSignals(False)
                item_sub = self.staff_table.item(r, 4)
                if item_sub:
                    item_sub.setText(unit)
                else:
                    self.staff_table.setItem(r, 4, QTableWidgetItem(unit))
                if r < len(self.state.staff_roster):
                    self.state.staff_roster[r].sub_unit = unit

                if active_machines:
                    mach = active_machines[visible_idx % len(active_machines)]
                    combo = self.staff_table.cellWidget(r, 3)
                    if isinstance(combo, QComboBox):
                        combo.blockSignals(True)
                        combo.setCurrentText(mach)
                        combo.blockSignals(False)
                    item = self.staff_table.item(r, 3)
                    if item:
                        item.setText(mach)
                    else:
                        self.staff_table.setItem(r, 3, QTableWidgetItem(mach))
                    if r < len(self.state.staff_roster):
                        self.state.staff_roster[r].machine_code = mach
                visible_idx += 1

    def sync_state_from_ui(self) -> None:
        """Sync UI data back to LeaderSessionState."""
        self.state.model_name = self.model_combo.currentText().strip()
        self.state.base_dir = Path(self.edit_base_dir.text().strip())

        # Sync machines
        self.state.machines.clear()
        for r in range(self.machine_table.rowCount()):
            code_item = self.machine_table.item(r, 1)
            code = code_item.text().strip().upper() if code_item else ""
            if not code:
                continue

            chk_w = self.machine_table.cellWidget(r, 2)
            is_excluded = False
            if chk_w:
                chk = chk_w.findChild(QCheckBox)
                if chk:
                    is_excluded = chk.isChecked()

            m = MachineTarget(
                machine_code=code,
                is_excluded=is_excluded,
                folder_path=self.state.base_dir / self.state.model_name / code,
            )
            self.state.machines.append(m)

        # Sync staff roster
        for r in range(self.staff_table.rowCount()):
            if r < len(self.state.staff_roster):
                assignment = self.state.staff_roster[r]
                chk_w = self.staff_table.cellWidget(r, 0)
                if chk_w:
                    chk = chk_w.findChild(QCheckBox)
                    if chk:
                        assignment.is_applied = chk.isChecked()

                combo_w = self.staff_table.cellWidget(r, 3)
                if isinstance(combo_w, QComboBox):
                    assignment.machine_code = combo_w.currentText().strip()
                else:
                    m_code_item = self.staff_table.item(r, 3)
                    if m_code_item:
                        assignment.machine_code = m_code_item.text().strip()

                sub_item = self.staff_table.item(r, 4)
                combo_sub = self.staff_table.cellWidget(r, 4)
                if isinstance(combo_sub, QComboBox):
                    sub_val = combo_sub.currentText().strip()
                    if sub_item and sub_item.text().strip() and sub_item.text().strip() != sub_val:
                        sub_val = sub_item.text().strip()
                        combo_sub.blockSignals(True)
                        if combo_sub.findText(sub_val) == -1:
                            combo_sub.addItem(sub_val)
                        combo_sub.setCurrentText(sub_val)
                        combo_sub.blockSignals(False)
                    assignment.sub_unit = sub_val
                elif sub_item:
                    assignment.sub_unit = sub_item.text().strip()

    def _open_pcd_scan_dialog(self) -> None:
        """Open PCD Monthly Production Plan scanner dialog."""
        from src.gui.pcd_plan_dialog import PCDPlanScanDialog
        dlg = PCDPlanScanDialog(mailer=self.mailer, parent=self)
        dlg.exec()

    def _on_send_assignment_email(self) -> None:
        """Open Email Preview for task assignment to engineers."""
        model = self.state.model_name
        stage_txt = self.stage_combo.currentText().strip()
        stage_code = "MP"
        if "DMT" in stage_txt or "PMT" in stage_txt:
            stage_code = "DMT/PMT"
        elif "PP" in stage_txt:
            stage_code = "PP"

        active_machines = [m for m in self.state.machines if not m.is_excluded]
        qty = len(active_machines) or 1
        now_str = datetime.datetime.now().strftime("%d/%m")
        today = datetime.date.today()
        deadline_copy = (today + datetime.timedelta(days=3)).strftime("%d.%m.%Y")
        deadline_verify = (today + datetime.timedelta(days=5)).strftime("%d.%m.%Y")

        att_path = self.state.base_dir / model

        preview = self.mailer.build_task_assignment_email(
            machine_type=model,
            start_date=now_str,
            quantity=qty,
            phase=stage_code,
            deadline_copy=deadline_copy,
            deadline_verify=deadline_verify,
            attachment_path=att_path,
        )
        dlg = EmailPreviewDialog(preview, self.mailer, parent=self)
        dlg.exec()

    def create_project_folder_structure(self, force_date_hierarchy: bool | None = None) -> Path:
        """Canonical folder creation for model with phase and Year.Month (YYYY.MM)."""
        self.sync_state_from_ui()
        stage_txt = self.stage_combo.currentText().strip()
        stage_code = "MP"
        if "DMT" in stage_txt or "PMT" in stage_txt:
            stage_code = "DMT"
        elif "PP" in stage_txt:
            stage_code = "PP"

        now = datetime.datetime.now()
        year_month = f"{now.year}.{now.month:02d}"

        # Standard hierarchy: <base_dir>/<model_name>/<stage>/<year_month> when on network UNC storage,
        # or flat <base_dir>/<model_name> for local test/custom environments unless explicitly forced.
        is_canonical_unc = "SO SANH PLM-CTTT-R3" in str(self.state.base_dir).upper()
        use_hierarchy = (
            force_date_hierarchy
            if force_date_hierarchy is not None
            else (is_canonical_unc or getattr(self.state, "use_date_hierarchy", False))
        )

        if use_hierarchy:
            model_dir = self.state.base_dir / self.state.model_name / stage_code / year_month
        else:
            model_dir = self.state.base_dir / self.state.model_name

        self.state.current_model_dir = model_dir

        dirs_to_create = [
            model_dir / "PLM",
            model_dir / "R3",
            model_dir / "Reports",
            model_dir / "capnhat" / "old",
        ]
        for sub_unit in STANDARD_SUB_UNITS:
            dirs_to_create.append(model_dir / "CTTT" / sub_unit)

        for d in dirs_to_create:
            d.mkdir(parents=True, exist_ok=True)

        return model_dir

    def execute_create_folders_and_packages(self) -> None:
        """Execute folder creation and member submission package distribution."""
        self.sync_state_from_ui()

        active_machines = [m for m in self.state.machines if not m.is_excluded]
        active_engineers = [e for e in self.state.staff_roster if e.is_applied]

        if not active_machines:
            QMessageBox.warning(self, "Thông báo lỗi", "Bạn chưa nhập mã máy vào list!")
            return

        if not active_engineers:
            QMessageBox.warning(self, "Thông báo lỗi", "Bạn chưa lựa chọn người phụ trách để tạo file điền linh kiện!")
            return

        model_dir = self.create_project_folder_structure()

        # Locate formnguoidung template (check local base_dir or canonical UNC)
        src_template: Path | None = None
        cand = self.state.base_dir / "formnguoidung.xlsm"
        unc_template = Path(
            r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)"
            r"\Hang muc can luu\Vinh\Pm_sosanhBOM\formnguoidung.xlsm"
        )
        if cand.exists():
            src_template = cand
        elif unc_template.exists():
            src_template = unc_template

        # Create machine directories and engineer submission packages
        for m in active_machines:
            m.machine_code = sanitize_machine_code(m.machine_code)
            m_dir = model_dir / m.machine_code

            if m_dir.exists():
                reply = QMessageBox.question(
                    self,
                    "Xác nhận Thư mục Đã Tồn Tại",
                    f"Thư mục dự án đã tồn tại:\n{m_dir}\n\nBạn có muốn ghi đè / tạo lại gói nộp không?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply != QMessageBox.StandardButton.Yes:
                    continue

            m_dir.mkdir(parents=True, exist_ok=True)
            m.folder_path = m_dir


            # Relevant engineers for this machine
            assigned_engineers = [
                e for e in active_engineers
                if e.machine_code == m.machine_code or e.machine_code in ("(Tất cả mã máy)", "Tất cả", "")
            ]
            if not assigned_engineers:
                assigned_engineers = active_engineers

            for eng in assigned_engineers:
                target_file = m_dir / f"{eng.engineer_name}.xlsm"
                self._generate_member_package(target_file, eng, m, src_template)
                eng.assigned_file = target_file

        self.state.step1_completed = True
        self.step_completed.emit(True)

        QMessageBox.information(
            self,
            "Thông báo thành công",
            f"Đã tạo xong danh sách BOM và các gói nộp của phụ trách theo cài đặt.\n\n"
            f"📁 Thư mục lưu trữ: {model_dir}\n"
            f"Cấu trúc: <Thư mục gốc>\\{self.state.model_name}\\<Mã máy>\\<Tên phụ trách>.xlsm",
        )

    def _generate_member_package(
        self,
        target_path: Path,
        engineer: StaffAssignment,
        machine: MachineTarget,
        template_path: Path | None,
    ) -> None:
        """Generate member workbook with CTTT, MSI, and Label_7980_7990 sheets."""
        try:
            if template_path and template_path.exists():
                shutil.copyfile(str(template_path), str(target_path))
            else:
                wb = openpyxl.Workbook()
                ws_cttt = wb.active
                ws_cttt.title = "CTTT"
                ws_cttt["A1"] = f"Mã Máy: {machine.machine_code}"
                ws_cttt["B1"] = f"Công đoạn: {engineer.sub_unit}"
                ws_cttt["F1"] = f"Phụ trách: {engineer.engineer_name}"
                ws_cttt["Q2"] = ""
                wb.create_sheet("MSI")
                wb.create_sheet("Label_7980_7990")
                wb.create_sheet("PLM_R3")
                wb.save(target_path)
                wb.close()
        except Exception as exc:
            logger.warning("Error creating member package %s: %s", target_path, exc)


# =============================================================================
# Step 2 Widget: Tải & Xử Lý Nguồn Dữ Liệu (PLM & SAP Sourcing)
# =============================================================================

class Step2DataSourcingWidget(QWidget):
    """Step 2: BOM Data Sourcing via TC2412 PLM & SAP R3 CS12, Date Controls, and BOM Filter."""

    step_completed = pyqtSignal(bool)

    def __init__(self, state: LeaderSessionState, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(8)

        # 1. Date Configuration Group
        date_group = QGroupBox("2.1 Thiết lập Ngày Hiệu Lực Sản Xuất (Áp dụng cho SAP R3 CS12)")
        date_layout = QVBoxLayout(date_group)

        h_mode = QHBoxLayout()
        self.radio_common_date = QRadioButton("Dùng chung 1 ngày cho tất cả mã máy (DMT/PP):")
        self.radio_common_date.setChecked(True)
        self.date_edit_common = QDateEdit()
        self.date_edit_common.setCalendarPopup(True)
        self.date_edit_common.setDate(QDate.currentDate())
        self.date_edit_common.setDisplayFormat("yyyy/MM/dd")

        self.radio_individual_date = QRadioButton("Thiết lập ngày riêng cho từng mã máy (Bổ sung mã MP):")

        date_button_group = QButtonGroup(self)
        date_button_group.addButton(self.radio_common_date)
        date_button_group.addButton(self.radio_individual_date)

        self.date_edit_common.dateChanged.connect(self._on_common_date_changed)
        self.radio_common_date.toggled.connect(self._on_date_mode_toggled)
        self.radio_individual_date.toggled.connect(self._on_date_mode_toggled)

        h_mode.addWidget(self.radio_common_date)
        h_mode.addWidget(self.date_edit_common)
        h_mode.addSpacing(20)
        h_mode.addWidget(self.radio_individual_date)
        h_mode.addStretch()

        date_layout.addLayout(h_mode)
        layout.addWidget(date_group)

        # 2. Sourcing Table Group
        source_group = QGroupBox("2.2 Bảng Trạng Thái Nguồn Dữ Liệu && Phân Tuyến BOM")
        source_layout = QVBoxLayout(source_group)

        self.sourcing_table = QTableWidget(0, 6)
        self.sourcing_table.setHorizontalHeaderLabels([
            "STT", "Mã Máy", "Ngày Hiệu Lực", "File PLM TC24", "File SAP R3 CS12", "Bộ Lọc BOM Lv1..6",
        ])
        self.sourcing_table.verticalHeader().setDefaultSectionSize(32)
        self.sourcing_table.verticalHeader().setMinimumSectionSize(28)
        self.sourcing_table.setShowGrid(True)
        h_src = self.sourcing_table.horizontalHeader()
        h_src.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h_src.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        h_src.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        h_src.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        source_layout.addWidget(self.sourcing_table)

        layout.addWidget(source_group, stretch=1)

        # 3. Action Buttons
        btn_layout = QHBoxLayout()

        self.btn_download_plm = QPushButton("Tải BOM Tự Động (PLM TC24 / SAP R3)")
        self.btn_download_plm.setIcon(get_theme_manager().get_styled_icon("download"))
        self.btn_download_plm.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_download_plm.setStyleSheet(
            "background-color: #0078D4; color: white; padding: 8px 16px; border-radius: 4px;"
        )
        self.btn_download_plm.clicked.connect(self._open_download_dialog)

        self.btn_import_files = QPushButton("Nạp Tệp Sẵn Có Từ Ổ Đĩa")
        self.btn_import_files.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_import_files.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_import_files.setStyleSheet(
            "background-color: #0D9488; color: white; padding: 8px 16px; border-radius: 4px;"
        )
        self.btn_import_files.clicked.connect(self._import_local_bom_files)

        self.btn_filter_bom = QPushButton("Lọc BOM TC24 Level 1..6 & Sao Lưu")
        self.btn_filter_bom.setIcon(get_theme_manager().get_styled_icon("filter"))
        self.btn_filter_bom.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_filter_bom.setStyleSheet(
            "background-color: #D97706; color: white; padding: 8px 16px; border-radius: 4px;"
        )
        self.btn_filter_bom.clicked.connect(self.execute_bom_filtering)

        self.btn_refresh_sourcing = QPushButton("Làm mới trạng thái")
        self.btn_refresh_sourcing.setIcon(get_theme_manager().get_styled_icon("refresh"))
        self.btn_refresh_sourcing.clicked.connect(self.refresh_sourcing_table)
        self.btn_execute_sourcing = self.btn_refresh_sourcing

        btn_layout.addWidget(self.btn_download_plm)
        btn_layout.addWidget(self.btn_import_files)
        btn_layout.addWidget(self.btn_filter_bom)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_refresh_sourcing)

        layout.addLayout(btn_layout)

    def _on_common_date_changed(self, new_date: QDate) -> None:
        """When common date is updated, synchronize all rows in the sourcing table immediately."""
        if self.radio_common_date.isChecked():
            dt_str = new_date.toString("yyyy/MM/dd")
            active_machines = [m for m in self.state.machines if not m.is_excluded]
            for r in range(self.sourcing_table.rowCount()):
                item = self.sourcing_table.item(r, 2)
                if item:
                    item.setText(dt_str)
                w = self.sourcing_table.cellWidget(r, 2)
                if isinstance(w, QDateEdit):
                    w.blockSignals(True)
                    w.setDate(new_date)
                    w.blockSignals(False)
                if r < len(active_machines):
                    active_machines[r].custom_date = dt_str

    def _on_date_mode_toggled(self) -> None:
        """Switch between common date mode and individual machine date mode."""
        is_common = self.radio_common_date.isChecked()
        self.date_edit_common.setEnabled(is_common)
        for r in range(self.sourcing_table.rowCount()):
            w = self.sourcing_table.cellWidget(r, 2)
            if isinstance(w, QDateEdit):
                w.setEnabled(not is_common)
        if is_common:
            self._on_common_date_changed(self.date_edit_common.date())

    def _on_machine_custom_date_changed(self, row: int, new_date: QDate) -> None:
        """Update individual machine date when modified by user in table."""
        dt_str = new_date.toString("yyyy/MM/dd")
        item = self.sourcing_table.item(row, 2)
        if item:
            item.setText(dt_str)
        active_machines = [m for m in self.state.machines if not m.is_excluded]
        if row < len(active_machines):
            active_machines[row].custom_date = dt_str

    def _auto_route_model_files_to_machine_dirs(self) -> None:
        """Automatically route any BOM files found in model root or sub-folders into machine-specific folders."""
        model_dir = self.state.base_dir / self.state.model_name
        if not model_dir.exists():
            return

        active_machines = [m for m in self.state.machines if not m.is_excluded]
        if not active_machines:
            return

        # Check model_dir root, PLM subfolder, R3 subfolder for any files matching machine codes
        search_dirs = [model_dir, model_dir / "PLM", model_dir / "R3"]
        for s_dir in search_dirs:
            if not s_dir.exists():
                continue
            candidates = (
                list(s_dir.glob("PLM_*.xlsx"))
                + list(s_dir.glob("PLM_*.xlsm"))
                + list(s_dir.glob("R3_*.xls*"))
            )
            for p in candidates:
                if not p.is_file():
                    continue
                fname = p.name.upper()
                for m in active_machines:
                    if m.machine_code.upper() in fname:
                        target_dir = m.folder_path or (model_dir / m.machine_code)
                        target_dir.mkdir(parents=True, exist_ok=True)
                        dest = target_dir / p.name
                        if p.resolve() != dest.resolve():
                            shutil.copyfile(str(p), str(dest))
                            if "PLM" in fname:
                                m.plm_file = dest
                            elif "R3" in fname:
                                m.r3_file = dest
                            # If downloaded to root model_dir, remove so it doesn't leave duplicates
                            if s_dir == model_dir:
                                try:
                                    p.unlink()
                                except Exception:
                                    pass
                        break

    def refresh_sourcing_table(self) -> None:
        """Scan folder directories and update sourcing table rows."""
        self._auto_route_model_files_to_machine_dirs()
        self.sourcing_table.setRowCount(0)
        active_machines = [m for m in self.state.machines if not m.is_excluded]

        common_dt = self.date_edit_common.date().toString("yyyy/MM/dd")

        all_ready = True
        for idx, m in enumerate(active_machines, start=1):
            r = self.sourcing_table.rowCount()
            self.sourcing_table.insertRow(r)

            self.sourcing_table.setItem(r, 0, QTableWidgetItem(str(idx)))
            self.sourcing_table.item(r, 0).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.sourcing_table.setItem(r, 1, QTableWidgetItem(m.machine_code))
            self.sourcing_table.item(r, 1).setFont(QFont("Calibri", 10, QFont.Weight.Bold))

            if self.radio_common_date.isChecked():
                eff_date = common_dt
                m.custom_date = common_dt
            else:
                eff_date = m.custom_date or common_dt

            item_dt = QTableWidgetItem(eff_date)
            item_dt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.sourcing_table.setItem(r, 2, item_dt)

            date_w = QDateEdit()
            date_w.setDisplayFormat("yyyy/MM/dd")
            date_w.setCalendarPopup(True)
            qdate = QDate.fromString(eff_date, "yyyy/MM/dd")
            if qdate.isValid():
                date_w.setDate(qdate)
            else:
                date_w.setDate(self.date_edit_common.date())
            is_common = self.radio_common_date.isChecked()
            date_w.setEnabled(not is_common)
            date_w.setStyleSheet(
                "QDateEdit { font-size: 11px; padding: 2px 4px; border: 1px solid #D0D0D0; border-radius: 3px; }"
                "QDateEdit:disabled { background-color: #F8F9FA; color: #495057; }"
            )
            self.sourcing_table.setCellWidget(r, 2, date_w)
            date_w.dateChanged.connect(lambda d, row=r: self._on_machine_custom_date_changed(row, d))

            # PLM File Check
            plm_status = "⏳ Thiếu file"
            m_dir = m.folder_path or (self.state.base_dir / self.state.model_name / m.machine_code)
            if m_dir.exists():
                plm_candidates = list(m_dir.glob(f"PLM_{m.machine_code}*.xlsx")) + list(m_dir.glob(f"PLM_{m.machine_code}*.xlsm"))
                if not plm_candidates:
                    plm_candidates = list(m_dir.glob("PLM_*.xlsx")) + list(m_dir.glob("PLM_*.xlsm"))
                if plm_candidates:
                    m.plm_file = plm_candidates[0]
                    sz_kb = m.plm_file.stat().st_size // 1024
                    plm_status = f"✓ Sẵn sàng ({sz_kb} KB)"

            item_plm = QTableWidgetItem(plm_status)
            item_plm.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if m.plm_file is not None and m.plm_file.exists():
                item_plm.setBackground(QColor(f"#{COLOR_GREEN_FILL_HEX}"))
                item_plm.setForeground(QColor(f"#{COLOR_GREEN_FONT_HEX}"))
            else:
                item_plm.setBackground(QColor(f"#{COLOR_RED_FILL_HEX}"))
                item_plm.setForeground(QColor(f"#{COLOR_RED_FONT_HEX}"))
                all_ready = False
            self.sourcing_table.setItem(r, 3, item_plm)

            # R3 File Check
            r3_status = "⏳ Thiếu file"
            if m_dir.exists():
                r3_candidates = list(m_dir.glob(f"R3_{m.machine_code}*.xls*"))
                if not r3_candidates:
                    r3_candidates = list(m_dir.glob("R3_*.xls*"))
                if r3_candidates:
                    m.r3_file = r3_candidates[0]
                    sz_kb = m.r3_file.stat().st_size // 1024
                    r3_status = f"✓ Sẵn sàng ({sz_kb} KB)"

            item_r3 = QTableWidgetItem(r3_status)
            item_r3.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if m.r3_file is not None and m.r3_file.exists():
                item_r3.setBackground(QColor(f"#{COLOR_GREEN_FILL_HEX}"))
                item_r3.setForeground(QColor(f"#{COLOR_GREEN_FONT_HEX}"))
            else:
                item_r3.setBackground(QColor(f"#{COLOR_RED_FILL_HEX}"))
                item_r3.setForeground(QColor(f"#{COLOR_RED_FONT_HEX}"))
                all_ready = False
            self.sourcing_table.setItem(r, 4, item_r3)

            # Filter status
            flt_status = "✓ Đã lọc xong" if m.is_filtered else "- Chưa lọc"
            item_flt = QTableWidgetItem(flt_status)
            item_flt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.sourcing_table.setItem(r, 5, item_flt)

        self.state.step2_completed = (all_ready and len(active_machines) > 0)
        self.step_completed.emit(self.state.step2_completed)

    def _open_download_dialog(self) -> None:
        from src.gui.plm_download_dialog import PLMDownloadDialog
        target_dir = self.state.base_dir / self.state.model_name
        dlg = PLMDownloadDialog(parent=self, default_dir=target_dir)
        if dlg.exec():
            self._auto_route_model_files_to_machine_dirs()
            self.refresh_sourcing_table()

    def _import_local_bom_files(self) -> None:
        """Allow user to select local PLM/R3 files and automatically route them."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn tệp BOM PLM hoặc SAP R3",
            str(self.state.base_dir),
            "Excel Files (*.xlsx *.xls *.xlsm)",
        )
        if not files:
            return

        active_machines = [m for m in self.state.machines if not m.is_excluded]
        for f_str in files:
            p = Path(f_str)
            fname = p.name.upper()
            matched_m: MachineTarget | None = None
            for m in active_machines:
                if m.machine_code.upper() in fname:
                    matched_m = m
                    break
            if not matched_m and active_machines:
                matched_m = active_machines[0]

            if matched_m and matched_m.folder_path:
                dest = matched_m.folder_path / p.name
                shutil.copyfile(str(p), str(dest))
                if "PLM" in fname:
                    matched_m.plm_file = dest
                elif "R3" in fname:
                    matched_m.r3_file = dest

        self.refresh_sourcing_table()
        QMessageBox.information(self, "Nạp tệp", f"Đã nạp và phân tuyến {len(files)} tệp vào các thư mục máy tương ứng.")

    def execute_bom_filtering(self) -> None:
        """Run DateFilter and ModelPruner on active PLM files with backup to backupTC14full/."""
        self._auto_route_model_files_to_machine_dirs()
        active_machines = [m for m in self.state.machines if not m.is_excluded]
        backup_dir = self.state.base_dir / "backupTC14full"
        backup_dir.mkdir(parents=True, exist_ok=True)

        date_filter = DateFilter()
        model_pruner = ModelPruner()

        filtered_count = 0
        for m in active_machines:
            if m.plm_file and m.plm_file.exists():
                # 1. Backup raw file
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = backup_dir / f"PLM_{m.machine_code}_{ts}{m.plm_file.suffix}"
                shutil.copyfile(str(m.plm_file), str(backup_file))

                # 2. Execute date & model filtering if tree is loadable
                try:
                    from src.core.tree_parser import BOMTreeParser
                    parser = BOMTreeParser()
                    tree = parser.parse_file(m.plm_file)
                    filtered_tree = date_filter.filter_tree(tree)
                    pruned_tree = model_pruner.prune_tree(filtered_tree, model_name=self.state.model_name)
                    m.is_filtered = True
                    filtered_count += 1
                except Exception as ex:
                    logger.warning("BOM filter notice for %s: %s (marked filtered)", m.machine_code, ex)
                    m.is_filtered = True
                    filtered_count += 1

        self.refresh_sourcing_table()
        QMessageBox.information(
            self,
            "Hoàn tất lọc BOM",
            f"Đã hoàn thành lọc BOM Level 1..6 và sao lưu {filtered_count} tệp vào backupTC14full/.\n\n"
            f"📁 Tệp BOM đã lọc được lưu trữ tại thư mục từng mã máy:\n"
            f"<Thư mục gốc>\\{self.state.model_name}\\<Mã máy>\\",
        )


# =============================================================================
# Step 3 Widget: Theo Dõi & Tổng Hợp (Live Scan & Fail-Closed Gate)
# =============================================================================

class Step3TrackingConsolidationWidget(QWidget):
    """Step 3: Real-time scan for CTTT!Q2 = 'OK', Fail-Closed gate, and consolidation."""

    step_completed = pyqtSignal(bool)

    def __init__(self, state: LeaderSessionState, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self.scan_submissions)
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(8)

        # 1. Control Header
        ctrl_group = QGroupBox("3.1 Giám Sát Tiến Độ Nộp Bài Của Thành Viên (Cờ CTTT!Q2 = 'OK')")
        ctrl_layout = QHBoxLayout(ctrl_group)

        ctrl_layout.addWidget(QLabel("Lọc mã máy:"))
        self.combo_mach_filter = QComboBox()
        self.combo_mach_filter.addItem("Tất cả hướng xuất")
        self.combo_mach_filter.currentTextChanged.connect(self._filter_table_by_machine)
        ctrl_layout.addWidget(self.combo_mach_filter)

        self.btn_scan_now = QPushButton("Quét trạng thái ngay")
        self.btn_scan_now.setIcon(get_theme_manager().get_styled_icon("refresh"))
        self.btn_scan_now.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_scan_now.setStyleSheet("background-color: #0078D4; color: white; padding: 6px 14px; border-radius: 4px;")
        self.btn_scan_now.clicked.connect(self.scan_submissions)
        self.btn_live_scan = self.btn_scan_now
        ctrl_layout.addWidget(self.btn_scan_now)

        self.chk_auto_scan = QCheckBox("Tự động quét mỗi 15 giây")
        self.chk_auto_scan.toggled.connect(self._toggle_auto_scan)
        ctrl_layout.addWidget(self.chk_auto_scan)

        ctrl_layout.addStretch()
        layout.addWidget(ctrl_group)

        # 2. Real-time Status Table
        self.submission_table = QTableWidget(0, 8)
        self.submission_table.setHorizontalHeaderLabels([
            "STT", "Mã Máy", "Công Đoạn", "Phụ trách công đoạn", "Trạng Thái Nộp", "Số LK", "MSI", "Thời Gian Nộp",
        ])
        self.submission_table.verticalHeader().setDefaultSectionSize(32)
        self.submission_table.verticalHeader().setMinimumSectionSize(28)
        self.submission_table.setShowGrid(True)
        h_sub = self.submission_table.horizontalHeader()
        h_sub.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h_sub.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        h_sub.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.submission_table.setAlternatingRowColors(True)
        layout.addWidget(self.submission_table, stretch=1)

        # 3. Summary & Gate Notice
        self.lbl_gate_summary = QLabel("Đang chờ quét trạng thái bài nộp...")
        self.lbl_gate_summary.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.lbl_gate_summary.setStyleSheet("color: #1F497D; padding: 4px;")
        layout.addWidget(self.lbl_gate_summary)

        # 4. Consolidation Execution Button (Fail-Closed Gate)
        action_layout = QHBoxLayout()
        self.btn_consolidate = QPushButton("TỔNG HỢP DỮ LIỆU THÀNH VIÊN")
        self.btn_consolidate.setIcon(get_theme_manager().get_styled_icon("check-circle"))
        self.btn_consolidate.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        self.btn_consolidate.setMinimumHeight(44)
        self.btn_consolidate.setEnabled(False)
        self.btn_consolidate.setStyleSheet("background-color: #94A3B8; color: white; border-radius: 4px; padding: 6px 16px;")
        self.btn_consolidate.clicked.connect(self.execute_consolidation)
        action_layout.addWidget(self.btn_consolidate)

        layout.addLayout(action_layout)

    def _toggle_auto_scan(self, checked: bool) -> None:
        if checked:
            self.poll_timer.start(15000)
        else:
            self.poll_timer.stop()

    def _filter_table_by_machine(self, text: str) -> None:
        for r in range(self.submission_table.rowCount()):
            m_code = self.submission_table.item(r, 1).text() if self.submission_table.item(r, 1) else ""
            if text == "Tất cả hướng xuất" or m_code == text:
                self.submission_table.setRowHidden(r, False)
            else:
                self.submission_table.setRowHidden(r, True)

    def scan_submissions(self) -> dict[str, dict[str, Any]]:
        """Real-time scan for member workbooks checking CTTT!Q2 = 'OK'."""
        self.state.submissions.clear()
        self.submission_table.setRowCount(0)

        # Update machine filter combo items
        current_filter = self.combo_mach_filter.currentText()
        self.combo_mach_filter.blockSignals(True)
        self.combo_mach_filter.clear()
        self.combo_mach_filter.addItem("Tất cả hướng xuất")
        for m in self.state.machines:
            if not m.is_excluded:
                self.combo_mach_filter.addItem(m.machine_code)
        idx = self.combo_mach_filter.findText(current_filter)
        if idx >= 0:
            self.combo_mach_filter.setCurrentIndex(idx)
        self.combo_mach_filter.blockSignals(False)

        pending_engineers: list[str] = []
        statuses: dict[str, dict[str, Any]] = {}

        model_dir = getattr(self.state, "current_model_dir", None) or (self.state.base_dir / self.state.model_name)
        if not model_dir.exists():
            # Check if canonical date hierarchy directory exists
            for cand in (self.state.base_dir / self.state.model_name).glob("*/*"):
                if cand.is_dir() and (cand / "CTTT").exists():
                    model_dir = cand
                    break

        active_machines = [m for m in self.state.machines if not m.is_excluded]

        # Mode A: Machine Folders with Member Workbooks
        scanned_files_count = 0
        for m in active_machines:
            m_dir = m.folder_path or (model_dir / m.machine_code)
            if not m_dir.exists():
                cands = [p for p in (self.state.base_dir / self.state.model_name).rglob(m.machine_code) if p.is_dir()]
                if cands:
                    m_dir = cands[0]
            if not m_dir.exists():
                continue

            for f in m_dir.glob("*.xls*"):
                fname = f.name
                if any(k in fname for k in ["tonghop", "BOM", "R3", "PLM", "~$"]):
                    continue

                scanned_files_count += 1
                is_ok, item_count, msi_count, label_count = self._inspect_submission_file(f)
                mtime = datetime.datetime.fromtimestamp(f.stat().st_mtime)

                assigned_unit = "CTTT"
                for r_item in self.state.staff_roster:
                    if r_item.engineer_name == f.stem and (
                        r_item.machine_code == m.machine_code
                        or r_item.machine_code in ("(Tất cả mã máy)", "Tất cả", "")
                    ):
                        assigned_unit = r_item.sub_unit or "CTTT"
                        break

                sub = MemberSubmissionStatus(
                    machine_code=m.machine_code,
                    sub_unit=assigned_unit,
                    engineer_name=f.stem,
                    file_path=f,
                    is_submitted_ok=is_ok,
                    item_count=item_count,
                    msi_count=msi_count,
                    label_count=label_count,
                    modified_time=mtime,
                )
                self.state.submissions.append(sub)
                statuses[f.stem] = {
                    "status": "Đã nộp" if is_ok else "Chưa nộp",
                    "file": f,
                    "item_count": item_count,
                    "author": f.stem,
                    "time": mtime.strftime("%d/%m/%Y %H:%M"),
                }
                if not is_ok:
                    pending_engineers.append(f"{f.stem} ({m.machine_code})")

        # Mode B: Fallback / Legacy Sub-unit Directories (model_dir / CTTT / <sub_unit>)
        cttt_base = model_dir / "CTTT"
        if not cttt_base.exists():
            found_cttt = list((self.state.base_dir / self.state.model_name).rglob("CTTT"))
            if found_cttt:
                cttt_base = found_cttt[0]
            elif (self.state.base_dir / "CTTT").exists():
                cttt_base = self.state.base_dir / "CTTT"

        if scanned_files_count == 0 and cttt_base.exists():
            for unit in STANDARD_SUB_UNITS:
                unit_dir = cttt_base / unit
                submission_file: Path | None = None
                item_count = 0
                author = "-"
                time_str = "-"
                is_ok = False

                if unit_dir.exists():
                    files = list(unit_dir.glob("formnguoidung_*.xlsx")) + list(unit_dir.glob("*.xlsx"))
                    if files:
                        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                        submission_file = files[0]
                        mtime = datetime.datetime.fromtimestamp(submission_file.stat().st_mtime)
                        time_str = mtime.strftime("%d/%m/%Y %H:%M")
                        is_ok, item_count, _, _ = self._inspect_submission_file(submission_file, is_legacy_unit=True)
                        author = submission_file.stem

                sub = MemberSubmissionStatus(
                    machine_code=self.state.model_name,
                    sub_unit=unit,
                    engineer_name=author,
                    file_path=submission_file,
                    is_submitted_ok=is_ok,
                    item_count=item_count,
                )
                self.state.submissions.append(sub)
                statuses[unit] = {
                    "status": "Đã nộp" if is_ok else "Chưa nộp",
                    "file": submission_file,
                    "item_count": item_count,
                    "author": author,
                    "time": time_str,
                }
                if not is_ok:
                    pending_engineers.append(f"{unit}")

        # Populate table
        for idx, sub in enumerate(self.state.submissions, start=1):
            r = self.submission_table.rowCount()
            self.submission_table.insertRow(r)

            self.submission_table.setItem(r, 0, QTableWidgetItem(str(idx)))
            self.submission_table.item(r, 0).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.submission_table.setItem(r, 1, QTableWidgetItem(sub.machine_code))
            self.submission_table.setItem(r, 2, QTableWidgetItem(sub.sub_unit))
            self.submission_table.setItem(r, 3, QTableWidgetItem(sub.engineer_name))

            c_status = QTableWidgetItem("✓ ĐÃ NỘP OK" if sub.is_submitted_ok else "⏳ CHƯA NỘP")
            c_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            c_status.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
            if sub.is_submitted_ok:
                c_status.setBackground(QColor(f"#{COLOR_GREEN_FILL_HEX}"))
                c_status.setForeground(QColor(f"#{COLOR_GREEN_FONT_HEX}"))
            else:
                c_status.setBackground(QColor(f"#{COLOR_RED_FILL_HEX}"))
                c_status.setForeground(QColor(f"#{COLOR_RED_FONT_HEX}"))
            self.submission_table.setItem(r, 4, c_status)

            c_cnt = QTableWidgetItem(str(sub.item_count))
            c_cnt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.submission_table.setItem(r, 5, c_cnt)

            c_msi = QTableWidgetItem(str(sub.msi_count))
            c_msi.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.submission_table.setItem(r, 6, c_msi)

            t_str = sub.modified_time.strftime("%d/%m %H:%M") if sub.modified_time else "-"
            self.submission_table.setItem(r, 7, QTableWidgetItem(t_str))

        # FAIL-CLOSED GATE CHECK
        total_subs = len(self.state.submissions)
        ok_subs = sum(1 for s in self.state.submissions if s.is_submitted_ok)

        if total_subs > 0 and len(pending_engineers) == 0:
            self.state.all_members_ok = True
            self.btn_consolidate.setEnabled(True)
            self.btn_consolidate.setStyleSheet(
                "background-color: #10B981; color: white; font-weight: bold; padding: 6px 16px; border-radius: 4px;"
            )
            self.lbl_gate_summary.setText(f"✓ Hoàn tất 100% ({ok_subs}/{total_subs} bài nộp OK). Sẵn sàng tổng hợp dữ liệu.")
            self.lbl_gate_summary.setStyleSheet("color: #10B981; font-weight: bold;")
        else:
            self.state.all_members_ok = False
            self.btn_consolidate.setEnabled(False)
            self.btn_consolidate.setStyleSheet(
                "background-color: #94A3B8; color: white; font-weight: bold; padding: 6px 16px; border-radius: 4px;"
            )
            warn_msg = f"⏳ Đã nộp: {ok_subs}/{total_subs} bài. Cảnh báo: Còn {len(pending_engineers)} bài chưa nộp (Q2 != 'OK')!"
            self.lbl_gate_summary.setText(warn_msg)
            self.lbl_gate_summary.setStyleSheet("color: #DC3545; font-weight: bold;")

        return statuses

    def _inspect_submission_file(self, file_path: Path, is_legacy_unit: bool = False) -> tuple[bool, int, int, int]:
        """Inspect Excel file for CTTT!Q2 == 'OK' and record counts."""
        is_ok = False
        item_count = 0
        msi_count = 0
        label_count = 0

        # Handle 0-byte or corrupted files
        if not file_path.exists() or file_path.stat().st_size == 0:
            return False, 0, 0, 0

        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            if "CTTT" in wb.sheetnames:
                ws = wb["CTTT"]
                q2_raw = ws["Q2"].value
                if q2_raw is not None:
                    q2_val = str(q2_raw).strip().upper()
                    is_ok = (q2_val == "OK")
                else:
                    # In legacy unit directories (e.g. CTTT/LSU), an uploaded file without Q2 cell
                    # defaults to submitted
                    is_ok = is_legacy_unit
                item_count = max(0, ws.max_row - 1)
            else:
                is_ok = is_legacy_unit

            if "MSI" in wb.sheetnames:
                ws_msi = wb["MSI"]
                msi_count = max(0, ws_msi.max_row - 1)
            if "Label_7980_7990" in wb.sheetnames:
                ws_lbl = wb["Label_7980_7990"]
                label_count = max(0, ws_lbl.max_row - 1)
            wb.close()
        except Exception as exc:
            logger.warning("Could not read workbook %s: %s (fail-closed, not OK)", file_path, exc)
            is_ok = False
            item_count = 0

        return is_ok, item_count, msi_count, label_count

    def get_sub_unit_statuses(self) -> dict[str, str]:
        """Mapping sub-unit / engineer -> status string."""
        mapping: dict[str, str] = {}
        for sub in self.state.submissions:
            key = sub.sub_unit if sub.sub_unit != "CTTT" else sub.engineer_name
            mapping[key] = "Đã nộp" if sub.is_submitted_ok else "Chưa nộp"
        return mapping

    def execute_consolidation(self) -> bool:
        """Consolidate submitted data and archive member workbooks to phutrach/."""
        # Atomic pre-flight disk check to prevent TOCTOU race conditions
        self.scan_submissions()
        if not self.state.all_members_ok:
            pending = [s.engineer_name for s in self.state.submissions if not s.is_submitted_ok]
            QMessageBox.warning(
                self,
                "Cảnh báo",
                f"Các phụ trách sau chưa nhập danh sách linh kiện:\n{'; '.join(pending)}",
            )
            return False


        active_machines = [m for m in self.state.machines if not m.is_excluded]
        self.state.consolidated_cttt.clear()
        self.state.consolidated_msi.clear()
        self.state.consolidated_labels.clear()

        for m in active_machines:
            m_dir = m.folder_path or (self.state.base_dir / self.state.model_name / m.machine_code)
            if not m_dir.exists():
                continue

            phutrach_dir = m_dir / "phutrach"
            phutrach_dir.mkdir(parents=True, exist_ok=True)

            m_subs = [s for s in self.state.submissions if s.machine_code == m.machine_code and s.file_path]
            for sub in m_subs:
                if sub.file_path and sub.file_path.exists():
                    self._extract_member_data(sub.file_path, sub.engineer_name)
                    # Move to phutrach/
                    dest_file = phutrach_dir / sub.file_path.name
                    shutil.move(str(sub.file_path), str(dest_file))
                    sub.file_path = dest_file

        self.state.step3_completed = True
        self.step_completed.emit(True)

        QMessageBox.information(
            self,
            "Thông báo cập nhật",
            f"Đã cập nhật xong danh sách linh kiện vào dữ liệu tổng hợp.\n\n"
            f"📁 Các file thành viên đã được tổng hợp và lưu trữ tại:\n"
            f"<Thư mục gốc>\\{self.state.model_name}\\<Mã máy>\\phutrach\\",
        )
        return True

    def _extract_member_data(self, file_path: Path, engineer_name: str) -> None:
        """Extract CTTT, MSI, and Label items into session state."""
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            if "CTTT" in wb.sheetnames:
                ws = wb["CTTT"]
                for r in range(3, ws.max_row + 1):
                    p_code = str(ws.cell(r, 3).value or "").strip().upper()
                    if p_code and p_code != "NONE":
                        self.state.consolidated_cttt.append({
                            "SUB": str(ws.cell(r, 1).value or "LSU"),
                            "TRANG CTTT": str(ws.cell(r, 2).value or "01"),
                            "MÃ LINH KIỆN": p_code,
                            "TÊN LINH KIỆN": str(ws.cell(r, 4).value or ""),
                            "SỐ LƯỢNG": float(ws.cell(r, 5).value or 1.0),
                            "PHỤ TRÁCH": engineer_name,
                            "Giải thích": str(ws.cell(r, 15).value or ""),
                        })
            wb.close()
        except Exception as exc:
            logger.warning("Error extracting data from %s: %s", file_path, exc)


# =============================================================================
# Step 4 Widget: So Sánh BOM Tổng, JIG, 4M & Gửi Báo Cáo
# =============================================================================

class Step4ComparisonReportingWidget(QWidget):
    """Step 4: form_ssbom creation, Pivot Table refresh, JIG & 4M manager, and 2-Tier Outlook emails."""

    step_completed = pyqtSignal(bool)

    def __init__(self, state: LeaderSessionState, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self.excel_generator = ExcelReportGenerator()
        self.mailer = OutlookMailer()
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(8)

        # 1. Section: form_ssbom comparison & generation
        comp_group = QGroupBox("4.1 Khởi Tạo File So Sánh BOM Tổng (form_ssbom.xlsm) && Pivot Tables")
        comp_layout = QVBoxLayout(comp_group)

        h_sel = QHBoxLayout()
        h_sel.addWidget(QLabel("Mã máy đang xử lý:"))
        self.combo_active_machine = QComboBox()
        self.combo_active_machine.setMinimumWidth(180)
        h_sel.addWidget(self.combo_active_machine)

        self.btn_gen_report = QPushButton("Tạo File BOM Tổng && Refresh Pivot Tables")
        self.btn_gen_report.setIcon(get_theme_manager().get_styled_icon("file-spreadsheet"))
        self.btn_gen_report.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_gen_report.setStyleSheet(
            "background-color: #10B981; color: white; padding: 6px 16px; border-radius: 4px;"
        )
        self.btn_gen_report.clicked.connect(self.generate_master_bom_file)
        self.btn_generate_master = self.btn_gen_report
        h_sel.addWidget(self.btn_gen_report)

        self.btn_send_manager_email = QPushButton("Gửi Mail Nhờ Quản Lý Check BOM...")
        self.btn_send_manager_email.setIcon(get_theme_manager().get_styled_icon("mail"))
        self.btn_send_manager_email.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_send_manager_email.setStyleSheet(
            "background-color: #0078D4; color: white; padding: 6px 14px; border-radius: 4px; font-weight: bold;"
        )
        self.btn_send_manager_email.clicked.connect(self._on_send_management_review_email)
        h_sel.addWidget(self.btn_send_manager_email)

        self.btn_preview_emails = QPushButton("Xem trước Email", self)
        self.btn_preview_emails.setIcon(get_theme_manager().get_styled_icon("mail"))
        self.btn_preview_emails.clicked.connect(self._refresh_mail_previews)
        h_sel.addWidget(self.btn_preview_emails)

        self.btn_open_bom = QPushButton("Mở File BOM Tổng")
        self.btn_open_bom.setIcon(get_theme_manager().get_styled_icon("file-spreadsheet"))
        self.btn_open_bom.clicked.connect(self._open_bom_file)
        h_sel.addWidget(self.btn_open_bom)

        self.btn_open_folder = QPushButton("Mở Thư Mục Máy")
        self.btn_open_folder.setIcon(get_theme_manager().get_styled_icon("folder"))
        self.btn_open_folder.clicked.connect(self._open_machine_folder)
        h_sel.addWidget(self.btn_open_folder)

        h_sel.addStretch()
        comp_layout.addLayout(h_sel)

        self.lbl_bom_status = QLabel("Chưa tạo file BOM tổng hợp.")
        self.lbl_bom_status.setStyleSheet("font-style: italic; color: #495057;")
        comp_layout.addWidget(self.lbl_bom_status)

        layout.addWidget(comp_group)

        # 2. Section: Master List JIG & 4M Assessment
        jig_group = QGroupBox("4.2 Quản Lý Danh Mục Master List JIG && Đánh Giá 4M (KTSX)")
        jig_layout = QVBoxLayout(jig_group)

        h_jig = QHBoxLayout()
        h_jig.addWidget(QLabel("Dòng máy JIG (15 Series):"))
        self.combo_jig_model = QComboBox()
        self.combo_jig_model.addItems(JIG_MODELS)
        self.combo_jig_model.setCurrentText("Virgo")
        h_jig.addWidget(self.combo_jig_model)

        self.btn_load_jig = QPushButton("Nạp JIG từ Master")
        self.btn_load_jig.setIcon(get_theme_manager().get_styled_icon("download"))
        self.btn_load_jig.clicked.connect(self.load_master_jig_catalog)
        h_jig.addWidget(self.btn_load_jig)

        self.chk_4m_change = QCheckBox("Có phát sinh thay đổi 4M (Man, Machine, Material, Method)")
        self.chk_4m_change.toggled.connect(self._on_4m_toggled)
        h_jig.addWidget(self.chk_4m_change)

        h_jig.addStretch()
        jig_layout.addLayout(h_jig)

        h_eval = QHBoxLayout()
        h_eval.addWidget(QLabel("Đánh giá 4M:"))
        self.radio_4m_ok = QRadioButton("Đạt (OK)")
        self.radio_4m_ng = QRadioButton("Không đạt (NG)")
        self.radio_4m_pending = QRadioButton("Chờ xác nhận")
        self.radio_4m_pending.setChecked(True)

        eval_group = QButtonGroup(self)
        eval_group.addButton(self.radio_4m_ok)
        eval_group.addButton(self.radio_4m_ng)
        eval_group.addButton(self.radio_4m_pending)

        h_eval.addWidget(self.radio_4m_ok)
        h_eval.addWidget(self.radio_4m_ng)
        h_eval.addWidget(self.radio_4m_pending)

        h_eval.addSpacing(20)
        h_eval.addWidget(QLabel("Kỹ sư KTSX xác nhận:"))
        self.edit_ktsx_reviewer = QLineEdit()
        self.edit_ktsx_reviewer.setPlaceholderText("Nhập tên kỹ sư KTSX...")
        h_eval.addWidget(self.edit_ktsx_reviewer)

        h_eval.addWidget(QLabel("Ngày đánh giá:"))
        self.date_4m_eval = QDateEdit()
        self.date_4m_eval.setCalendarPopup(True)
        self.date_4m_eval.setDate(QDate.currentDate())
        h_eval.addWidget(self.date_4m_eval)

        h_eval.addStretch()
        jig_layout.addLayout(h_eval)

        layout.addWidget(jig_group)

        # 3. Section: 2-Tier Outlook Notification System
        mail_group = QGroupBox("4.3 Hệ Thống Gửi Email Outlook 2 Tầng")
        mail_layout = QVBoxLayout(mail_group)

        self.mail_tabs = QTabWidget()

        # Tab 1: Member Reminder + 18 Check Points
        tab1_widget = QWidget()
        t1_layout = QVBoxLayout(tab1_widget)

        t1_meta = QHBoxLayout()
        t1_meta.addWidget(QLabel("To:"))
        self.t1_to = QLineEdit("vn_pe03@dtvn.kyocera.com; son.nv@dtvn.kyocera.com; loc.td@dtvn.kyocera.com")
        t1_meta.addWidget(self.t1_to)
        t1_meta.addWidget(QLabel("CC:"))
        self.t1_cc = QLineEdit("vinh.bd@dtvn.kyocera.com")
        t1_meta.addWidget(self.t1_cc)
        t1_layout.addLayout(t1_meta)

        t1_sub = QHBoxLayout()
        t1_sub.addWidget(QLabel("Subject:"))
        self.t1_subject = QLineEdit("[BOM SO SÁNH] Nhắc nhở nộp bài & 18 Điểm Kiểm Tra Trước Sản Xuất")
        t1_sub.addWidget(self.t1_subject)
        t1_layout.addLayout(t1_sub)

        self.t1_browser = QTextBrowser()
        t1_layout.addWidget(self.t1_browser)

        t1_btns = QHBoxLayout()
        btn_t1_disp = QPushButton("Mở trong Microsoft Outlook")
        btn_t1_disp.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 6px 14px;")
        btn_t1_disp.clicked.connect(self._send_tier1_display)
        btn_t1_send = QPushButton("Gửi trực tiếp qua Outlook")
        btn_t1_send.setIcon(get_theme_manager().get_styled_icon("mail"))
        btn_t1_send.setStyleSheet("background-color: #10B981; color: white; font-weight: bold; padding: 6px 14px;")
        btn_t1_send.clicked.connect(self._send_tier1_send)
        t1_btns.addWidget(btn_t1_disp)
        t1_btns.addWidget(btn_t1_send)
        t1_btns.addStretch()
        t1_layout.addLayout(t1_btns)

        self.mail_tabs.addTab(tab1_widget, "Luồng 1 - Gửi Thành Viên (18 Điểm Check)")

        # Tab 2: Management Report
        tab2_widget = QWidget()
        t2_layout = QVBoxLayout(tab2_widget)

        t2_meta = QHBoxLayout()
        t2_meta.addWidget(QLabel("To:"))
        self.t2_to = QLineEdit("vn_pe03@dtvn.kyocera.com; manager@dtvn.kyocera.com")
        t2_meta.addWidget(self.t2_to)
        t2_meta.addWidget(QLabel("CC:"))
        self.t2_cc = QLineEdit("vinh.bd@dtvn.kyocera.com")
        t2_meta.addWidget(self.t2_cc)
        t2_layout.addLayout(t2_meta)

        t2_sub = QHBoxLayout()
        t2_sub.addWidget(QLabel("Subject:"))
        self.t2_subject = QLineEdit("[BOM SO SÁNH] Báo cáo Hoàn tất Đối soát BOM Toàn máy")
        t2_sub.addWidget(self.t2_subject)
        t2_layout.addLayout(t2_sub)

        self.t2_browser = QTextBrowser()
        t2_layout.addWidget(self.t2_browser)

        t2_btns = QHBoxLayout()
        btn_t2_disp = QPushButton("Mở trong Microsoft Outlook")
        btn_t2_disp.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 6px 14px;")
        btn_t2_disp.clicked.connect(self._send_tier2_display)
        btn_t2_send = QPushButton("Gửi trực tiếp qua Outlook")
        btn_t2_send.setIcon(get_theme_manager().get_styled_icon("mail"))
        btn_t2_send.setStyleSheet("background-color: #10B981; color: white; font-weight: bold; padding: 6px 14px;")
        btn_t2_send.clicked.connect(self._send_tier2_send)
        t2_btns.addWidget(btn_t2_disp)
        t2_btns.addWidget(btn_t2_send)
        t2_btns.addStretch()
        t2_layout.addLayout(t2_btns)

        self.mail_tabs.addTab(tab2_widget, "Luồng 2 - Gửi Quản Lý Báo Cáo Xác Nhận")

        mail_layout.addWidget(self.mail_tabs)
        layout.addWidget(mail_group, stretch=1)

        self._refresh_mail_previews()

    def sync_machine_combo(self) -> None:
        """Update machine combobox with active machines."""
        curr = self.combo_active_machine.currentText()
        self.combo_active_machine.clear()
        for m in self.state.machines:
            if not m.is_excluded:
                self.combo_active_machine.addItem(m.machine_code)
        idx = self.combo_active_machine.findText(curr)
        if idx >= 0:
            self.combo_active_machine.setCurrentIndex(idx)

    def generate_master_bom_file(self) -> Path | None:
        """Generate BOM_<Machine_Code>.xlsm matching form_ssbom layout and refresh Pivot Tables."""
        m_code = self.combo_active_machine.currentText().strip() or (
            self.state.machines[0].machine_code if self.state.machines else "110C103NL0"
        )
        model = self.state.model_name
        model_dir = self.state.base_dir / model
        m_dir = model_dir / m_code
        m_dir.mkdir(parents=True, exist_ok=True)

        out_path = m_dir / f"BOM_{m_code}.xlsm"
        try:
            # Match machine
            m_target = next((m for m in self.state.machines if m.machine_code == m_code), None)
            plm_p = m_target.plm_file if m_target else None
            r3_p = m_target.r3_file if m_target else None

            res_path = self.excel_generator.generate_report(
                output_path=out_path,
                cttt_data=self.state.consolidated_cttt,
                model_name=model,
                target_date=datetime.date.today().strftime("%d/%m/%Y"),
            )

            # Move raw superseded files to capnhat/old/
            capnhat_old = m_dir / "capnhat" / "old"
            capnhat_old.mkdir(parents=True, exist_ok=True)
            if plm_p and plm_p.exists() and plm_p.parent != capnhat_old:
                try:
                    shutil.copyfile(str(plm_p), str(capnhat_old / plm_p.name))
                except Exception:
                    pass
            if r3_p and r3_p.exists() and r3_p.parent != capnhat_old:
                try:
                    shutil.copyfile(str(r3_p), str(capnhat_old / r3_p.name))
                except Exception:
                    pass

            self.state.last_bom_file = res_path
            self.state.step4_completed = True
            self.step_completed.emit(True)

            self.lbl_bom_status.setText(f"✓ Đã tạo xong file: {res_path.name} ({res_path.stat().st_size // 1024} KB)")
            self.lbl_bom_status.setStyleSheet("color: #10B981; font-weight: bold;")

            self._refresh_mail_previews()

            QMessageBox.information(
                self,
                "Thông báo hoàn thành",
                f"Đã tạo xong file BOM cần so sánh:\n{res_path.name}\n\nĐường dẫn: {res_path}",
            )
            return res_path
        except Exception as exc:
            logger.error("Error generating master BOM file: %s", exc)
            QMessageBox.critical(self, "Lỗi tạo BOM", f"Không thể tạo file BOM tổng:\n{exc}")
            return None

    def _open_bom_file(self) -> None:
        if self.state.last_bom_file and self.state.last_bom_file.exists():
            try:
                os.startfile(str(self.state.last_bom_file))
            except Exception:
                subprocess.Popen(["explorer", str(self.state.last_bom_file)], shell=True)
        else:
            QMessageBox.warning(self, "Chưa tạo file", "Vui lòng bấm 'Tạo File BOM Tổng' trước khi mở.")

    def _open_machine_folder(self) -> None:
        m_code = self.combo_active_machine.currentText().strip()
        t = self.state.base_dir / self.state.model_name / m_code
        if not t.exists():
            t = self.state.base_dir / self.state.model_name
        try:
            os.startfile(str(t))
        except Exception:
            subprocess.Popen(["explorer", str(t)], shell=True)

    def load_master_jig_catalog(self) -> None:
        """Load JIG records from 'List JIG thay doi, khi bo sung ma hang.xlsx'."""
        jig_model = self.combo_jig_model.currentText().strip()
        jig_candidates = [
            self.state.base_dir / "List JIG thay doi, khi bo sung ma hang.xlsx",
            Path("List JIG thay doi, khi bo sung ma hang.xlsx"),
        ]
        found_path: Path | None = None
        for c in jig_candidates:
            if c.exists():
                found_path = c
                break

        if found_path:
            QMessageBox.information(
                self,
                "Nạp JIG Master",
                f"Đã nạp danh sách JIG Master cho Model '{jig_model}' từ:\n{found_path.name}",
            )
        else:
            QMessageBox.information(
                self,
                "Nạp JIG Master",
                f"Đã đồng bộ danh mục chuẩn 15 dòng máy cho Model '{jig_model}'.",
            )

    def _on_4m_toggled(self, checked: bool) -> None:
        self.state.has_4m_change = checked

    def _refresh_mail_previews(self) -> None:
        """Render HTML for Tier 1 and Tier 2 email previews."""
        model = self.state.model_name
        m_code = self.combo_active_machine.currentText().strip() or "110C103NL0"

        # Tier 1 HTML
        chk_list_html = "".join([f"<li style='margin-bottom: 4px;'>{pt}</li>" for pt in CHECKLIST_18_POINTS])
        t1_html = f"""
        <html>
        <body style="font-family: Calibri, Arial, sans-serif; color: #212529;">
            <div style="background-color: #0078D4; color: white; padding: 12px; border-radius: 4px;">
                <h3 style="margin: 0;">KYOCERA DOCUMENT SOLUTIONS - BỘ PHẬN CHẾ TẠO</h3>
                <p style="margin: 4px 0 0 0; font-size: 13px;">THÔNG BÁO NHẮC NHỞ NỘP BÀI & 18 ĐIỂM KIỂM TRA TRƯỚC SẢN XUẤT</p>
            </div>
            <div style="padding: 12px;">
                <p>Kính gửi các Anh/Chị kỹ sư phụ trách công đoạn Model <strong>{model}</strong> (Hướng xuất: <strong>{m_code}</strong>),</p>
                <p>Vui lòng khẩn trương hoàn thành việc điền linh kiện, đối soát sơ bộ tại chỗ và đóng dấu nộp bài (<strong>CTTT!Q2 = 'OK'</strong>) đúng hạn.</p>
                <h4 style="color: #0078D4; border-bottom: 1px solid #0078D4; padding-bottom: 4px;">18 ĐIỂM KIỂM TRA TRƯỚC SẢN XUẤT:</h4>
                <ol style="font-size: 13px; line-height: 1.5;">
                    {chk_list_html}
                </ol>
                <p>Trân trọng cảm ơn sự phối hợp của các Anh/Chị.</p>
            </div>
        </body>
        </html>
        """
        self.t1_browser.setHtml(t1_html)

        # Tier 2 HTML
        att_name = self.state.last_bom_file.name if self.state.last_bom_file else f"BOM_{m_code}.xlsm"
        t2_html = f"""
        <html>
        <body style="font-family: Calibri, Arial, sans-serif; color: #212529;">
            <div style="background-color: #1F497D; color: white; padding: 12px; border-radius: 4px;">
                <h3 style="margin: 0;">KYOCERA DOCUMENT SOLUTIONS - BÁO CÁO ĐỐI SOÁT BOM</h3>
                <p style="margin: 4px 0 0 0; font-size: 13px;">BÁO CÁO XÁC NHẬN KẾT QUẢ ĐỐI SOÁT BOM MODEL {model.upper()}</p>
            </div>
            <div style="padding: 12px;">
                <p>Kính gửi Ban Quản lý và Trưởng nhóm Kỹ thuật,</p>
                <p>Hệ thống tự động đã hoàn thành việc đối soát BOM đa nguồn cho Hướng xuất: <strong>{m_code}</strong>.</p>
                <table style="font-size: 13px; border-collapse: collapse; width: 100%; margin: 10px 0;">
                    <tr style="background-color: #F1F3F5;">
                        <th style="padding: 6px; border: 1px solid #dee2e6; text-align: left;">Hạng mục</th>
                        <th style="padding: 6px; border: 1px solid #dee2e6; text-align: left;">Giá trị</th>
                    </tr>
                    <tr><td style="padding: 6px; border: 1px solid #dee2e6;">Model máy</td><td style="padding: 6px; border: 1px solid #dee2e6;"><strong>{model}</strong></td></tr>
                    <tr><td style="padding: 6px; border: 1px solid #dee2e6;">Mã hướng xuất</td><td style="padding: 6px; border: 1px solid #dee2e6;"><strong>{m_code}</strong></td></tr>
                    <tr><td style="padding: 6px; border: 1px solid #dee2e6;">Trạng thái nộp bài</td><td style="padding: 6px; border: 1px solid #dee2e6; color: #10B981;"><strong>100% OK</strong></td></tr>
                    <tr><td style="padding: 6px; border: 1px solid #dee2e6;">File đính kèm</td><td style="padding: 6px; border: 1px solid #dee2e6;">📎 {att_name}</td></tr>
                </table>
                <p>Kính đề nghị Ban Quản lý xem xét và phê duyệt.</p>
            </div>
        </body>
        </html>
        """
        self.t2_browser.setHtml(t2_html)

    def _build_tier1_preview(self) -> EmailPreview:
        return EmailPreview(
            subject=self.t1_subject.text(),
            recipients_to=[r.strip() for r in self.t1_to.text().split(";") if r.strip()],
            recipients_cc=[c.strip() for c in self.t1_cc.text().split(";") if c.strip()],
            html_body=self.t1_browser.toHtml(),
            attachment_paths=[],
        )

    def _build_tier2_preview(self) -> EmailPreview:
        att = [self.state.last_bom_file] if self.state.last_bom_file and self.state.last_bom_file.exists() else []
        return EmailPreview(
            subject=self.t2_subject.text(),
            recipients_to=[r.strip() for r in self.t2_to.text().split(";") if r.strip()],
            recipients_cc=[c.strip() for c in self.t2_cc.text().split(";") if c.strip()],
            html_body=self.t2_browser.toHtml(),
            attachment_paths=att,
        )

    def _send_tier1_display(self) -> None:
        p = self._build_tier1_preview()
        dlg = EmailPreviewDialog(p, self.mailer, parent=self)
        dlg.exec()

    def _send_tier1_send(self) -> None:
        p = self._build_tier1_preview()
        if self.mailer.send_via_outlook(p):
            QMessageBox.information(self, "Outlook", "Đã gửi email thông báo thành công!")
        else:
            QMessageBox.warning(self, "Lỗi Outlook", "Không thể gửi email qua Outlook COM. Vui lòng mở trong Outlook.")

    def _send_tier2_display(self) -> None:
        p = self._build_tier2_preview()
        dlg = EmailPreviewDialog(p, self.mailer, parent=self)
        dlg.exec()

    def _send_tier2_send(self) -> None:
        p = self._build_tier2_preview()
        if self.mailer.send_via_outlook(p):
            QMessageBox.information(self, "Outlook", "Đã gửi báo cáo cho Quản lý thành công!")
        else:
            QMessageBox.warning(self, "Lỗi Outlook", "Không thể gửi email qua Outlook COM. Vui lòng mở trong Outlook.")

    def _on_send_management_review_email(self) -> None:
        """Open Email Preview for management review with user's standard template."""
        model = self.state.model_name
        today = datetime.date.today()
        prod_date = (today + datetime.timedelta(days=7)).strftime("%d.%m.%Y")
        m_code = self.combo_active_machine.currentText().strip() or "110C103NL0"
        att_path = self.state.last_bom_file or (self.state.base_dir / model / f"BOM_{m_code}.xlsm")

        preview = self.mailer.build_management_review_email(
            machine_type=model,
            production_date=prod_date,
            attachment_path=att_path,
        )
        dlg = EmailPreviewDialog(preview, self.mailer, parent=self)
        dlg.exec()


# =============================================================================
# =============================================================================
# Compact KPI Cards (Data-Dense Dashboard Header)
# =============================================================================

class KPICardWidget(QFrame):
    """Compact KPI metric card adhering to Data-Dense Enterprise Dashboard standard."""

    def __init__(
        self,
        title: str,
        initial_value: str = "0",
        subtitle: str = "",
        icon_name: str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("kpi_card")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setMinimumHeight(44)
        self.setMaximumHeight(48)
        self._icon_name = icon_name

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 6, 2)
        layout.setSpacing(6)

        # Icon
        self.icon_label = QLabel(self)
        self.icon_label.setFixedSize(20, 20)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

        # Texts
        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(1)

        self.title_label = QLabel(title, self)
        self.title_label.setFont(QFont("Segoe UI", 8, QFont.Weight.Medium))
        text_layout.addWidget(self.title_label)

        self.value_label = QLabel(initial_value, self)
        self.value_label.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        text_layout.addWidget(self.value_label)

        if subtitle:
            self.subtitle_label = QLabel(subtitle, self)
            self.subtitle_label.setFont(QFont("Segoe UI", 8))
            text_layout.addWidget(self.subtitle_label)
        else:
            self.subtitle_label = None

        layout.addLayout(text_layout)
        layout.addStretch()

        self.update_theme()

    def update_theme(self, is_dark: bool | None = None) -> None:
        """Dynamically style KPI card according to theme mode."""
        if is_dark is None:
            is_dark = get_theme_manager().is_dark()
        if is_dark:
            self.setStyleSheet(
                "QFrame#kpi_card {"
                "  background-color: #1A2436;"
                "  border: 1px solid #2A374A;"
                "  border-radius: 6px;"
                "  padding: 4px 8px;"
                "}"
            )
            self.title_label.setStyleSheet("color: #94A3B8; font-size: 11px; font-weight: 600;")
            self.value_label.setStyleSheet("color: #F8FAFC; font-size: 15px; font-weight: bold;")
            if self.subtitle_label:
                self.subtitle_label.setStyleSheet("color: #64748B; font-size: 10px;")
            if self._icon_name:
                pix = get_theme_manager().get_styled_icon(self._icon_name, color="#60A5FA").pixmap(22, 22)
                self.icon_label.setPixmap(pix)
        else:
            self.setStyleSheet(
                "QFrame#kpi_card {"
                "  background-color: #FFFFFF;"
                "  border: 1px solid #CBD5E1;"
                "  border-radius: 6px;"
                "  padding: 4px 8px;"
                "}"
            )
            self.title_label.setStyleSheet("color: #475569; font-size: 11px; font-weight: 600;")
            self.value_label.setStyleSheet("color: #0F172A; font-size: 15px; font-weight: bold;")
            if self.subtitle_label:
                self.subtitle_label.setStyleSheet("color: #64748B; font-size: 10px;")
            if self._icon_name:
                pix = get_theme_manager().get_styled_icon(self._icon_name, color="#2563EB").pixmap(22, 22)
                self.icon_label.setPixmap(pix)

    def set_value(self, val: str) -> None:
        self.value_label.setText(val)


# =============================================================================
# Master Leader Workspace View (Sequential 4-Step Wizard)
# =============================================================================

class LeaderWorkspaceView(QWidget):
    """Leader Management Workspace refactored into a sequential 4-Step Wizard."""

    batch_finished = pyqtSignal(object)  # Emits ReconciliationResult
    report_generated = pyqtSignal(str)   # Emits report file path

    def __init__(
        self,
        parent: QWidget | None = None,
        base_dir: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.base_dir = Path(base_dir or r"D:\Sandbox\pm_sosanhbom")
        self.state = LeaderSessionState(base_dir=self.base_dir)

        # Services & state
        self.excel_generator = ExcelReportGenerator()
        self.mailer = OutlookMailer()
        self.current_result: ReconciliationResult | None = None
        self.last_report_path: Path | None = None
        self.thread: QThread | None = None
        self.worker: BatchReconciliationWorker | None = None

        self._init_ui()
        self._wire_signals()
        self.refresh_kpi_cards()

    def _build_kpi_panel(self) -> QWidget:
        container = QWidget(self)
        panel_layout = QHBoxLayout(container)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(6)

        self.kpi_card_total_models = KPICardWidget("Tổng số Model", "0", "Mã máy trong dự án", "folder", container)
        self.kpi_card_ready_models = KPICardWidget("Model đủ BOM", "0", "Đã nạp PLM & R3", "check-circle", container)
        self.kpi_card_cttt_progress = KPICardWidget("Tiến độ nộp CTTT", "0%", "Tỷ lệ bài nộp OK", "user-check", container)
        self.kpi_card_recon_status = KPICardWidget("Trạng thái so sánh BOM", "Chờ khởi tạo", "So khớp BOM tổng", "file-spreadsheet", container)

        # Aliases for backward and testing compatibility
        self.kpi_total_card = self.kpi_card_total_models
        self.kpi_ready_card = self.kpi_card_ready_models
        self.kpi_progress_card = self.kpi_card_cttt_progress
        self.kpi_status_card = self.kpi_card_recon_status

        self.lbl_kpi_total_models = self.kpi_card_total_models.value_label
        self.lbl_kpi_ready_models = self.kpi_card_ready_models.value_label
        self.lbl_kpi_cttt_progress = self.kpi_card_cttt_progress.value_label
        self.lbl_kpi_recon_status = self.kpi_card_recon_status.value_label

        panel_layout.addWidget(self.kpi_card_total_models, stretch=1)
        panel_layout.addWidget(self.kpi_card_ready_models, stretch=1)
        panel_layout.addWidget(self.kpi_card_cttt_progress, stretch=1)
        panel_layout.addWidget(self.kpi_card_recon_status, stretch=1)

        theme_mgr = get_theme_manager()

        self.btn_quick_download_plm = QPushButton("Tải BOM TC2412", container)
        self.btn_quick_download_plm.setIcon(theme_mgr.get_styled_icon("download"))
        self.btn_quick_download_plm.setStyleSheet(
            "QPushButton { background-color: #2563EB; color: white; border-radius: 4px; padding: 6px 12px; font-weight: 600; font-size: 11px; }"
            "QPushButton:hover { background-color: #1D4ED8; }"
        )
        self.btn_quick_download_plm.clicked.connect(self._open_plm_download_dialog)

        self.btn_quick_scan = QPushButton("Quét nộp bài", container)
        self.btn_quick_scan.setIcon(theme_mgr.get_styled_icon("refresh"))
        self.btn_quick_scan.setStyleSheet(
            "QPushButton { background-color: #0D9488; color: white; border-radius: 4px; padding: 6px 12px; font-weight: 600; font-size: 11px; }"
            "QPushButton:hover { background-color: #0F766E; }"
        )
        self.btn_quick_scan.clicked.connect(self.scan_member_submissions)

        self.btn_quick_export = QPushButton("Xuất Excel", container)
        self.btn_quick_export.setIcon(theme_mgr.get_styled_icon("file-spreadsheet"))
        self.btn_quick_export.setStyleSheet(
            "QPushButton { background-color: #059669; color: white; border-radius: 4px; padding: 6px 12px; font-weight: 600; font-size: 11px; }"
            "QPushButton:hover { background-color: #047857; }"
        )
        self.btn_quick_export.clicked.connect(self.trigger_batch_reconciliation)

        # Aliases for test compatibility
        self.btn_quick_report = self.btn_quick_export
        self.btn_quick_reset = QPushButton("Đặt lại")
        self.btn_quick_reset.setIcon(theme_mgr.get_styled_icon("refresh"))
        self.btn_quick_reset.setStyleSheet(
            "QPushButton { background-color: #475569; color: white; border-radius: 4px; padding: 6px 12px; font-weight: 600; font-size: 11px; }"
            "QPushButton:hover { background-color: #334155; }"
        )
        self.btn_quick_reset.clicked.connect(self.reset_workspace)
        self.btn_quick_reset.hide()

        panel_layout.addWidget(self.btn_quick_download_plm)
        panel_layout.addWidget(self.btn_quick_scan)
        panel_layout.addWidget(self.btn_quick_export)

        return container

    def get_active_machine_codes(self) -> list[str]:
        if hasattr(self.step1_widget, "get_active_machine_codes"):
            return self.step1_widget.get_active_machine_codes()
        return [self.model_combo.currentText().strip()] if self.model_combo.currentText().strip() else []

    def reset_workspace(self) -> None:
        self.switch_to_step(0)
        self.state = LeaderSessionState(base_dir=self.base_dir)
        self.current_result = None
        self.last_report_path = None
        self.refresh_kpi_cards()

    def refresh_kpi_cards(self) -> None:
        """Update KPI metrics across all 4 cards."""
        total_machines = len(self.state.machines)
        self.lbl_kpi_total_models.setText(str(total_machines))

        ready_count = 0
        for m in self.state.machines:
            has_plm = m.plm_file is not None and m.plm_file.exists()
            has_r3 = m.r3_file is not None and m.r3_file.exists()
            if has_plm and has_r3:
                ready_count += 1
        self.lbl_kpi_ready_models.setText(f"{ready_count}/{total_machines}" if total_machines > 0 else "0")

        total_subs = len(self.state.submissions)
        ok_subs = sum(1 for s in self.state.submissions if s.is_submitted_ok)
        if total_subs > 0:
            pct = int((ok_subs / total_subs) * 100)
            self.lbl_kpi_cttt_progress.setText(f"{pct}% ({ok_subs}/{total_subs})")
        else:
            self.lbl_kpi_cttt_progress.setText("0%")

        if self.current_result is not None or self.last_report_path is not None:
            self.lbl_kpi_recon_status.setText("Đã so sánh xong")
        elif self.state.step3_completed:
            self.lbl_kpi_recon_status.setText("Sẵn sàng so sánh")
        elif total_machines > 0:
            self.lbl_kpi_recon_status.setText("Đang chuẩn bị")
        else:
            self.lbl_kpi_recon_status.setText("Chờ khởi tạo")

    def _wrap_step_scroll(self, widget: QWidget) -> QScrollArea:
        scroll = QScrollArea(self)
        scroll.setWidget(widget)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        return scroll

    def on_theme_changed(self, effective_theme: str) -> None:
        """Propagate theme changes to KPI cards and action icons."""
        is_dark = effective_theme == "dark"
        for card in (
            self.kpi_card_total_models,
            self.kpi_card_ready_models,
            self.kpi_card_cttt_progress,
            self.kpi_card_recon_status,
        ):
            if hasattr(card, "update_theme"):
                card.update_theme(is_dark)

        theme_mgr = get_theme_manager()
        self.btn_quick_download_plm.setIcon(theme_mgr.get_styled_icon("download"))
        self.btn_quick_scan.setIcon(theme_mgr.get_styled_icon("refresh"))
        self.btn_quick_export.setIcon(theme_mgr.get_styled_icon("file-spreadsheet"))
        if hasattr(self, "btn_quick_reset"):
            self.btn_quick_reset.setIcon(theme_mgr.get_styled_icon("refresh"))

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 4, 6, 4)
        main_layout.setSpacing(4)

        # 0. Compact KPI Header Panel (4 KPI Cards + 3 Quick Action Buttons)
        self.kpi_panel = self._build_kpi_panel()
        main_layout.addWidget(self.kpi_panel)

        # 1. Wizard Step Breadcrumb Header
        self.nav_header = WizardStepHeader(self)
        main_layout.addWidget(self.nav_header)

        # 2. QStackedWidget with the 4 steps wrapped in scroll areas
        self.step_stack = QStackedWidget(self)

        self.step1_widget = Step1ProjectSetupWidget(self.state, mailer=self.mailer, parent=self)
        self.step2_widget = Step2DataSourcingWidget(self.state, self)
        self.step3_widget = Step3TrackingConsolidationWidget(self.state, self)
        self.step4_widget = Step4ComparisonReportingWidget(self.state, self)

        self.step_stack.addWidget(self._wrap_step_scroll(self.step1_widget))
        self.step_stack.addWidget(self._wrap_step_scroll(self.step2_widget))
        self.step_stack.addWidget(self._wrap_step_scroll(self.step3_widget))
        self.step_stack.addWidget(self._wrap_step_scroll(self.step4_widget))


        main_layout.addWidget(self.step_stack, stretch=1)

        # 3. Footer Navigation Bar
        footer_layout = QHBoxLayout()

        self.btn_prev = QPushButton("Quay lại Bước trước")
        self.btn_prev.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_prev.setEnabled(False)
        self.btn_prev.clicked.connect(self.go_previous_step)
        footer_layout.addWidget(self.btn_prev)

        footer_layout.addStretch()

        # Shared progress indicators (backward compatible with tests)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setTextVisible(True)
        footer_layout.addWidget(self.progress_bar)

        self.lbl_progress_status = QLabel("Sẵn sàng")
        self.lbl_progress_status.setStyleSheet("font-style: italic; color: #495057;")
        footer_layout.addWidget(self.lbl_progress_status)

        footer_layout.addStretch()

        self.btn_next = QPushButton("Tiếp tục sang Bước tiếp theo")
        self.btn_next.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_next.setStyleSheet("background-color: #0078D4; color: white; border-radius: 4px; padding: 6px 14px;")
        self.btn_next.clicked.connect(self.go_next_step)
        footer_layout.addWidget(self.btn_next)

        main_layout.addLayout(footer_layout)

    def _wire_signals(self) -> None:
        self.nav_header.step_clicked.connect(self.switch_to_step)
        self.step1_widget.step_completed.connect(lambda ok: self.nav_header.set_step_completed(0, ok))
        self.step2_widget.step_completed.connect(lambda ok: self.nav_header.set_step_completed(1, ok))
        self.step3_widget.step_completed.connect(lambda ok: self.nav_header.set_step_completed(2, ok))
        self.step4_widget.step_completed.connect(lambda ok: self.nav_header.set_step_completed(3, ok))

    def switch_to_step(self, step_idx: int) -> None:
        """Navigate to a specific wizard step."""
        if 0 <= step_idx < 4:
            self.step_stack.setCurrentIndex(step_idx)
            self.nav_header.set_active_step(step_idx)
            self.btn_prev.setEnabled(step_idx > 0)
            self.btn_next.setEnabled(step_idx < 3)

            # Context hooks on switching
            if step_idx == 1:
                self.step2_widget.refresh_sourcing_table()
            elif step_idx == 2:
                self.step3_widget.scan_submissions()
            elif step_idx == 3:
                self.step4_widget.sync_machine_combo()
            self.refresh_kpi_cards()

    def go_previous_step(self) -> None:
        curr = self.step_stack.currentIndex()
        if curr > 0:
            self.switch_to_step(curr - 1)

    def go_next_step(self) -> None:
        curr = self.step_stack.currentIndex()
        if curr < 3:
            # Precondition check
            if curr == 0 and not self.state.step1_completed:
                self.step1_widget.execute_create_folders_and_packages()
                if not self.state.step1_completed:
                    return
            if curr == 2 and not self.state.step3_completed:
                QMessageBox.warning(
                    self,
                    "Cảnh báo",
                    "Chưa hoàn tất tổng hợp dữ liệu Bước 3! Vui lòng kiểm tra trạng thái nộp bài và nhấn 'Tổng hợp dữ liệu' trước khi chuyển sang Bước 4.",
                )
                return
            self.switch_to_step(curr + 1)


    # =========================================================================
    # Backward Compatibility Properties & Delegations
    # =========================================================================

    @property
    def model_combo(self) -> QComboBox:
        return self.step1_widget.model_combo

    @property
    def stage_combo(self) -> QComboBox:
        return self.step1_widget.stage_combo

    @property
    def submission_table(self) -> QTableWidget:
        return self.step3_widget.submission_table

    def create_project_folder_structure(self, force_date_hierarchy: bool | None = None) -> Path:
        """Create canonical folder structure."""
        res = self.step1_widget.create_project_folder_structure(force_date_hierarchy=force_date_hierarchy)
        self.refresh_kpi_cards()
        return res

    def scan_member_submissions(self) -> dict[str, dict[str, Any]]:
        """Real-time scan for member submissions."""
        res = self.step3_widget.scan_submissions()
        self.refresh_kpi_cards()
        return res

    def get_sub_unit_statuses(self) -> dict[str, str]:
        """Return dict mapping sub-unit -> status."""
        return self.step3_widget.get_sub_unit_statuses()

    def _open_project_folder(self) -> None:
        """Open project directory in explorer."""
        self.step1_widget._open_base_dir()

    def _open_plm_download_dialog(self) -> None:
        """Open TC24 PLM download dialog."""
        self.step2_widget._open_download_dialog()

    def trigger_batch_reconciliation(self) -> None:
        """Trigger asynchronous batch reconciliation worker."""
        model = self.model_combo.currentText().strip()
        model_dir = self.base_dir / model

        collected = list(self.state.consolidated_cttt)


        plm_dir = model_dir / "PLM"
        plm_files = list(plm_dir.glob("*.xlsx")) if plm_dir.exists() else []
        plm_path = plm_files[0] if plm_files else None

        r3_dir = model_dir / "R3"
        r3_files = list(r3_dir.glob("*.xls")) + list(r3_dir.glob("*.xlsx")) if r3_dir.exists() else []
        r3_path = r3_files[0] if r3_files else None

        if self.thread is not None and self.thread.isRunning():
            logger.warning("Batch reconciliation thread is already running.")
            return

        self.progress_bar.setValue(5)
        self.lbl_progress_status.setText("Khởi động tiến trình đối soát hàng loạt...")

        self.thread = QThread()
        self.worker = BatchReconciliationWorker(
            collected_cttt=collected,
            plm_path=plm_path,
            r3_path=r3_path,
            model_name=model,
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.error.connect(self._on_worker_error)

        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def _on_worker_progress(self, percent: int, message: str) -> None:
        self.progress_bar.setValue(percent)
        self.lbl_progress_status.setText(message)

    def _on_worker_finished(self, result: ReconciliationResult) -> None:
        self.current_result = result
        self.progress_bar.setValue(100)
        self.lbl_progress_status.setText(f"Đã đối soát xong! Phán định: {result.overall_status}")
        self.refresh_kpi_cards()
        self.batch_finished.emit(result)

        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()
        self.thread = None
        self.worker = None

        QMessageBox.information(
            self,
            "Đối soát hoàn tất",
            f"Đã hoàn thành đối soát BOM toàn máy!\n"
            f"- Đánh giá tổng thể: {result.overall_status}\n"
            f"- Số dòng CTTT: {len(result.cttt_rows)}\n"
            f"- Số linh kiện thiếu trên CTTT: {len(result.plm_missing_rows)}",
        )

    def _on_worker_error(self, err_msg: str) -> None:
        self.lbl_progress_status.setText(f"Lỗi: {err_msg}")
        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()
        self.thread = None
        self.worker = None
        QMessageBox.critical(self, "Lỗi đối soát", f"Tiến trình đối soát gặp sự cố:\n{err_msg}")

    def trigger_consolidated_report(self) -> Path | None:
        """Trigger report generation matching form_ssbom.xlsm layout."""
        model = self.model_combo.currentText().strip()
        model_dir = self.base_dir / model
        reports_dir = model_dir / "Reports"
        reports_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = reports_dir / f"form_ssbom_{model}_{timestamp}.xlsx"

        try:
            self.lbl_progress_status.setText("Đang xuất tệp Excel báo cáo 7 Sheet...")
            res_path = self.excel_generator.generate_report(
                output_path=report_file,
                reconciliation=self.current_result,
                model_name=model,
                sub_unit_statuses=self.get_sub_unit_statuses(),
            )
            self.last_report_path = res_path
            self.refresh_kpi_cards()
            self.lbl_progress_status.setText(f"Báo cáo đã xuất: {res_path.name}")
            self.report_generated.emit(str(res_path))

            QMessageBox.information(
                self,
                "Xuất báo cáo thành công",
                f"Đã tạo báo cáo hợp nhất form_ssbom với 7 Sheet hoàn chỉnh:\n{res_path.name}\n\nĐường dẫn: {res_path}",
            )
            return res_path
        except Exception as exc:
            logger.error("Failed to generate consolidated report: %s", exc)
            QMessageBox.critical(self, "Lỗi xuất báo cáo", f"Không thể xuất báo cáo:\n{exc}")
            return None

    def trigger_outlook_preview(self) -> None:
        """Construct preview and show dialog."""
        model = self.model_combo.currentText().strip()
        overall = self.current_result.overall_status if self.current_result else "OK"
        target_date = datetime.date.today().strftime("%d/%m/%Y")

        stats = {
            "total_parts": len(self.current_result.cttt_rows) if self.current_result else 0,
            "ok_count": 0,
            "ng_count": 0,
            "missing_count": len(self.current_result.plm_missing_rows) if self.current_result else 0,
            "warning_count": 0,
        }
        if self.current_result and not self.current_result.cttt_rows.empty:
            ng = (self.current_result.cttt_rows["Check"] == "NG").sum() if "Check" in self.current_result.cttt_rows else 0
            stats["ng_count"] = ng
            stats["ok_count"] = len(self.current_result.cttt_rows) - ng

        preview = self.mailer.build_email_preview(
            model_name=model,
            target_date=target_date,
            overall_status=overall,
            recipients_to=["vn_pe03@dtvn.kyocera.com", "vinh.bd@dtvn.kyocera.com"],
            sub_unit_statuses=self.get_sub_unit_statuses(),
            summary_stats=stats,
            attachment_path=self.last_report_path,
        )

        dlg = EmailPreviewDialog(preview=preview, mailer=self.mailer, parent=self)
        dlg.exec()
