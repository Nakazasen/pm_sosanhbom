"""Dialog for scanning PCD Monthly Production Plan and Auto-generating Project Folders."""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Any, List, Optional

from PyQt6.QtCore import QDate, QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.member_database import MemberDatabaseManager
from src.gui.styles.theme_manager import get_theme_manager
from src.reporting.outlook_mailer import EmailPreview, OutlookMailer
from src.services.machine_dict_service import DEFAULT_DICT_PATH, MachineDictService, MachineInfo
from src.services.pcd_plan_service import PCDPlanItem, PCDPlanScanResult, PCDPlanService

logger = logging.getLogger(__name__)

DEFAULT_BOM_STORAGE_ROOT = (
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept"
    r"\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3"
)
DEFAULT_MEMBER_TEMPLATE = (
    r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)"
    r"\Hang muc can luu\Vinh\Pm_sosanhBOM\formnguoidung.xlsm"
)


class PCDSharePointDownloadWorker(QThread):
    """Background worker to download PCD plan from SharePoint Online using headless Edge."""

    progress = pyqtSignal(str)
    finished = pyqtSignal(Path)
    error = pyqtSignal(str)

    def __init__(self, year: int, month: int, parent: Optional[QObject] = None):
        super().__init__(parent)
        self.year = year
        self.month = month

    def run(self) -> None:
        try:
            from src.services.pcd_sharepoint_automation import download_pcd_plan_from_sharepoint

            file_path = download_pcd_plan_from_sharepoint(
                year=self.year,
                month=self.month,
                progress_callback=lambda msg: self.progress.emit(msg),
            )
            self.finished.emit(file_path)
        except Exception as exc:
            self.error.emit(str(exc))


class PCDPlanScanDialog(QDialog):
    """Dialog to scan monthly PCD production plan, map machine types, and auto-setup project folders."""

    projects_created = pyqtSignal(list)  # Emits list of created project directories

    def __init__(
        self,
        plan_service: PCDPlanService | None = None,
        dict_service: MachineDictService | None = None,
        mailer: OutlookMailer | None = None,
        db_manager: MemberDatabaseManager | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.plan_service = plan_service or PCDPlanService()
        self.dict_service = dict_service or MachineDictService()
        self.mailer = mailer or OutlookMailer()
        self.db_manager = db_manager or MemberDatabaseManager(base_dir=getattr(parent, "base_dir", None))

        self.setWindowTitle("Quét Kế Hoạch Sản Xuất Tháng PCD & Khởi Tạo Dự Án Tự Động")
        self.resize(1020, 680)

        self._scanned_items: list[PCDPlanItem] = []
        self._init_ui()
        self._auto_detect_plan()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        # 1. Top Section: Date & File Selection
        top_group = QGroupBox("1. Thông tin Kế hoạch Sản xuất Tháng (PCD)")
        top_layout = QVBoxLayout(top_group)
        top_layout.setContentsMargins(8, 8, 8, 8)
        top_layout.setSpacing(6)

        h_date = QHBoxLayout()
        h_date.addWidget(QLabel("Năm:"))
        self.spin_year = QSpinBox()
        self.spin_year.setRange(2020, 2035)
        self.spin_year.setValue(datetime.datetime.now().year)
        h_date.addWidget(self.spin_year)

        h_date.addWidget(QLabel("Tháng:"))
        self.combo_month = QComboBox()
        for m in range(1, 13):
            self.combo_month.addItem(f"Tháng {m:02d}", m)
        self.combo_month.setCurrentIndex(datetime.datetime.now().month - 1)
        h_date.addWidget(self.combo_month)

        self.btn_auto_find = QPushButton("Tìm theo ngày hiện tại")
        self.btn_auto_find.setIcon(get_theme_manager().get_styled_icon("search"))
        self.btn_auto_find.clicked.connect(self._auto_detect_plan)
        h_date.addWidget(self.btn_auto_find)

        h_date.addStretch()

        self.lbl_file_type = QLabel("Chưa phát hiện tệp kế hoạch")
        self.lbl_file_type.setStyleSheet("font-weight: bold; color: #666;")
        h_date.addWidget(self.lbl_file_type)

        top_layout.addLayout(h_date)

        h_file = QHBoxLayout()
        h_file.addWidget(QLabel("Tệp kế hoạch Excel:"))
        self.edit_plan_file = QLineEdit()
        self.edit_plan_file.setPlaceholderText("Đường dẫn tệp (*定期計画後明細計画.xlsx hoặc *月初明細計画.xlsx)...")
        h_file.addWidget(self.edit_plan_file, stretch=1)

        self.btn_browse_plan = QPushButton("Chọn tệp...")
        self.btn_browse_plan.clicked.connect(self._browse_plan_file)
        h_file.addWidget(self.btn_browse_plan)

        self.btn_auto_download = QPushButton("🚀 Tải Tự Động Từ SharePoint")
        self.btn_auto_download.setToolTip("Tự động điều khiển trình duyệt Edge chạy ngầm kết nối SharePoint PCD, tìm và tải file kế hoạch của tháng được chọn")
        self.btn_auto_download.setStyleSheet(
            "QPushButton { background-color: #107C41; color: white; font-weight: bold; padding: 4px 12px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #0B5A2F; }"
            "QPushButton:disabled { background-color: #94A3B8; color: #FFFFFF; }"
        )
        self.btn_auto_download.clicked.connect(self._on_auto_download_clicked)
        h_file.addWidget(self.btn_auto_download)

        self.btn_open_sharepoint = QPushButton("🌐 Mở SharePoint PCD...")
        self.btn_open_sharepoint.setToolTip("Mở trang web Kế hoạch R3 tổng của phòng Quản lý sản xuất (PCD) trên trình duyệt")
        self.btn_open_sharepoint.setStyleSheet(
            "QPushButton { background-color: #0F6CBD; color: white; font-weight: bold; padding: 4px 10px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #115EA3; }"
        )
        self.btn_open_sharepoint.clicked.connect(self._on_open_sharepoint_clicked)
        h_file.addWidget(self.btn_open_sharepoint)

        self.btn_scan = QPushButton("Quét Dữ Liệu (Sheet 詳細日程)")
        self.btn_scan.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_scan.setStyleSheet("background-color: #0078D4; color: white; padding: 4px 14px; border-radius: 4px;")
        self.btn_scan.clicked.connect(self._on_scan_clicked)
        h_file.addWidget(self.btn_scan)

        top_layout.addLayout(h_file)

        # Progress and status bar for SharePoint download
        h_status = QHBoxLayout()
        self.pbar_download = QProgressBar()
        self.pbar_download.setRange(0, 0)
        self.pbar_download.setFixedHeight(12)
        self.pbar_download.setVisible(False)
        h_status.addWidget(self.pbar_download, stretch=1)

        self.lbl_download_status = QLabel("")
        self.lbl_download_status.setStyleSheet("font-size: 11px;")
        h_status.addWidget(self.lbl_download_status)
        top_layout.addLayout(h_status)

        lbl_hint = QLabel(
            "💡 <b>Tự động hóa phòng ban:</b> Bấm <b>'🚀 Tải Tự Động Từ SharePoint'</b> để phần mềm tự động lấy file kế hoạch mới nhất từ SharePoint PCD mà không cần cấu hình thủ công. Hoặc mở SharePoint và bấm <b>'Thêm lối tắt vào OneDrive'</b> để tự đồng bộ về máy."
        )
        lbl_hint.setWordWrap(True)
        lbl_hint.setStyleSheet("color: #94A3B8; font-size: 11px; padding: 2px 4px;")
        top_layout.addWidget(lbl_hint)

        layout.addWidget(top_group)

        # 2. Results Table: Detected NEW materials
        mid_group = QGroupBox("2. Danh sách Mã Máy Mới (履歴 = 'NEW' & Mã bắt đầu bằng 'T1' hoặc '11')")
        mid_layout = QVBoxLayout(mid_group)
        mid_layout.setContentsMargins(8, 8, 8, 8)
        mid_layout.setSpacing(6)

        h_tools = QHBoxLayout()
        self.btn_select_all = QPushButton("Chọn tất cả")
        self.btn_select_all.clicked.connect(lambda: self._set_all_rows_checked(True))
        self.btn_deselect_all = QPushButton("Bỏ chọn")
        self.btn_deselect_all.clicked.connect(lambda: self._set_all_rows_checked(False))
        self.btn_reload_dict = QPushButton("Tải lại từ điển Loại máy")
        self.btn_reload_dict.setIcon(get_theme_manager().get_styled_icon("refresh-cw"))
        self.btn_reload_dict.clicked.connect(self._reload_dictionary)

        self.btn_add_to_dict = QPushButton("+ Thêm mã vào từ điển")
        self.btn_add_to_dict.setToolTip("Ghi ngược mã máy mới vào file_loaimay_nhommail.xlsx")
        self.btn_add_to_dict.clicked.connect(self._open_add_to_dict_dialog)

        h_tools.addWidget(self.btn_select_all)
        h_tools.addWidget(self.btn_deselect_all)
        h_tools.addWidget(self.btn_reload_dict)
        h_tools.addWidget(self.btn_add_to_dict)
        h_tools.addStretch()

        self.lbl_stats = QLabel("Chưa có dữ liệu.")
        self.lbl_stats.setStyleSheet("font-weight: bold; color: #1F497D;")
        h_tools.addWidget(self.lbl_stats)

        mid_layout.addLayout(h_tools)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            "Chọn",
            "Mã Vật Tư (Material)",
            "Mã Máy (4 ký tự)",
            "Trạng Thái",
            "Tên Loại Máy (Từ Điển)",
            "Giai Đoạn (Phase)",
            "Đường Dẫn Lưu Trữ Dự Kiến",
        ])
        self.table.setShowGrid(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        h_header = self.table.horizontalHeader()
        h_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 45)
        h_header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h_header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        h_header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        h_header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        h_header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        h_header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)

        mid_layout.addWidget(self.table)
        layout.addWidget(mid_group)

        # 3. Bottom Action Bar
        bot_group = QGroupBox("3. Thiết lập Khởi tạo & Thông báo Email")
        bot_layout = QVBoxLayout(bot_group)
        bot_layout.setContentsMargins(8, 6, 8, 6)
        bot_layout.setSpacing(6)

        h_opt = QHBoxLayout()
        self.chk_send_email = QCheckBox("Gửi Email Thông Báo Yêu Cầu Phụ Trách Sau Khi Tạo Thư Mục")
        self.chk_send_email.setChecked(True)
        h_opt.addWidget(self.chk_send_email)

        self.chk_test_mode = QCheckBox("Chế độ Thử Nghiệm (Test Mode - Gửi tới vinh.bd@dtvn.kyocera.com)")
        self.chk_test_mode.setChecked(True)
        h_opt.addWidget(self.chk_test_mode)

        h_opt.addStretch()
        bot_layout.addLayout(h_opt)

        h_action = QHBoxLayout()
        self.btn_execute = QPushButton("Khởi Tạo Thư Mục & Sinh Gói File Làm Việc Cho Mã Đã Chọn")
        self.btn_execute.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        self.btn_execute.setMinimumHeight(36)
        self.btn_execute.setStyleSheet(
            "background-color: #2E7D32; color: white; padding: 6px 18px; border-radius: 4px; font-weight: bold;"
        )
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setMinimumHeight(36)
        self.btn_close.clicked.connect(self.reject)

        h_action.addWidget(self.btn_execute, stretch=1)
        h_action.addWidget(self.btn_close)
        bot_layout.addLayout(h_action)

        layout.addWidget(bot_group)

    def _auto_detect_plan(self) -> None:
        """Attempt to locate the plan file automatically from OneDrive or local paths."""
        year = self.spin_year.value()
        month = self.combo_month.currentData()

        found = self.plan_service.find_monthly_plan_file(year=year, month=month)
        if found:
            p, ftype = found
            self.edit_plan_file.setText(str(p))
            if ftype == "end_of_period":
                self.lbl_file_type.setText("⭐ Ưu tiên: File kế hoạch cuối kỳ (*定期計画後明細計画)")
                self.lbl_file_type.setStyleSheet("color: #2E7D32; font-weight: bold;")
            elif ftype == "beginning_of_period":
                self.lbl_file_type.setText("📄 File kế hoạch đầu kỳ (*月初明細計画)")
                self.lbl_file_type.setStyleSheet("color: #0078D4; font-weight: bold;")
            else:
                self.lbl_file_type.setText("📄 Đã phát hiện tệp kế hoạch")
                self.lbl_file_type.setStyleSheet("color: #333; font-weight: bold;")
        else:
            self.edit_plan_file.clear()
            self.lbl_file_type.setText("⚠️ Chưa tìm thấy file kế hoạch cho tháng đang chọn. Vui lòng bấm 'Chọn tệp...'.")
            self.lbl_file_type.setStyleSheet("color: #D32F2F;")

    def _browse_plan_file(self) -> None:
        start_dir = str(self.plan_service.base_dir) if self.plan_service.base_dir.exists() else ""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn tệp Kế hoạch Sản xuất Tháng (*.xlsx)",
            start_dir,
            "Excel Files (*.xlsx *.xlsm);;All Files (*.*)",
        )
        if path:
            self.edit_plan_file.setText(path)
            self._save_pcd_base_dir_to_settings(Path(path))
            if "定期計画後明細計画" in path:
                self.lbl_file_type.setText("⭐ Ưu tiên: File kế hoạch cuối kỳ (*定期計画後明細計画)")
                self.lbl_file_type.setStyleSheet("color: #2E7D32; font-weight: bold;")
            elif "月初明細計画" in path:
                self.lbl_file_type.setText("📄 File kế hoạch đầu kỳ (*月初明細計画)")
                self.lbl_file_type.setStyleSheet("color: #0078D4; font-weight: bold;")
            else:
                self.lbl_file_type.setText("📄 Đã chọn tệp kế hoạch thủ công")
                self.lbl_file_type.setStyleSheet("color: #333; font-weight: bold;")

    def _save_pcd_base_dir_to_settings(self, file_path: Path) -> None:
        """Persist directory of selected plan file to settings.json so next scans find it automatically."""
        import json
        try:
            cfg_path = Path("config/settings.json")
            if not cfg_path.exists():
                cfg_path = Path("../config/settings.json")
            cfg: dict[str, Any] = {}
            if cfg_path.exists():
                with open(cfg_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            if "paths" not in cfg or not isinstance(cfg["paths"], dict):
                cfg["paths"] = {}
            parent_dir = str(file_path.parent)
            cfg["paths"]["pcd_base_dir"] = parent_dir
            cfg_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = cfg_path.with_suffix(".tmp")
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
            tmp.replace(cfg_path)
            logger.info("Saved custom pcd_base_dir to %s: %s", cfg_path, parent_dir)
        except Exception as e:
            logger.debug("Could not save pcd_base_dir to settings: %s", e)

    def _on_auto_download_clicked(self) -> None:
        """Trigger background browser automation to download plan from SharePoint."""
        year = self.spin_year.value()
        month = self.combo_month.currentData()

        self.btn_auto_download.setEnabled(False)
        self.btn_scan.setEnabled(False)
        self.btn_browse_plan.setEnabled(False)
        self.pbar_download.setVisible(True)
        self.lbl_download_status.setText(f"Đang kết nối SharePoint PCD Tháng {month:02d}/{year}...")
        self.lbl_download_status.setStyleSheet("color: #0F6CBD; font-weight: bold;")

        self._download_worker = PCDSharePointDownloadWorker(year=year, month=month, parent=self)
        self._download_worker.progress.connect(self._on_download_progress)
        self._download_worker.finished.connect(self._on_download_finished)
        self._download_worker.error.connect(self._on_download_error)
        self._download_worker.start()

    def _on_download_progress(self, msg: str) -> None:
        """Update status label with current download step."""
        clean_msg = msg.replace("[*]", "").replace("[+]", "").strip()
        self.lbl_download_status.setText(clean_msg)

    def _on_download_finished(self, file_path: Path) -> None:
        """Handle successful SharePoint download: populate input, save dir, and scan."""
        self.pbar_download.setVisible(False)
        self.btn_auto_download.setEnabled(True)
        self.btn_scan.setEnabled(True)
        self.btn_browse_plan.setEnabled(True)
        self.lbl_download_status.setText(f"✓ Đã tải xong: {file_path.name}")
        self.lbl_download_status.setStyleSheet("color: #2E7D32; font-weight: bold;")

        self.edit_plan_file.setText(str(file_path))
        self._save_pcd_base_dir_to_settings(file_path)
        # Automatically trigger scanning the downloaded plan file!
        self._on_scan_clicked()

    def _on_download_error(self, err_msg: str) -> None:
        """Handle download failure with actionable notice."""
        self.pbar_download.setVisible(False)
        self.btn_auto_download.setEnabled(True)
        self.btn_scan.setEnabled(True)
        self.btn_browse_plan.setEnabled(True)
        self.lbl_download_status.setText("Tải thất bại.")
        self.lbl_download_status.setStyleSheet("color: #D32F2F; font-weight: bold;")
        QMessageBox.warning(
            self,
            "Tải Kế Hoạch SharePoint",
            f"Không thể tự động tải tệp kế hoạch từ SharePoint PCD:\n\n{err_msg}\n\n"
            f"Bạn có thể bấm '🌐 Mở SharePoint PCD...' để kiểm tra trên trình duyệt hoặc tự chọn tệp kế hoạch bằng nút 'Chọn tệp...'.",
        )

    def _on_open_sharepoint_clicked(self) -> None:
        """Open official PCD SharePoint library in default web browser."""
        from src.services.pcd_plan_service import open_pcd_sharepoint_url, OFFICIAL_PCD_SHAREPOINT_URL
        success = open_pcd_sharepoint_url()
        if not success:
            QMessageBox.warning(
                self,
                "Không thể mở trình duyệt",
                f"Không thể tự động mở trình duyệt web. Vui lòng truy cập đường dẫn:\n{OFFICIAL_PCD_SHAREPOINT_URL}",
            )

    def _on_scan_clicked(self) -> None:
        file_path = self.edit_plan_file.text().strip()
        if not file_path or not Path(file_path).exists():
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn tệp Kế hoạch Sản xuất (*.xlsx) hợp lệ!")
            return

        try:
            self.dict_service.load()
            result = self.plan_service.parse_plan_file(file_path)
            self._scanned_items = result.items
            self._populate_table()
            if self._scanned_items:
                self.lbl_stats.setText(f"Đã tìm thấy {len(self._scanned_items)} mã máy mới (NEW) thỏa mãn điều kiện.")
            else:
                self.lbl_stats.setText("Quét xong: Không có mã máy mới (NEW) nào trong tệp này.")
                QMessageBox.information(
                    self,
                    "Kết quả quét",
                    f"Đã đọc thành công tệp nhưng không tìm thấy dòng nào thỏa mãn:\n"
                    f"• Cột H (履歴) = 'NEW'\n"
                    f"• Cột D (Material) bắt đầu bằng '11' hoặc 'T1'\n\n"
                    f"Tổng số dòng đã quét: {result.total_rows_scanned}.",
                )
        except Exception as exc:
            logger.error("Error scanning plan file: %s", exc)
            self._handle_scan_error(file_path, exc)

    def _handle_scan_error(self, file_path: str, exc: Exception) -> None:
        p = Path(file_path)
        from src.services.pcd_plan_service import is_cloud_placeholder, is_onedrive_running, start_onedrive

        if is_cloud_placeholder(p) and not is_onedrive_running():
            reply = QMessageBox.question(
                self,
                "Tệp chưa tải về từ OneDrive",
                f"Tệp '{p.name}' đang ở chế độ lưu trữ trực tuyến trên đám mây (Cloud-only) "
                f"và ứng dụng OneDrive hiện chưa được mở trên máy tính.\n\n"
                f"Bạn có muốn khởi động OneDrive ngay bây giờ để tải tệp về máy không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                started = start_onedrive()
                if started:
                    QMessageBox.information(
                        self,
                        "Đã khởi động OneDrive",
                        "Đã khởi động tiến trình OneDrive trong nền.\n"
                        "Vui lòng đợi vài giây để OneDrive đồng bộ tệp về máy rồi bấm 'Quét Dữ Liệu' lại.",
                    )
                else:
                    QMessageBox.warning(
                        self,
                        "Không thể mở OneDrive",
                        "Không tìm thấy file thực thi OneDrive.exe. Vui lòng mở OneDrive từ Start Menu của Windows.",
                    )
            return

        QMessageBox.critical(
            self,
            "Lỗi quét tệp kế hoạch",
            f"Không thể đọc tệp kế hoạch:\n{p.name}\n\n"
            f"Nguyên nhân:\n{exc}\n\n"
            f"Gợi ý khắc phục:\n"
            f"1. Đảm bảo bạn đã chọn đúng tệp Kế hoạch Sản xuất Tháng (có chứa sheet '詳細日程').\n"
            f"2. Nếu tệp đang lưu trên OneDrive, hãy mở tệp trực tiếp trong Microsoft Excel để tải về máy trước khi quét.",
        )

    def _populate_table(self) -> None:
        self.table.setRowCount(0)
        year = self.spin_year.value()
        month = self.combo_month.currentData()
        year_month_str = f"{year}.{month:02d}"

        self.table.setRowCount(len(self._scanned_items))

        for row_idx, item in enumerate(self._scanned_items):
            # Col 0: Checkbox
            chk = QCheckBox()
            chk.setChecked(True)
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk_layout.addWidget(chk)
            self.table.setCellWidget(row_idx, 0, chk_widget)

            # Col 1: Material Code
            item_mat = QTableWidgetItem(item.material_code)
            item_mat.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
            self.table.setItem(row_idx, 1, item_mat)

            # Col 2: Machine Code (4 chars)
            item_code = QTableWidgetItem(item.machine_code_4char)
            self.table.setItem(row_idx, 2, item_code)

            # Col 3: Status
            item_st = QTableWidgetItem(item.history_status)
            item_st.setForeground(QColor("#2E7D32"))
            self.table.setItem(row_idx, 3, item_st)

            # Col 4: Machine Name / Model
            model_name = item.machine_name or "(Chưa định nghĩa)"
            item_model = QTableWidgetItem(model_name)
            if not item.machine_name:
                item_model.setForeground(QColor("#C62828"))
                item_model.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
            self.table.setItem(row_idx, 4, item_model)

            # Col 5: Phase selector (Combo)
            phase_combo = QComboBox()
            if item.is_trial:
                phase_combo.addItems(["DMT", "PMT", "PP", "MP"])
                phase_combo.setCurrentText("DMT")
            else:
                phase_combo.addItems(["MP", "PP", "DMT", "PMT"])
                phase_combo.setCurrentText("MP")
            self.table.setCellWidget(row_idx, 5, phase_combo)

            # Col 6: Expected Storage Path
            clean_model = item.machine_name or "Unknown_Model"
            phase_str = phase_combo.currentText()
            expected_path = f"{DEFAULT_BOM_STORAGE_ROOT}\\{clean_model}\\{phase_str}\\{year_month_str}\\{item.material_code}"
            item_path = QTableWidgetItem(expected_path)
            item_path.setToolTip(expected_path)
            self.table.setItem(row_idx, 6, item_path)

            # Update path on phase change
            phase_combo.currentTextChanged.connect(
                lambda val, r=row_idx, m=clean_model, ym=year_month_str, mat=item.material_code: self._update_row_path(r, m, val, ym, mat)
            )

    def _update_row_path(self, row: int, model: str, phase: str, ym: str, mat: str) -> None:
        expected_path = f"{DEFAULT_BOM_STORAGE_ROOT}\\{model}\\{phase}\\{ym}\\{mat}"
        item_path = self.table.item(row, 6)
        if item_path:
            item_path.setText(expected_path)
            item_path.setToolTip(expected_path)

    def _set_all_rows_checked(self, checked: bool) -> None:
        for r in range(self.table.rowCount()):
            w = self.table.cellWidget(r, 0)
            if w:
                chk = w.findChild(QCheckBox)
                if chk:
                    chk.setChecked(checked)

    def _reload_dictionary(self) -> None:
        if self.dict_service.load(force_reload=True):
            QMessageBox.information(self, "Thông báo", "Đã tải lại từ điển Loại máy thành công!")
            # Re-match items
            for item in self._scanned_items:
                matched = self.dict_service.extract_and_lookup_material(item.material_code)
                item.machine_info = matched
                item.machine_name = matched.machine_name if matched else ""
            self._populate_table()
        else:
            QMessageBox.warning(self, "Lỗi", "Không thể tải lại từ điển từ file_loaimay_nhommail.xlsx!")

    def _open_add_to_dict_dialog(self) -> None:
        """Dialog to add a new machine code and write back 2-way to file_loaimay_nhommail.xlsx."""
        selected_rows = self.table.selectionModel().selectedRows()
        default_code = ""
        if selected_rows:
            r = selected_rows[0].row()
            default_code = self.table.item(r, 2).text() if self.table.item(r, 2) else ""

        dlg = QDialog(self)
        dlg.setWindowTitle("Bổ Sung Mã Máy Vào Từ Điển (Ghi 2 chiều)")
        dlg.resize(450, 220)
        d_layout = QVBoxLayout(dlg)

        f_layout = QHBoxLayout()
        f_layout.addWidget(QLabel("Mã máy (4 ký tự):"))
        edit_code = QLineEdit(default_code)
        f_layout.addWidget(edit_code)
        d_layout.addLayout(f_layout)

        f_model = QHBoxLayout()
        f_model.addWidget(QLabel("Tên Loại máy:"))
        combo_model = QComboBox()
        combo_model.setEditable(True)
        # Populate existing models
        known_models = sorted(list({m.machine_name for m in self.dict_service.get_all_machines() if m.machine_name}))
        combo_model.addItems(known_models)
        f_model.addWidget(combo_model)
        d_layout.addLayout(f_model)

        f_var = QHBoxLayout()
        f_var.addWidget(QLabel("Kiểu / Biến thể:"))
        edit_var = QLineEdit()
        f_var.addWidget(edit_var)
        d_layout.addLayout(f_var)

        btn_box = QHBoxLayout()
        btn_save = QPushButton("Lưu vào File Excel Từ Điển")
        btn_save.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold; padding: 6px;")
        btn_cancel = QPushButton("Hủy")
        btn_cancel.clicked.connect(dlg.reject)

        def _do_save():
            c = edit_code.text().strip().upper()
            m = combo_model.currentText().strip()
            v = edit_var.text().strip()
            if not c or not m:
                QMessageBox.warning(dlg, "Thiếu thông tin", "Vui lòng nhập đầy đủ Mã máy và Tên loại máy!")
                return
            success = self.dict_service.add_or_update_machine_code(machine_name=m, new_code=c, variant=v)
            if success:
                QMessageBox.information(dlg, "Thành công", f"Đã bổ sung mã {c} cho loại máy {m} vào file_loaimay_nhommail.xlsx!")
                dlg.accept()
                self._reload_dictionary()
            else:
                QMessageBox.critical(dlg, "Lỗi", "Không thể ghi dữ liệu vào file Excel từ điển!")

        btn_save.clicked.connect(_do_save)
        btn_box.addWidget(btn_save)
        btn_box.addWidget(btn_cancel)
        d_layout.addLayout(btn_box)
        dlg.exec()

    def _on_execute_clicked(self) -> None:
        """Execute folder creation and member package distribution for selected items."""
        selected_indices = []
        for r in range(self.table.rowCount()):
            w = self.table.cellWidget(r, 0)
            if w:
                chk = w.findChild(QCheckBox)
                if chk and chk.isChecked():
                    selected_indices.append(r)

        if not selected_indices:
            QMessageBox.warning(self, "Chưa chọn", "Vui lòng tích chọn ít nhất 1 mã máy để khởi tạo dự án!")
            return

        # Prepare template
        template_path = Path(DEFAULT_MEMBER_TEMPLATE)
        if not template_path.exists():
            logger.warning("Default member template not found at %s", template_path)

        created_folders: list[Path] = []
        skipped_count = 0
        now_str = datetime.datetime.now().strftime("%d/%m")

        for r in selected_indices:
            mat_item = self.table.item(r, 1)
            mat_code = mat_item.text().strip() if mat_item else ""
            model_item = self.table.item(r, 4)
            model_name = model_item.text().strip() if model_item else "Unknown_Model"
            phase_combo = self.table.cellWidget(r, 5)
            phase = phase_combo.currentText().strip() if isinstance(phase_combo, QComboBox) else "MP"
            path_item = self.table.item(r, 6)
            folder_str = path_item.text().strip() if path_item else ""

            if not folder_str:
                continue

            target_folder = Path(folder_str)

            # Check if directory already exists
            if target_folder.exists():
                reply = QMessageBox.question(
                    self,
                    "Xác nhận Thư mục Đã Tồn Tại",
                    f"Thư mục dự án đã tồn tại:\n{target_folder}\n\n"
                    f"Bạn có muốn ghi đè / tải lại PLM và R3 mới nhất không?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply != QMessageBox.StandardButton.Yes:
                    skipped_count += 1
                    continue

            try:
                target_folder.mkdir(parents=True, exist_ok=True)
                created_folders.append(target_folder)

                # Generate member packages from template based on assigned machine models
                matched_members = self.db_manager.get_members_for_machine(model_name)
                if matched_members:
                    engineers = [m.account_id for m in matched_members]
                else:
                    # Fallback to active members in database if none specifically matched, or default sample
                    active_members = self.db_manager.get_members(active_only=True)
                    engineers = [m.account_id for m in active_members[:4]] if active_members else [
                        "Toan_mecha",
                        "Hung_mecha1",
                        "Huong_mecha2",
                        "Dien_electrical",
                    ]
                for eng_name in engineers:
                    dest_file = target_folder / f"{eng_name}.xlsm"
                    if template_path.exists():
                        import shutil
                        shutil.copy2(str(template_path), str(dest_file))
                    else:
                        # Fallback empty workbook
                        import openpyxl
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "CTTT"
                        ws["A1"] = f"Mã Máy: {mat_code}"
                        ws["F1"] = f"Phụ trách: {eng_name}"
                        wb.save(str(dest_file))
                        wb.close()

            except Exception as e:
                logger.error("Error creating folder %s: %s", target_folder, e)
                QMessageBox.critical(self, "Lỗi khởi tạo", f"Không thể tạo thư mục:\n{target_folder}\nLỗi: {e}")

        msg = f"Đã khởi tạo thành công {len(created_folders)} thư mục dự án và file phụ trách!"
        if skipped_count > 0:
            msg += f"\n(Đã bỏ qua {skipped_count} thư mục có sẵn theo lựa chọn của bạn)"
        QMessageBox.information(self, "Khởi tạo Hoàn tất", msg)

        # Prompt to send Email if checked - strictly separated by Model!
        if self.chk_send_email.isChecked() and created_folders:
            self.mailer.test_mode = self.chk_test_mode.isChecked()

            from collections import defaultdict
            model_groups: dict[str, list[tuple[str, Path]]] = defaultdict(list)
            for r in selected_indices:
                model_item = self.table.item(r, 4)
                mod_name = model_item.text().strip() if model_item else "Unknown_Model"
                phase_combo = self.table.cellWidget(r, 5)
                phase_str = phase_combo.currentText().strip() if isinstance(phase_combo, QComboBox) else "MP"
                path_item = self.table.item(r, 6)
                p = Path(path_item.text().strip()) if path_item else Path()
                model_groups[mod_name].append((phase_str, p))

            from src.gui.leader_view import EmailPreviewDialog
            total_groups = len(model_groups)
            for idx, (mod_name, row_list) in enumerate(model_groups.items(), start=1):
                qty = len(row_list)
                first_phase, first_folder = row_list[0]
                # The attachment path should be the model's base folder (parent of the machine code folder)
                att_dir = first_folder.parent if first_folder.parent and first_folder.parent != first_folder else first_folder

                preview = self.mailer.build_task_assignment_email(
                    machine_type=mod_name,
                    start_date=now_str,
                    quantity=qty,
                    phase=first_phase,
                    deadline_copy=(datetime.datetime.now() + datetime.timedelta(days=3)).strftime("%d.%m.%Y"),
                    deadline_verify=(datetime.datetime.now() + datetime.timedelta(days=5)).strftime("%d.%m.%Y"),
                    attachment_path=att_dir,
                )
                dlg = EmailPreviewDialog(preview=preview, mailer=self.mailer, parent=self)
                dlg.setWindowTitle(f"Xem trước Email Phân Công ({idx}/{total_groups}) - Model: {mod_name} ({qty} mã)")
                dlg.exec()

        self.projects_created.emit(created_folders)
        self.accept()
