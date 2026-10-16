"""BOM Update Dialog (Cập Nhật BOM Mới - PLM & SAP R3).

Direct implementation and modernization of Buttons 7 & 8 from legacy PE Dept VBA:
- Nút 7: '7. Cập nhật PLM mới' (Sub capnhat_PLM):
  * Ingests new PLM Teamcenter file into target BOM comparison workbook.
  * Archives current Sheet 'PLM' to 'PLM_old' (or 'PLM_old_1', 'PLM_old_2'...).
  * Vectorized legacy ham_match_index_mix inherits:
    - Col O: Giải thích (Explanation)
    - Col P: Người phụ trách (Engineer)
    - Col Q: Quản lý check (Manager review status)
  * Moves processed PLM file into 'capnhat\\old\\'.
  * Refreshes Pivot Table pivot_table_plm.
- Nút 8: '8. Cập nhật R3 mới' (Sub capnhat_R3):
  * Ingests new SAP R3 CS12 file into target BOM workbook.
  * Normalizes column structures (supports both Virgo CS12 and standard layout).
  * Moves processed R3 file into 'capnhat\\old\\'.
  * Rebuilds PivotTable3 at W1 and refreshes Pivot cache.
- Standalone / Multi-entry usage:
  * Can be triggered from Main Toolbar, Tools menu (Ctrl+U), Member View, or Leader View.
  * Supports auto-scanning machine folder (identical to VBA) or custom file selection.
  * Non-blocking background execution with live log console.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
import shutil
import subprocess
from typing import Any, Optional

from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from src.core.inheritance_engine import InheritanceEngine, InheritanceSummary
from src.gui.styles import (
    apply_smart_dialog_geometry,
    ensure_window_fits_screen,
    get_theme_manager,
)

logger = logging.getLogger(__name__)


class BOMUpdateWorker(QThread):
    """Worker thread running BOM update operations without blocking the GUI."""

    progress_signal = pyqtSignal(int, str)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, dict)

    def __init__(
        self,
        bom_file: Path,
        plm_file: Optional[Path] = None,
        r3_file: Optional[Path] = None,
        do_plm: bool = True,
        do_r3: bool = False,
        archive_raw_files: bool = True,
        parent: Optional[QObject] = None,
    ) -> None:
        super().__init__(parent)
        self.bom_file = bom_file
        self.plm_file = plm_file
        self.r3_file = r3_file
        self.do_plm = do_plm
        self.do_r3 = do_r3
        self.archive_raw_files = archive_raw_files

    def run(self) -> None:
        engine = InheritanceEngine()
        results: dict[str, Any] = {}
        total_steps = (1 if self.do_plm else 0) + (1 if self.do_r3 else 0)
        current_step = 0

        self.log_signal.emit(f"[*] Bắt đầu quy trình cập nhật BOM: {self.bom_file.name}")
        self.log_signal.emit(f"[*] Đường dẫn thư mục: {self.bom_file.parent}")

        archive_dir = self.bom_file.parent / "capnhat" / "old"
        if self.archive_raw_files:
            archive_dir.mkdir(parents=True, exist_ok=True)
            self.log_signal.emit(f"[+] Thư mục lưu trữ: {archive_dir}")

        # 1. Update PLM (Button 7 VBA)
        if self.do_plm:
            current_step += 1
            self.progress_signal.emit(int((current_step / total_steps) * 50), "Đang cập nhật PLM...")
            self.log_signal.emit("=" * 60)
            self.log_signal.emit(">>> [NÚT 7 VBA] CẬP NHẬT PLM MỚI & KẾ THỪA GIẢI TRÌNH")

            if not self.plm_file or not self.plm_file.is_file():
                err_msg = "Không tìm thấy file PLM mới để cập nhật!"
                self.log_signal.emit(f"[-] LỖI: {err_msg}")
                self.finished_signal.emit(False, err_msg, results)
                return

            self.log_signal.emit(f"[*] File PLM mới: {self.plm_file.name}")
            try:
                self.log_signal.emit("[*] Đang đọc dữ liệu BOM cũ và xác định sheet lưu trữ...")
                summary: InheritanceSummary = engine.update_workbook_with_inheritance(
                    workbook_path=self.bom_file,
                    new_plm_data=self.plm_file,
                    archive_dir=archive_dir if self.archive_raw_files else None,
                    source_raw_file=self.plm_file if self.archive_raw_files else None,
                )
                results["plm"] = summary
                self.log_signal.emit(f"[+] Kế thừa thành công: {summary.inherited_count} linh kiện được giữ nguyên giải trình!")
                self.log_signal.emit(f"[+] Tổng linh kiện mới: {summary.total_new_parts}")
                self.log_signal.emit(f"[+] Bản sao lưu sheet: '{summary.archived_sheet_name}'")
                if summary.archived_files:
                    self.log_signal.emit(f"[+] Đã lưu trữ file PLM thô vào: {summary.archived_files[0].name}")
                self.log_signal.emit("[+] Pivot Table PLM đã được làm mới thành công.")
            except Exception as exc:
                logger.error("Lỗi khi cập nhật PLM: %s", exc, exc_info=True)
                err_msg = f"Lỗi cập nhật PLM: {exc}"
                self.log_signal.emit(f"[-] LỖI: {err_msg}")
                self.finished_signal.emit(False, err_msg, results)
                return

        # 2. Update R3 (Button 8 VBA)
        if self.do_r3:
            current_step += 1
            self.progress_signal.emit(int((current_step / total_steps) * 100), "Đang cập nhật SAP R3...")
            self.log_signal.emit("=" * 60)
            self.log_signal.emit(">>> [NÚT 8 VBA] CẬP NHẬT SAP R3 (CS12) & ĐỒNG BỘ PIVOT")

            if not self.r3_file or not self.r3_file.is_file():
                err_msg = "Không tìm thấy file R3 mới để cập nhật!"
                self.log_signal.emit(f"[-] LỖI: {err_msg}")
                self.finished_signal.emit(False, err_msg, results)
                return

            self.log_signal.emit(f"[*] File R3 mới: {self.r3_file.name}")
            try:
                self.log_signal.emit("[*] Đang cập nhật Sheet R3 và chuẩn hóa định dạng...")
                summary_r3: InheritanceSummary = engine.update_r3_in_workbook(
                    workbook_path=self.bom_file,
                    new_r3_data=self.r3_file,
                    archive_dir=archive_dir if self.archive_raw_files else None,
                    source_raw_file=self.r3_file if self.archive_raw_files else None,
                )
                results["r3"] = summary_r3
                self.log_signal.emit(f"[+] Cập nhật R3 thành công: {summary_r3.total_new_parts} linh kiện!")
                self.log_signal.emit(f"[+] Bản sao lưu sheet: '{summary_r3.archived_sheet_name}'")
                if summary_r3.archived_files:
                    self.log_signal.emit(f"[+] Đã lưu trữ file R3 thô vào: {summary_r3.archived_files[0].name}")
                self.log_signal.emit("[+] PivotTable3 tại cột W1 đã được tái tạo và đồng bộ cache.")
            except Exception as exc:
                logger.error("Lỗi khi cập nhật R3: %s", exc, exc_info=True)
                err_msg = f"Lỗi cập nhật SAP R3: {exc}"
                self.log_signal.emit(f"[-] LỖI: {err_msg}")
                self.finished_signal.emit(False, err_msg, results)
                return

        self.progress_signal.emit(100, "Hoàn tất cập nhật!")
        self.log_signal.emit("=" * 60)
        self.log_signal.emit("[✓] TIẾN TRÌNH CẬP NHẬT BOM ĐÃ HOÀN TẤT THÀNH CÔNG VÀ AN TOÀN!")
        self.finished_signal.emit(True, "Cập nhật hoàn tất!", results)


class BOMUpdateDialog(QDialog):
    """High-productivity dialog implementing legacy VBA Buttons 7 & 8 for BOM updates."""

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        initial_dir: Optional[Path] = None,
        initial_bom_file: Optional[Path] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Cập Nhật Dữ Liệu BOM Mới (PLM & SAP R3)")
        self._is_compact = apply_smart_dialog_geometry(
            self,
            ideal_w=880,
            ideal_h=620,
            min_w=640,
            min_h=380,
            max_w_ratio=0.94,
            max_h_ratio=0.86,
            margin_bottom=20,
        )

        self.theme_mgr = get_theme_manager()

        self.current_dir: Optional[Path] = initial_dir
        self.current_bom_file: Optional[Path] = initial_bom_file
        self.current_plm_file: Optional[Path] = None
        self.current_r3_file: Optional[Path] = None
        self._worker: Optional[BOMUpdateWorker] = None

        self._init_ui()
        self.theme_mgr.theme_changed.connect(self._apply_theme_styles)
        self._apply_theme_styles()

        if self.current_dir and self.current_dir.is_dir():
            self.txt_folder.setText(str(self.current_dir))
            self._scan_folder(self.current_dir)
        elif self.current_bom_file and self.current_bom_file.is_file():
            self.txt_folder.setText(str(self.current_bom_file.parent))
            self._scan_folder(self.current_bom_file.parent)

    def showEvent(self, event: Any) -> None:
        """Ensure dialog is cleanly clamped within available geometry before display."""
        super().showEvent(event)
        ensure_window_fits_screen(self, margin_bottom=20)

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        # 1. Header Banner
        header = QFrame(self)
        header.setStyleSheet(
            "QFrame { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E3A8A, stop:1 #2563EB); "
            "border-radius: 8px; padding: 10px 12px; color: white; }"
        )
        h_layout = QVBoxLayout(header)
        lbl_title = QLabel("🔄 CẬP NHẬT DỮ LIỆU BOM MỚI (PLM & SAP R3)", header)
        lbl_title.setStyleSheet("font-size: 14px; font-weight: bold; color: #FFFFFF;")
        h_layout.addWidget(lbl_title)

        lbl_desc = QLabel(
            "Tính năng tương ứng <b>Nút 7 ('Cập nhật PLM mới')</b> và <b>Nút 8 ('Cập nhật R3 mới')</b> từ bản VBA PE Dept.<br/>"
            "Tự động kế thừa nguyên vẹn 3 cột giải trình: <b>Cột O (Giải thích)</b>, <b>Cột P (Người phụ trách)</b>, <b>Cột Q (Quản lý check)</b> "
            "và đồng bộ lại Pivot Table.",
            header,
        )
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet("font-size: 11px; color: #E0E7FF; line-height: 1.3;")
        h_layout.addWidget(lbl_desc)
        layout.addWidget(header)

        # 2. Mode Selector: Folder Auto-scan vs Custom Files
        grp_source = QGroupBox("Chọn nguồn dữ liệu BOM cần cập nhật", self)
        src_layout = QVBoxLayout(grp_source)
        src_layout.setContentsMargins(10, 10, 10, 8)
        src_layout.setSpacing(6)

        # Folder row
        h_folder = QHBoxLayout()
        h_folder.addWidget(QLabel("Thư mục máy:", self))
        self.txt_folder = QLineEdit(self)
        self.txt_folder.setPlaceholderText("Chọn hoặc dán đường dẫn thư mục chứa BOM (VD: D:\\Sandbox\\pm_sosanhbom\\20261002 hoặc thư mục máy)")
        self.txt_folder.textChanged.connect(self._on_folder_text_changed)
        h_folder.addWidget(self.txt_folder)

        self.btn_browse_folder = QPushButton("Duyệt thư mục...", self)
        self.btn_browse_folder.setIcon(self.theme_mgr.get_styled_icon("folder"))
        self.btn_browse_folder.clicked.connect(self._browse_folder)
        h_folder.addWidget(self.btn_browse_folder)

        self.btn_rescan = QPushButton("Quét lại", self)
        self.btn_rescan.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        self.btn_rescan.clicked.connect(lambda: self._scan_folder(Path(self.txt_folder.text().strip())))
        h_folder.addWidget(self.btn_rescan)
        src_layout.addLayout(h_folder)

        # Detected files status cards
        self.status_box = QFrame(grp_source)
        self.status_box.setObjectName("status_box")
        s_layout = QVBoxLayout(self.status_box)
        s_layout.setContentsMargins(10, 6, 10, 6)
        s_layout.setSpacing(4)

        # BOM row
        h_bom = QHBoxLayout()
        self.lbl_title_bom = QLabel("<b>1. File BOM So Sánh:</b>", self.status_box)
        h_bom.addWidget(self.lbl_title_bom)
        self.lbl_detected_bom = QLabel("Chưa xác định", self.status_box)
        h_bom.addWidget(self.lbl_detected_bom, 1)
        self.btn_pick_bom = QPushButton("Chọn file...", self.status_box)
        self.btn_pick_bom.clicked.connect(self._pick_custom_bom)
        h_bom.addWidget(self.btn_pick_bom)
        s_layout.addLayout(h_bom)

        # PLM row
        h_plm = QHBoxLayout()
        self.lbl_title_plm = QLabel("<b>2. File PLM Mới (Nút 7):</b>", self.status_box)
        h_plm.addWidget(self.lbl_title_plm)
        self.lbl_detected_plm = QLabel("Chưa xác định", self.status_box)
        h_plm.addWidget(self.lbl_detected_plm, 1)
        self.btn_pick_plm = QPushButton("Chọn file...", self.status_box)
        self.btn_pick_plm.clicked.connect(self._pick_custom_plm)
        h_plm.addWidget(self.btn_pick_plm)
        s_layout.addLayout(h_plm)

        # R3 row
        h_r3 = QHBoxLayout()
        self.lbl_title_r3 = QLabel("<b>3. File SAP R3 Mới (Nút 8):</b>", self.status_box)
        h_r3.addWidget(self.lbl_title_r3)
        self.lbl_detected_r3 = QLabel("Chưa xác định", self.status_box)
        h_r3.addWidget(self.lbl_detected_r3, 1)
        self.btn_pick_r3 = QPushButton("Chọn file...", self.status_box)
        self.btn_pick_r3.clicked.connect(self._pick_custom_r3)
        h_r3.addWidget(self.btn_pick_r3)
        s_layout.addLayout(h_r3)

        src_layout.addWidget(self.status_box)

        # Options
        self.chk_archive_old = QCheckBox(
            "Tự động chuyển file PLM/R3 cũ vào thư mục 'capnhat\\old\\' (chuẩn hóa lưu trữ như VBA)", self
        )
        self.chk_archive_old.setChecked(True)
        src_layout.addWidget(self.chk_archive_old)

        layout.addWidget(grp_source)

        # 3. Action Buttons (Modern 1:1 VBA Buttons 7 & 8)
        act_box = QFrame(self)
        h_act = QHBoxLayout(act_box)
        h_act.setContentsMargins(0, 0, 0, 0)
        h_act.setSpacing(10)

        # Button 7
        self.btn_run_plm = QPushButton("🔄 7. Cập nhật PLM mới", self)
        self.btn_run_plm.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        self.btn_run_plm.setToolTip("Cập nhật file PLM Teamcenter mới vào sheet 'PLM' và kế thừa giải trình cột O, P, Q")
        self.btn_run_plm.clicked.connect(lambda: self._start_update(do_plm=True, do_r3=False))
        h_act.addWidget(self.btn_run_plm)

        # Button 8
        self.btn_run_r3 = QPushButton("🔄 8. Cập nhật R3 mới", self)
        self.btn_run_r3.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        self.btn_run_r3.setToolTip("Cập nhật file SAP R3 CS12 mới vào sheet 'R3' và làm mới Pivot Table W1")
        self.btn_run_r3.clicked.connect(lambda: self._start_update(do_plm=False, do_r3=True))
        h_act.addWidget(self.btn_run_r3)

        # Both
        self.btn_run_both = QPushButton("⚡ Cập nhật cả PLM & R3", self)
        self.btn_run_both.setIcon(self.theme_mgr.get_styled_icon("rotate-ccw"))
        self.btn_run_both.setToolTip("Cập nhật liên hoàn đồng thời cả file PLM và file SAP R3 mới")
        self.btn_run_both.clicked.connect(lambda: self._start_update(do_plm=True, do_r3=True))
        h_act.addWidget(self.btn_run_both)

        layout.addWidget(act_box)

        # 4. Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("QProgressBar { height: 16px; border-radius: 4px; text-align: center; }")
        layout.addWidget(self.progress_bar)

        # 5. Live Console
        grp_log = QGroupBox("Nhật ký thực thi chi tiết", self)
        log_layout = QVBoxLayout(grp_log)
        log_layout.setContentsMargins(8, 8, 8, 8)
        self.txt_log = QPlainTextEdit(self)
        self.txt_log.setReadOnly(True)
        self.txt_log.setFont(QFont("Consolas", 10))
        self.txt_log.setMinimumHeight(90)
        self.txt_log.setMaximumHeight(180)
        self.txt_log.setStyleSheet(
            "QPlainTextEdit { background-color: #0F172A; color: #F8FAFC; border: 1px solid #334155; border-radius: 4px; padding: 6px; }"
        )
        log_layout.addWidget(self.txt_log)
        layout.addWidget(grp_log, 1)

        # 6. Dedicated Bottom Bar (Guaranteed never clipped)
        bottom_frame = QFrame(self)
        bottom_frame.setMinimumHeight(42)
        h_bot = QHBoxLayout(bottom_frame)
        h_bot.setContentsMargins(0, 4, 0, 0)
        h_bot.setSpacing(10)

        self.btn_open_target_folder = QPushButton("📂 Mở thư mục chứa file BOM", bottom_frame)
        self.btn_open_target_folder.setIcon(self.theme_mgr.get_styled_icon("folder"))
        self.btn_open_target_folder.setMinimumHeight(32)
        self.btn_open_target_folder.clicked.connect(self._open_target_folder)
        h_bot.addWidget(self.btn_open_target_folder)

        self.btn_open_archive_folder = QPushButton("📂 Mở thư mục lưu trữ (capnhat\\old)", bottom_frame)
        self.btn_open_archive_folder.setIcon(self.theme_mgr.get_styled_icon("folder"))
        self.btn_open_archive_folder.setMinimumHeight(32)
        self.btn_open_archive_folder.clicked.connect(self._open_archive_folder)
        h_bot.addWidget(self.btn_open_archive_folder)

        h_bot.addStretch()

        self.btn_close = QPushButton("Đóng", bottom_frame)
        self.btn_close.setMinimumHeight(32)
        self.btn_close.setMinimumWidth(90)
        self.btn_close.clicked.connect(self.close)
        h_bot.addWidget(self.btn_close)

        layout.addWidget(bottom_frame)

    def _apply_theme_styles(self) -> None:
        """Apply high-contrast theme styling across all widgets, buttons, and status cards."""
        is_dark = self.theme_mgr.is_dark()
        if is_dark:
            # Dark mode styling
            self.status_box.setStyleSheet(
                "QFrame#status_box { background-color: #1E293B; border: 1px solid #334155; border-radius: 6px; padding: 6px 10px; }"
            )
            label_title_style = "color: #F8FAFC; font-weight: bold; font-size: 12px;"
            self.lbl_title_bom.setStyleSheet(label_title_style)
            self.lbl_title_plm.setStyleSheet(label_title_style)
            self.lbl_title_r3.setStyleSheet(label_title_style)

            pick_btn_style = (
                "QPushButton { background-color: #334155; color: #F1F5F9; border: 1px solid #475569; "
                "border-radius: 4px; padding: 4px 10px; font-weight: 500; font-size: 11px; } "
                "QPushButton:hover { background-color: #475569; color: #FFFFFF; }"
            )
            self.btn_pick_bom.setStyleSheet(pick_btn_style)
            self.btn_pick_plm.setStyleSheet(pick_btn_style)
            self.btn_pick_r3.setStyleSheet(pick_btn_style)

            # Action buttons styles for Dark Mode
            self._btn_plm_enabled_style = (
                "QPushButton { background-color: #2563EB; color: #FFFFFF; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #3B82F6; } "
                "QPushButton:hover { background-color: #1D4ED8; }"
            )
            self._btn_plm_disabled_style = (
                "QPushButton { background-color: #1E293B; color: #64748B; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #334155; }"
            )

            self._btn_r3_enabled_style = (
                "QPushButton { background-color: #D97706; color: #FFFFFF; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #F59E0B; } "
                "QPushButton:hover { background-color: #B45309; }"
            )
            self._btn_r3_disabled_style = (
                "QPushButton { background-color: #1E293B; color: #64748B; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #334155; }"
            )

            self._btn_both_enabled_style = (
                "QPushButton { background-color: #7C3AED; color: #FFFFFF; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #8B5CF6; } "
                "QPushButton:hover { background-color: #6D28D9; }"
            )
            self._btn_both_disabled_style = (
                "QPushButton { background-color: #1E293B; color: #64748B; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #334155; }"
            )

            # Bottom buttons
            bot_btn_style = (
                "QPushButton { background-color: #1E293B; color: #F1F5F9; border: 1px solid #334155; "
                "border-radius: 4px; padding: 6px 14px; font-weight: 500; font-size: 12px; } "
                "QPushButton:hover { background-color: #334155; }"
            )
            self.btn_open_target_folder.setStyleSheet(bot_btn_style)
            self.btn_open_archive_folder.setStyleSheet(bot_btn_style)
            self.btn_close.setStyleSheet(
                "QPushButton { background-color: #334155; color: #FFFFFF; border: 1px solid #475569; "
                "border-radius: 4px; padding: 6px 20px; font-weight: bold; font-size: 12px; } "
                "QPushButton:hover { background-color: #475569; }"
            )
        else:
            # Light mode styling
            self.status_box.setStyleSheet(
                "QFrame#status_box { background-color: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 6px; padding: 6px 10px; }"
            )
            label_title_style = "color: #0F172A; font-weight: bold; font-size: 12px;"
            self.lbl_title_bom.setStyleSheet(label_title_style)
            self.lbl_title_plm.setStyleSheet(label_title_style)
            self.lbl_title_r3.setStyleSheet(label_title_style)

            pick_btn_style = (
                "QPushButton { background-color: #E2E8F0; color: #0F172A; border: 1px solid #CBD5E1; "
                "border-radius: 4px; padding: 4px 10px; font-weight: 500; font-size: 11px; } "
                "QPushButton:hover { background-color: #CBD5E1; }"
            )
            self.btn_pick_bom.setStyleSheet(pick_btn_style)
            self.btn_pick_plm.setStyleSheet(pick_btn_style)
            self.btn_pick_r3.setStyleSheet(pick_btn_style)

            # Action buttons styles for Light Mode
            self._btn_plm_enabled_style = (
                "QPushButton { background-color: #2563EB; color: #FFFFFF; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; } "
                "QPushButton:hover { background-color: #1D4ED8; }"
            )
            self._btn_plm_disabled_style = (
                "QPushButton { background-color: #E2E8F0; color: #94A3B8; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #CBD5E1; }"
            )

            self._btn_r3_enabled_style = (
                "QPushButton { background-color: #D97706; color: #FFFFFF; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; } "
                "QPushButton:hover { background-color: #B45309; }"
            )
            self._btn_r3_disabled_style = (
                "QPushButton { background-color: #E2E8F0; color: #94A3B8; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #CBD5E1; }"
            )

            self._btn_both_enabled_style = (
                "QPushButton { background-color: #7C3AED; color: #FFFFFF; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; } "
                "QPushButton:hover { background-color: #6D28D9; }"
            )
            self._btn_both_disabled_style = (
                "QPushButton { background-color: #E2E8F0; color: #94A3B8; font-weight: bold; font-size: 12px; "
                "padding: 8px 14px; border-radius: 6px; border: 1px solid #CBD5E1; }"
            )

            # Bottom buttons
            bot_btn_style = (
                "QPushButton { background-color: #F1F5F9; color: #0F172A; border: 1px solid #CBD5E1; "
                "border-radius: 4px; padding: 6px 14px; font-weight: 500; font-size: 12px; } "
                "QPushButton:hover { background-color: #E2E8F0; }"
            )
            self.btn_open_target_folder.setStyleSheet(bot_btn_style)
            self.btn_open_archive_folder.setStyleSheet(bot_btn_style)
            self.btn_close.setStyleSheet(
                "QPushButton { background-color: #0F172A; color: #FFFFFF; border: 1px solid #334155; "
                "border-radius: 4px; padding: 6px 20px; font-weight: bold; font-size: 12px; } "
                "QPushButton:hover { background-color: #1E293B; }"
            )

        self._update_button_states()
        self._update_status_labels_styling()

    def _update_status_labels_styling(self) -> None:
        """Apply theme-appropriate contrast colors to file detection status labels."""
        is_dark = self.theme_mgr.is_dark()
        # BOM status
        if self.current_bom_file and self.current_bom_file.is_file():
            self.lbl_detected_bom.setStyleSheet(
                "color: #4ADE80; font-weight: bold; font-size: 12px;" if is_dark else "color: #16A34A; font-weight: bold; font-size: 12px;"
            )
        else:
            self.lbl_detected_bom.setStyleSheet(
                "color: #F87171; font-weight: bold; font-style: italic; font-size: 12px;" if is_dark else "color: #DC2626; font-weight: bold; font-style: italic; font-size: 12px;"
            )

        # PLM status
        if self.current_plm_file and self.current_plm_file.is_file():
            self.lbl_detected_plm.setStyleSheet(
                "color: #60A5FA; font-weight: bold; font-size: 12px;" if is_dark else "color: #2563EB; font-weight: bold; font-size: 12px;"
            )
        else:
            self.lbl_detected_plm.setStyleSheet(
                "color: #94A3B8; font-style: italic; font-size: 12px;" if is_dark else "color: #64748B; font-style: italic; font-size: 12px;"
            )

        # R3 status
        if self.current_r3_file and self.current_r3_file.is_file():
            self.lbl_detected_r3.setStyleSheet(
                "color: #FBBF24; font-weight: bold; font-size: 12px;" if is_dark else "color: #D97706; font-weight: bold; font-size: 12px;"
            )
        else:
            self.lbl_detected_r3.setStyleSheet(
                "color: #94A3B8; font-style: italic; font-size: 12px;" if is_dark else "color: #64748B; font-style: italic; font-size: 12px;"
            )

    # =========================================================================
    # Folder Scanning & File Auto-Discovery
    # =========================================================================

    def _browse_folder(self) -> None:
        start_dir = self.txt_folder.text().strip() or str(Path.cwd())
        selected = QFileDialog.getExistingDirectory(self, "Chọn thư mục chứa BOM và file cập nhật", start_dir)
        if selected:
            self.txt_folder.setText(selected)
            self._scan_folder(Path(selected))

    def _on_folder_text_changed(self, text: str) -> None:
        clean = text.strip().strip('"').strip("'")
        p = Path(clean)
        if p.is_dir():
            self._scan_folder(p)

    def _scan_folder(self, folder: Path) -> None:
        if not folder or not folder.is_dir():
            return
        self.current_dir = folder
        self.txt_log.appendPlainText(f"[*] Đang quét thư mục: {folder}")

        # 1. Search for BOM file: BOM*.xlsx / xlsm
        bom_candidates = [
            f for f in folder.iterdir()
            if f.is_file() and f.name.upper().startswith("BOM") and f.suffix.lower() in (".xlsx", ".xlsm")
            and not f.name.startswith("~$")
        ]
        if bom_candidates:
            self.current_bom_file = bom_candidates[0]
            self.lbl_detected_bom.setText(f"✓ {self.current_bom_file.name} ({self.current_bom_file.stat().st_size // 1024} KB)")
            self.txt_log.appendPlainText(f"[+] Tìm thấy file BOM: {self.current_bom_file.name}")
        else:
            self.current_bom_file = None
            self.lbl_detected_bom.setText("Không tìm thấy file BOM* trong thư mục")
            self.txt_log.appendPlainText("[-] Không tìm thấy file BOM* trong thư mục.")

        # 2. Search for PLM file: PLM*.xlsx / xlsm / xls
        plm_candidates = [
            f for f in folder.iterdir()
            if f.is_file() and f.name.upper().startswith("PLM") and f.suffix.lower() in (".xlsx", ".xlsm", ".xls")
            and not f.name.startswith("~$")
        ]
        if plm_candidates:
            self.current_plm_file = plm_candidates[0]
            self.lbl_detected_plm.setText(f"✓ {self.current_plm_file.name} ({self.current_plm_file.stat().st_size // 1024} KB)")
            self.txt_log.appendPlainText(f"[+] Tìm thấy file PLM mới: {self.current_plm_file.name}")
        else:
            self.current_plm_file = None
            self.lbl_detected_plm.setText("Chưa tìm thấy file PLM* (sẽ bỏ qua Nút 7 nếu không cập nhật)")

        # 3. Search for R3 file: R3*.xlsx / xlsm / xls
        r3_candidates = [
            f for f in folder.iterdir()
            if f.is_file() and f.name.upper().startswith("R3") and f.suffix.lower() in (".xlsx", ".xlsm", ".xls")
            and not f.name.startswith("~$")
        ]
        if r3_candidates:
            self.current_r3_file = r3_candidates[0]
            self.lbl_detected_r3.setText(f"✓ {self.current_r3_file.name} ({self.current_r3_file.stat().st_size // 1024} KB)")
            self.txt_log.appendPlainText(f"[+] Tìm thấy file R3 mới: {self.current_r3_file.name}")
        else:
            self.current_r3_file = None
            self.lbl_detected_r3.setText("Chưa tìm thấy file R3* (sẽ bỏ qua Nút 8 nếu không cập nhật)")

        self._update_status_labels_styling()
        self._update_button_states()

    def _pick_custom_bom(self) -> None:
        start = str(self.current_dir) if self.current_dir else str(Path.cwd())
        f, _ = QFileDialog.getOpenFileName(self, "Chọn file BOM so sánh", start, "Excel Files (*.xlsx *.xlsm)")
        if f:
            p = Path(f)
            self.current_bom_file = p
            self.lbl_detected_bom.setText(f"✓ {p.name}")
            if not self.current_dir:
                self.current_dir = p.parent
                self.txt_folder.setText(str(p.parent))
            self.txt_log.appendPlainText(f"[+] Đã chọn file BOM: {self.current_bom_file.name}")
            self._update_status_labels_styling()
            self._update_button_states()

    def _pick_custom_plm(self) -> None:
        start = str(self.current_dir) if self.current_dir else str(Path.cwd())
        f, _ = QFileDialog.getOpenFileName(self, "Chọn file PLM mới", start, "Excel Files (*.xlsx *.xlsm *.xls)")
        if f:
            p = Path(f)
            self.current_plm_file = p
            self.lbl_detected_plm.setText(f"✓ {p.name}")
            self.txt_log.appendPlainText(f"[+] Đã chọn file PLM: {self.current_plm_file.name}")
            self._update_status_labels_styling()
            self._update_button_states()

    def _pick_custom_r3(self) -> None:
        start = str(self.current_dir) if self.current_dir else str(Path.cwd())
        f, _ = QFileDialog.getOpenFileName(self, "Chọn file SAP R3 mới", start, "Excel Files (*.xlsx *.xlsm *.xls)")
        if f:
            p = Path(f)
            self.current_r3_file = p
            self.lbl_detected_r3.setText(f"✓ {p.name}")
            self.txt_log.appendPlainText(f"[+] Đã chọn file R3: {self.current_r3_file.name}")
            self._update_status_labels_styling()
            self._update_button_states()

    def _update_button_states(self) -> None:
        has_bom = self.current_bom_file is not None and self.current_bom_file.is_file()
        has_plm = self.current_plm_file is not None and self.current_plm_file.is_file()
        has_r3 = self.current_r3_file is not None and self.current_r3_file.is_file()

        plm_enabled = has_bom and has_plm
        r3_enabled = has_bom and has_r3
        both_enabled = has_bom and (has_plm or has_r3)

        self.btn_run_plm.setEnabled(plm_enabled)
        self.btn_run_plm.setStyleSheet(
            getattr(self, "_btn_plm_enabled_style", "") if plm_enabled else getattr(self, "_btn_plm_disabled_style", "")
        )

        self.btn_run_r3.setEnabled(r3_enabled)
        self.btn_run_r3.setStyleSheet(
            getattr(self, "_btn_r3_enabled_style", "") if r3_enabled else getattr(self, "_btn_r3_disabled_style", "")
        )

        self.btn_run_both.setEnabled(both_enabled)
        self.btn_run_both.setStyleSheet(
            getattr(self, "_btn_both_enabled_style", "") if both_enabled else getattr(self, "_btn_both_disabled_style", "")
        )

    # =========================================================================
    # Execution
    # =========================================================================

    def _start_update(self, do_plm: bool, do_r3: bool) -> None:
        if not self.current_bom_file or not self.current_bom_file.is_file():
            QMessageBox.warning(self, "Chưa chọn file BOM", "Vui lòng chọn file BOM so sánh trước khi thực hiện!")
            return

        if do_plm and (not self.current_plm_file or not self.current_plm_file.is_file()):
            QMessageBox.warning(self, "Chưa có file PLM", "Không tìm thấy file PLM mới để cập nhật!")
            return

        if do_r3 and (not self.current_r3_file or not self.current_r3_file.is_file()):
            QMessageBox.warning(self, "Chưa có file R3", "Không tìm thấy file SAP R3 mới để cập nhật!")
            return

        # Disable buttons during work
        self.btn_run_plm.setEnabled(False)
        self.btn_run_r3.setEnabled(False)
        self.btn_run_both.setEnabled(False)
        self.progress_bar.setValue(10)

        self._worker = BOMUpdateWorker(
            bom_file=self.current_bom_file,
            plm_file=self.current_plm_file if do_plm else None,
            r3_file=self.current_r3_file if do_r3 else None,
            do_plm=do_plm,
            do_r3=do_r3,
            archive_raw_files=self.chk_archive_old.isChecked(),
            parent=self,
        )
        self._worker.progress_signal.connect(self._on_progress)
        self._worker.log_signal.connect(self._on_log)
        self._worker.finished_signal.connect(self._on_finished)
        self._worker.start()

    @pyqtSlot(int, str)
    def _on_progress(self, val: int, msg: str) -> None:
        self.progress_bar.setValue(val)
        self.progress_bar.setFormat(f"%p% - {msg}")

    @pyqtSlot(str)
    def _on_log(self, text: str) -> None:
        self.txt_log.appendPlainText(text)

    @pyqtSlot(bool, str, dict)
    def _on_finished(self, success: bool, message: str, results: dict) -> None:
        self._update_button_states()
        if success:
            detail_lines = []
            if "plm" in results:
                p_sum: InheritanceSummary = results["plm"]
                detail_lines.append(f"• PLM: {p_sum.inherited_count} linh kiện kế thừa giải trình, lưu sheet '{p_sum.archived_sheet_name}'")
            if "r3" in results:
                r_sum: InheritanceSummary = results["r3"]
                detail_lines.append(f"• SAP R3: {r_sum.total_new_parts} linh kiện, đồng bộ PivotTable3")

            QMessageBox.information(
                self,
                "Cập Nhật BOM Hoàn Tất",
                f"Đã cập nhật dữ liệu mới vào BOM thành công!\n\n"
                f"File BOM: {self.current_bom_file.name}\n"
                + "\n".join(detail_lines)
                + "\n\nFile nguồn thô đã được chuyển vào thư mục 'capnhat\\old\\'."
            )
            # Re-scan folder to update available files
            if self.current_dir:
                self._scan_folder(self.current_dir)
        else:
            QMessageBox.critical(self, "Lỗi Cập Nhật BOM", f"Tiến trình cập nhật BOM thất bại:\n\n{message}")

    # =========================================================================
    # Helpers
    # =========================================================================

    def _open_target_folder(self) -> None:
        folder = self.current_dir or (self.current_bom_file.parent if self.current_bom_file else None)
        if folder and folder.is_dir():
            try:
                os.startfile(str(folder))
            except Exception as e:
                QMessageBox.warning(self, "Không thể mở thư mục", str(e))
        else:
            QMessageBox.information(self, "Chưa có thư mục", "Chưa xác định được thư mục mục tiêu.")

    def _open_archive_folder(self) -> None:
        folder = self.current_dir or (self.current_bom_file.parent if self.current_bom_file else None)
        if folder:
            arch = folder / "capnhat" / "old"
            if arch.is_dir():
                try:
                    os.startfile(str(arch))
                except Exception as e:
                    QMessageBox.warning(self, "Không thể mở thư mục", str(e))
            else:
                QMessageBox.information(self, "Chưa tạo thư mục lưu trữ", f"Thư mục lưu trữ chưa được tạo:\n{arch}")
