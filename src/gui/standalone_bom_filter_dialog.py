"""Standalone BOM Filtering Dialog (Áp Dụng Bộ Lọc BOM Cho Tệp Có Sẵn).

Allows engineers and leaders to select existing PLM Excel files (.xlsx, .xlsm, .xls)
or entire folders on disk and re-run model-specific decomposition pruning (ModelPruner/BolocBom)
without needing to re-connect or re-download from Teamcenter TC24.

Includes:
- Business notice banner clarifying:
  * Safe additional pruning (xóa thêm / lọc sót) on previously filtered files.
  * Child component loss warning and requirement of backup/Teamcenter for full tree restore.
- Non-blocking execution via QThread (StandalonePruneWorker) with progress bar and live log.
- Intelligent model auto-detection with manual override from DB/registry.
- Automatic safety backup (.bak / backup_before_filter/) with custom location support.
- Dual-theme aware UI (Kyocera Light / Dark mode).
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any, Optional

from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.bom_filter_manager import BOMFilterManager, get_bom_filter_manager
from src.core.model_pruner import ModelPruner, auto_detect_model_from_path
from src.gui.styles import (
    apply_smart_dialog_geometry,
    ensure_window_fits_screen,
    get_theme_manager,
)

logger = logging.getLogger(__name__)


class StandalonePruneWorker(QThread):
    """Background worker thread executing BOM pruning across multiple Excel files."""

    progress_signal = pyqtSignal(int, int, str)  # current, total, filename
    log_signal = pyqtSignal(str)                 # log message line
    file_finished_signal = pyqtSignal(int, dict) # row_index, result_dict
    all_finished_signal = pyqtSignal(dict)       # summary stats

    def __init__(
        self,
        files: list[tuple[int, Path, str]], # list of (row_idx, path, effective_model)
        filter_manager: Optional[BOMFilterManager] = None,
        create_backup: bool = True,
        backup_dir: Optional[Path] = None,
        output_as_new: bool = False,
    ) -> None:
        super().__init__()
        self.files = files
        self.filter_manager = filter_manager
        self.create_backup = create_backup
        self.backup_dir = backup_dir
        self.output_as_new = output_as_new
        self._is_cancelled = False

    def cancel(self) -> None:
        """Signal worker to stop processing remaining files."""
        self._is_cancelled = True

    def run(self) -> None:
        pruner = ModelPruner(filter_manager=self.filter_manager)
        total = len(self.files)
        success_count = 0
        error_count = 0
        total_deleted_rows = 0

        self.log_signal.emit(f"🚀 Bắt đầu quá trình lọc cho {total} tệp PLM...")

        for idx, (row_idx, file_path, target_model) in enumerate(self.files, start=1):
            if self._is_cancelled:
                self.log_signal.emit("⏹ Quá trình đã bị người dùng hủy bỏ.")
                break

            self.progress_signal.emit(idx, total, file_path.name)
            self.log_signal.emit(f"[{idx}/{total}] Đang xử lý: {file_path.name} (Model: {target_model})...")

            out_path: Optional[Path] = None
            if self.output_as_new:
                out_path = file_path.with_name(f"{file_path.stem}_filtered{file_path.suffix}")

            try:
                res = pruner.prune_file(
                    input_path=file_path,
                    output_path=out_path,
                    model_name=target_model,
                    create_backup=self.create_backup,
                    backup_dir=self.backup_dir,
                )

                if res.get("backup_path"):
                    self.log_signal.emit(f"   [+] Bản sao lưu: {Path(res['backup_path']).name}")

                initial_r = res.get("initial_rows", 0)
                final_r = res.get("final_rows", 0)
                del_r = res.get("deleted_rows", 0)
                rules_c = res.get("rules_count", 0)
                total_deleted_rows += del_r
                success_count += 1

                self.log_signal.emit(
                    f"   [✓] Hoàn tất ({rules_c} quy tắc): Ban đầu {initial_r:,} dòng -> "
                    f"Còn lại {final_r:,} dòng (Đã cắt {del_r:,} dòng linh kiện con)."
                )
                self.file_finished_signal.emit(row_idx, res)

            except Exception as exc:
                error_count += 1
                logger.error("Lỗi khi lọc tệp %s: %s", file_path, exc)
                self.log_signal.emit(f"   [-] LỖI khi xử lý {file_path.name}: {exc}")
                err_res = {
                    "input_path": str(file_path),
                    "output_path": str(out_path or file_path),
                    "backup_path": None,
                    "model_name": target_model,
                    "initial_rows": 0,
                    "final_rows": 0,
                    "deleted_rows": 0,
                    "rules_count": 0,
                    "success": False,
                    "error": str(exc),
                }
                self.file_finished_signal.emit(row_idx, err_res)

        summary = {
            "total": total,
            "processed": idx if not self._is_cancelled else idx - 1,
            "success": success_count,
            "error": error_count,
            "total_deleted_rows": total_deleted_rows,
            "cancelled": self._is_cancelled,
        }
        self.log_signal.emit(
            f"🎉 Kết thúc: Đã xử lý {summary['processed']}/{total} tệp "
            f"({success_count} thành công, {error_count} lỗi). "
            f"Tổng số dòng đã cắt tỉa: {total_deleted_rows:,} dòng."
        )
        self.all_finished_signal.emit(summary)


class StandaloneBOMFilterDialog(QDialog):
    """Enterprise Dialog to browse and prune existing PLM files without Teamcenter re-download."""

    def __init__(
        self,
        initial_model: Optional[str] = None,
        initial_path: Optional[Path | str] = None,
        filter_manager: Optional[BOMFilterManager] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("⚡ Áp Dụng Bộ Lọc BOM Cho Tệp Có Sẵn (Standalone BOM Pruner)")
        self._is_compact = apply_smart_dialog_geometry(
            self,
            ideal_w=980,
            ideal_h=650,
            min_w=680,
            min_h=380,
            max_w_ratio=0.94,
            max_h_ratio=0.86,
            margin_bottom=20,
        )

        self.filter_manager = filter_manager or get_bom_filter_manager()
        self.pruner = ModelPruner(filter_manager=self.filter_manager)
        self.initial_model = initial_model or ""
        self.initial_path = Path(initial_path) if initial_path else None

        self.detected_files: list[Path] = []
        self.worker: Optional[StandalonePruneWorker] = None

        self._init_ui()
        self._apply_theme()
        self._populate_models()

        if self.initial_path:
            if self.initial_path.is_file():
                self.rb_mode_files.setChecked(True)
                self.txt_source_path.setText(str(self.initial_path))
                self._load_selected_files([self.initial_path])
            elif self.initial_path.is_dir():
                self.rb_mode_folder.setChecked(True)
                self.txt_source_path.setText(str(self.initial_path))
                self._scan_folder_files()

    def showEvent(self, event: Any) -> None:
        """Ensure dialog fits completely on screen without taskbar occlusion."""
        super().showEvent(event)
        ensure_window_fits_screen(self, margin_bottom=20)

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 10, 12, 10)
        main_layout.setSpacing(8)

        # 1. Header Banner
        header_frame = QFrame()
        header_frame.setObjectName("dialog_header_frame")
        h_layout = QHBoxLayout(header_frame)
        h_layout.setContentsMargins(12, 8, 12, 8)
        h_layout.setSpacing(10)

        icon_lbl = QLabel()
        icon_lbl.setPixmap(get_theme_manager().get_styled_icon("filter").pixmap(28, 28))
        h_layout.addWidget(icon_lbl)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        lbl_title = QLabel("Cắt Tỉa BOM PLM Trực Tiếp Từ Tệp Có Sẵn")
        lbl_title.setFont(QFont("Calibri", 12, QFont.Weight.Bold))
        title_vbox.addWidget(lbl_title)

        lbl_sub = QLabel(
            "Chạy lại bộ quy tắc cắt tỉa (BolocBom) cho file Excel PLM trên ổ đĩa. "
            "Tiết kiệm thời gian mở và đợi Teamcenter."
        )
        lbl_sub.setFont(QFont("Calibri", 9))
        title_vbox.addWidget(lbl_sub)
        h_layout.addLayout(title_vbox)
        h_layout.addStretch()

        main_layout.addWidget(header_frame, 0)

        # 2. Critical Business Notice Card
        notice_box = QFrame()
        notice_box.setObjectName("business_notice_box")
        nb_layout = QHBoxLayout(notice_box)
        nb_layout.setContentsMargins(12, 8, 12, 8)
        nb_layout.setSpacing(10)

        warn_icon = QLabel("⚠️")
        warn_icon.setFont(QFont("Segoe UI Emoji", 16))
        nb_layout.addWidget(warn_icon)

        notice_text = QLabel(
            "<b>Lưu ý nghiệp vụ quan trọng khi lọc lại trên file có sẵn:</b><br/>"
            "• <b>Xóa thêm (lọc sót):</b> Nếu chi tiết cần xóa bổ sung vẫn còn trong file Excel, "
            "hệ thống sẽ quét và loại bỏ chính xác.<br/>"
            "• <b>Khôi phục linh kiện con:</b> File sau khi lọc đã bị xóa vĩnh viễn các dòng con. "
            "Để phục hồi cấu trúc cây gốc, hệ thống cần bản sao lưu (backup) hoặc tải lại từ Teamcenter.<br/>"
            "• <b>Cơ chế an toàn:</b> Tùy chọn sao lưu tự động (.bak) được bật mặc định bên dưới để bảo vệ dữ liệu."
        )
        notice_text.setFont(QFont("Calibri", 9))
        notice_text.setWordWrap(True)
        nb_layout.addWidget(notice_text, 1)

        main_layout.addWidget(notice_box, 0)

        # 3. Source Selection & Options (Horizontal Splitter or Groups)
        src_group = QGroupBox("1. Chọn Nguồn Tệp PLM")
        src_layout = QVBoxLayout(src_group)
        src_layout.setContentsMargins(8, 8, 8, 8)
        src_layout.setSpacing(6)

        # Mode selector radio buttons
        mode_row = QHBoxLayout()
        self.rb_mode_files = QRadioButton("Chọn một hoặc nhiều tệp Excel (.xlsx, .xlsm)")
        self.rb_mode_folder = QRadioButton("Quét toàn bộ thư mục chứa các tệp PLM")
        self.rb_mode_folder.setChecked(True)
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.rb_mode_files)
        self.mode_group.addButton(self.rb_mode_folder)
        self.rb_mode_files.toggled.connect(self._on_mode_changed)

        mode_row.addWidget(self.rb_mode_folder)
        mode_row.addWidget(self.rb_mode_files)
        mode_row.addStretch()
        src_layout.addLayout(mode_row)

        # Path browse row
        path_row = QHBoxLayout()
        self.txt_source_path = QLineEdit()
        self.txt_source_path.setPlaceholderText("Đường dẫn thư mục hoặc tệp...")
        path_row.addWidget(self.txt_source_path, 1)

        self.btn_browse = QPushButton("📁 Duyệt Tìm...")
        self.btn_browse.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_browse.clicked.connect(self._on_browse)
        path_row.addWidget(self.btn_browse)

        self.btn_rescan = QPushButton("🔄 Quét lại")
        self.btn_rescan.setFont(QFont("Calibri", 9))
        self.btn_rescan.clicked.connect(self._on_rescan)
        path_row.addWidget(self.btn_rescan)

        src_layout.addLayout(path_row)

        # Scan filter options
        opts_row = QHBoxLayout()
        self.chk_recursive = QCheckBox("Quét cả thư mục con (Recursive)")
        self.chk_recursive.setChecked(True)
        self.chk_recursive.stateChanged.connect(self._on_rescan)
        opts_row.addWidget(self.chk_recursive)

        self.chk_prefix_plm_only = QCheckBox("Chỉ nhận diện tệp có tiền tố 'PLM_' (ví dụ: PLM_*.xlsx)")
        self.chk_prefix_plm_only.setChecked(True)
        self.chk_prefix_plm_only.stateChanged.connect(self._on_rescan)
        opts_row.addWidget(self.chk_prefix_plm_only)

        opts_row.addStretch()
        src_layout.addLayout(opts_row)

        main_layout.addWidget(src_group, 0)

        # 4. Target Model & Backup Options
        cfg_group = QGroupBox("2. Cấu Hình Dòng Máy & Tùy Chọn Sao Lưu")
        cfg_layout = QHBoxLayout(cfg_group)
        cfg_layout.setContentsMargins(8, 8, 8, 8)
        cfg_layout.setSpacing(12)

        # Left sub-column: Model selection
        model_vbox = QVBoxLayout()
        m_row = QHBoxLayout()
        m_row.addWidget(QLabel("Dòng máy (Model):"))
        self.combo_model = QComboBox()
        self.combo_model.currentIndexChanged.connect(self._on_model_changed)
        m_row.addWidget(self.combo_model, 1)
        model_vbox.addLayout(m_row)

        self.chk_autodetect = QCheckBox("🔍 Tự động nhận diện dòng máy theo tên tệp / thư mục")
        self.chk_autodetect.setChecked(True)
        self.chk_autodetect.setToolTip("Ưu tiên nhận diện model từ tên file/thư mục; nếu không tìm thấy sẽ dùng model đã chọn ở trên")
        self.chk_autodetect.stateChanged.connect(self._refresh_detected_table_models)
        model_vbox.addWidget(self.chk_autodetect)

        self.lbl_rules_info = QLabel("Quy tắc: Đang tải...")
        self.lbl_rules_info.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.lbl_rules_info.setStyleSheet("color: #2563EB;")
        model_vbox.addWidget(self.lbl_rules_info)
        cfg_layout.addLayout(model_vbox, 1)

        # Right sub-column: Backup & Overwrite options
        backup_vbox = QVBoxLayout()
        self.chk_backup = QCheckBox("🛡️ Tự động tạo bản sao lưu (.bak) trước khi cắt tỉa (Khuyên dùng)")
        self.chk_backup.setChecked(True)
        backup_vbox.addWidget(self.chk_backup)

        out_mode_row = QHBoxLayout()
        self.rb_overwrite = QRadioButton("Ghi đè trực tiếp lên tệp gốc (Đã có backup)")
        self.rb_overwrite.setChecked(True)
        self.rb_save_new = QRadioButton("Lưu thành tệp mới (*_filtered.xlsx)")
        out_mode_row.addWidget(self.rb_overwrite)
        out_mode_row.addWidget(self.rb_save_new)
        backup_vbox.addLayout(out_mode_row)

        cfg_layout.addLayout(backup_vbox, 1)
        main_layout.addWidget(cfg_group, 0)

        # 5. Middle Splitter: Files Table + Execution Log
        mid_splitter = QSplitter(Qt.Orientation.Vertical)

        # Upper: Table of files
        table_container = QWidget()
        t_layout = QVBoxLayout(table_container)
        t_layout.setContentsMargins(0, 0, 0, 0)
        t_layout.setSpacing(4)

        t_header_row = QHBoxLayout()
        self.lbl_file_count = QLabel("Danh sách tệp phát hiện: 0 tệp")
        self.lbl_file_count.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        t_header_row.addWidget(self.lbl_file_count)
        t_header_row.addStretch()
        t_layout.addLayout(t_header_row)

        self.table_files = QTableWidget()
        self.table_files.setColumnCount(6)
        self.table_files.setHorizontalHeaderLabels([
            "STT", "Tên Tệp PLM", "Dòng Máy Áp Dụng", "Kích Thước", "Đường Dẫn", "Trạng Thái"
        ])
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table_files.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table_files.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_files.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table_files.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table_files.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.table_files.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_files.setAlternatingRowColors(True)
        t_layout.addWidget(self.table_files)
        mid_splitter.addWidget(table_container)

        # Lower: Real-time execution log
        log_container = QWidget()
        l_layout = QVBoxLayout(log_container)
        l_layout.setContentsMargins(0, 0, 0, 0)
        l_layout.setSpacing(4)

        lbl_log = QLabel("Nhật ký thực thi thời gian thực:")
        lbl_log.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        l_layout.addWidget(lbl_log)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 9))
        l_layout.addWidget(self.log_view)
        mid_splitter.addWidget(log_container)

        mid_splitter.setStretchFactor(0, 3)
        mid_splitter.setStretchFactor(1, 2)
        mid_splitter.setSizes([260, 160])
        main_layout.addWidget(mid_splitter, 1)

        # 6. Bottom Progress & Action Bar
        bottom_frame = QFrame()
        bottom_frame.setObjectName("dialog_bottom_frame")
        b_layout = QVBoxLayout(bottom_frame)
        b_layout.setContentsMargins(8, 6, 8, 6)
        b_layout.setSpacing(6)

        prog_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        prog_row.addWidget(self.progress_bar, 1)

        self.lbl_progress_status = QLabel("Sẵn sàng")
        self.lbl_progress_status.setFont(QFont("Calibri", 9))
        prog_row.addWidget(self.lbl_progress_status)
        b_layout.addLayout(prog_row)

        btn_row = QHBoxLayout()
        self.btn_open_folder = QPushButton("📂 Mở Thư Mục Chứa Tệp")
        self.btn_open_folder.setFont(QFont("Calibri", 9))
        self.btn_open_folder.clicked.connect(self._on_open_folder)
        btn_row.addWidget(self.btn_open_folder)

        btn_row.addStretch()

        self.btn_start = QPushButton("🚀 Bắt Đầu Lọc BOM")
        self.btn_start.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_start.setStyleSheet(
            "background-color: #2563EB; color: white; border-radius: 4px; padding: 6px 20px;"
        )
        self.btn_start.clicked.connect(self._on_start_pruning)
        btn_row.addWidget(self.btn_start)

        self.btn_stop = QPushButton("⏹ Dừng")
        self.btn_stop.setFont(QFont("Calibri", 9))
        self.btn_stop.setStyleSheet(
            "background-color: #DC2626; color: white; border-radius: 4px; padding: 6px 14px;"
        )
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._on_stop_pruning)
        btn_row.addWidget(self.btn_stop)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setFont(QFont("Calibri", 9))
        self.btn_close.clicked.connect(self.accept)
        btn_row.addWidget(self.btn_close)

        b_layout.addLayout(btn_row)
        main_layout.addWidget(bottom_frame, 0)

    def _apply_theme(self) -> None:
        """Apply dark or light styling consistent with Kyocera theme tokens."""
        is_dark = get_theme_manager().is_dark()
        bg_col = "#0B0F17" if is_dark else "#F8FAFC"
        fg_col = "#F1F5F9" if is_dark else "#0F172A"
        card_bg = "#111722" if is_dark else "#FFFFFF"
        border_col = "#2A374A" if is_dark else "#E2E8F0"

        self.setStyleSheet(f"""
            StandaloneBOMFilterDialog, QDialog {{ background-color: {bg_col}; color: {fg_col}; }}
            QFrame#dialog_header_frame, QFrame#dialog_bottom_frame {{
                background-color: {'#151D2A' if is_dark else '#F1F5F9'};
                border: 1px solid {border_col};
                border-radius: 6px;
            }}
            QFrame#business_notice_box {{
                background-color: {'#2A2415' if is_dark else '#FEF3C7'};
                border: 1px solid {'#D97706' if is_dark else '#F59E0B'};
                border-radius: 6px;
            }}
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {border_col};
                border-radius: 6px;
                margin-top: 6px;
                padding-top: 10px;
                color: {fg_col};
                background-color: {card_bg};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
                color: {'#38BDF8' if is_dark else '#0284C7'};
            }}
            QTableWidget {{
                background-color: {'#0E141F' if is_dark else '#FFFFFF'};
                color: {fg_col};
                border: 1px solid {border_col};
                gridline-color: {border_col};
            }}
            QPlainTextEdit {{
                background-color: {'#05080E' if is_dark else '#F8FAFC'};
                color: {'#A7F3D0' if is_dark else '#065F46'};
                border: 1px solid {border_col};
                border-radius: 4px;
            }}
        """)

    def _populate_models(self) -> None:
        """Populate model combobox from SQLite database and default rules."""
        self.combo_model.clear()
        models: list[str] = []
        try:
            models = self.filter_manager.get_all_models()
        except Exception:
            pass

        from src.core.model_pruner import COMMON_MODELS
        if not models:
            models = list(COMMON_MODELS)
        else:
            for cm in COMMON_MODELS:
                if cm not in models:
                    models.append(cm)

        if self.initial_model and self.initial_model not in models:
            models.append(self.initial_model)

        models = sorted(list(set(models)), key=lambda s: s.lower())
        self.combo_model.addItems(models)

        # Select initial model if specified
        if self.initial_model:
            clean_init = re.sub(r"\s+", "", self.initial_model.lower())
            for idx in range(self.combo_model.count()):
                text = re.sub(r"\s+", "", self.combo_model.itemText(idx).lower())
                if text == clean_init or clean_init in text or text in clean_init:
                    self.combo_model.setCurrentIndex(idx)
                    break

        self._on_model_changed()

    def _on_model_changed(self) -> None:
        """Update rule count indicator for selected model."""
        selected_model = self.combo_model.currentText().strip()
        if not selected_model:
            self.lbl_rules_info.setText("Quy tắc: Chưa chọn dòng máy")
            return

        rules = self.pruner.get_rules_for_model(selected_model)
        self.lbl_rules_info.setText(f"Bộ quy tắc: Có {len(rules)} quy tắc đang áp dụng cho [{selected_model}]")
        self._refresh_detected_table_models()

    def _on_mode_changed(self) -> None:
        """Handle toggling between file selection and folder selection."""
        is_folder = self.rb_mode_folder.isChecked()
        self.chk_recursive.setEnabled(is_folder)
        self.txt_source_path.clear()
        self.detected_files.clear()
        self._populate_table()

    def _on_browse(self) -> None:
        """Browse file or folder based on selected mode."""
        if self.rb_mode_folder.isChecked():
            start_dir = self.txt_source_path.text().strip() or str(Path.cwd())
            folder = QFileDialog.getExistingDirectory(self, "Chọn Thư Mục Chứa Tệp PLM", start_dir)
            if folder:
                self.txt_source_path.setText(folder)
                self._scan_folder_files()
        else:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Chọn Các Tệp PLM Excel",
                str(Path.cwd()),
                "Excel Files (*.xlsx *.xlsm);;All Files (*.*)",
            )
            if files:
                paths = [Path(f) for f in files]
                self.txt_source_path.setText("; ".join(p.name for p in paths))
                self._load_selected_files(paths)

    def _on_rescan(self) -> None:
        """Rescan currently entered path."""
        p_str = self.txt_source_path.text().strip()
        if not p_str:
            return
        if self.rb_mode_folder.isChecked():
            self._scan_folder_files()
        else:
            # If comma or semicolon separated
            paths = [Path(p.strip()) for p in re.split(r"[;,]", p_str) if p.strip()]
            self._load_selected_files(paths)

    def _scan_folder_files(self) -> None:
        """Scan directory for PLM Excel files."""
        dir_str = self.txt_source_path.text().strip()
        if not dir_str:
            return
        d_p = Path(dir_str)
        if not d_p.is_dir():
            QMessageBox.warning(self, "Lỗi đường dẫn", f"Thư mục không tồn tại: {d_p}")
            return

        pattern = "PLM_*.xls*" if self.chk_prefix_plm_only.isChecked() else "*.xls*"
        iterator = d_p.rglob(pattern) if self.chk_recursive.isChecked() else d_p.glob(pattern)

        files: list[Path] = []
        for f in iterator:
            if not f.is_file():
                continue
            if any(b in f.parts for b in ("backup_before_filter", "backupTC14full", ".bak", "old")):
                continue
            if f.name.startswith("~$") or f.stem.endswith("_filtered"):
                continue
            if f.suffix.lower() in (".xlsx", ".xlsm", ".xltm"):
                files.append(f)

        files.sort(key=lambda p: p.name.lower())
        self.detected_files = files
        self._populate_table()

    def _load_selected_files(self, paths: list[Path]) -> None:
        """Load an explicit list of files into table."""
        valid_files = [p for p in paths if p.is_file() and not p.name.startswith("~$")]
        self.detected_files = sorted(valid_files, key=lambda p: p.name.lower())
        self._populate_table()

    def _resolve_model_for_file(self, file_path: Path) -> str:
        """Determine which model should be applied to a specific file."""
        if self.chk_autodetect.isChecked():
            detected = auto_detect_model_from_path(file_path)
            if detected:
                return detected
        return self.combo_model.currentText().strip() or "Unknown"

    def _populate_table(self) -> None:
        """Populate QTableWidget with detected files."""
        self.table_files.setRowCount(0)
        self.lbl_file_count.setText(f"Danh sách tệp phát hiện: {len(self.detected_files)} tệp")

        for r_idx, f in enumerate(self.detected_files):
            self.table_files.insertRow(r_idx)
            
            # STT
            stt_item = QTableWidgetItem(str(r_idx + 1))
            stt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_files.setItem(r_idx, 0, stt_item)

            # Tên tệp
            name_item = QTableWidgetItem(f.name)
            name_item.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
            self.table_files.setItem(r_idx, 1, name_item)

            # Dòng máy áp dụng
            model_name = self._resolve_model_for_file(f)
            model_item = QTableWidgetItem(model_name)
            model_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_files.setItem(r_idx, 2, model_item)

            # Kích thước
            try:
                sz_kb = round(f.stat().st_size / 1024, 1)
                sz_str = f"{sz_kb:,} KB"
            except Exception:
                sz_str = "N/A"
            sz_item = QTableWidgetItem(sz_str)
            sz_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table_files.setItem(r_idx, 3, sz_item)

            # Đường dẫn
            path_item = QTableWidgetItem(str(f.parent))
            self.table_files.setItem(r_idx, 4, path_item)

            # Trạng thái
            status_item = QTableWidgetItem("Sẵn sàng")
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_files.setItem(r_idx, 5, status_item)

        self.btn_start.setEnabled(len(self.detected_files) > 0)

    def _refresh_detected_table_models(self) -> None:
        """Update the model column for all rows in the table."""
        for r_idx in range(self.table_files.rowCount()):
            if r_idx < len(self.detected_files):
                f = self.detected_files[r_idx]
                model_name = self._resolve_model_for_file(f)
                item = self.table_files.item(r_idx, 2)
                if item:
                    item.setText(model_name)

    def _on_start_pruning(self) -> None:
        """Start the background pruning operation."""
        if not self.detected_files:
            QMessageBox.information(self, "Thông báo", "Chưa có tệp nào được chọn để lọc.")
            return

        # Prepare worker file payload
        worker_files: list[tuple[int, Path, str]] = []
        for r_idx in range(self.table_files.rowCount()):
            f = self.detected_files[r_idx]
            model_item = self.table_files.item(r_idx, 2)
            target_model = model_item.text().strip() if model_item else self.combo_model.currentText().strip()
            worker_files.append((r_idx, f, target_model))

        unknown_count = sum(1 for _, _, m in worker_files if not m or m == "Unknown")
        if unknown_count > 0:
            reply = QMessageBox.question(
                self,
                "Xác nhận dòng máy",
                f"Có {unknown_count}/{len(worker_files)} tệp chưa xác định được dòng máy (Model: Unknown).\n"
                "Các tệp này sẽ không được áp dụng quy tắc cắt tỉa nào.\n\n"
                "Bạn có muốn tiếp tục xử lý không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        # Clear logs & reset progress
        self.log_view.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(worker_files))
        self.lbl_progress_status.setText(f"Đang xử lý (0/{len(worker_files)})...")

        # Disable controls during run
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_browse.setEnabled(False)
        self.btn_rescan.setEnabled(False)
        self.btn_close.setEnabled(False)

        # Launch worker thread
        self.worker = StandalonePruneWorker(
            files=worker_files,
            filter_manager=self.filter_manager,
            create_backup=self.chk_backup.isChecked(),
            output_as_new=self.rb_save_new.isChecked(),
        )
        self.worker.progress_signal.connect(self._on_worker_progress)
        self.worker.log_signal.connect(self._on_worker_log)
        self.worker.file_finished_signal.connect(self._on_worker_file_finished)
        self.worker.all_finished_signal.connect(self._on_worker_all_finished)
        self.worker.start()

    def _on_stop_pruning(self) -> None:
        """Request worker thread to cancel."""
        if self.worker and self.worker.isRunning():
            self.lbl_progress_status.setText("Đang yêu cầu dừng...")
            self.worker.cancel()

    @pyqtSlot(int, int, str)
    def _on_worker_progress(self, current: int, total: int, filename: str) -> None:
        self.progress_bar.setValue(current)
        self.lbl_progress_status.setText(f"Xử lý {current}/{total}: {filename}")

    @pyqtSlot(str)
    def _on_worker_log(self, message: str) -> None:
        self.log_view.appendPlainText(message)
        # Auto-scroll to end
        sb = self.log_view.verticalScrollBar()
        if sb:
            sb.setValue(sb.maximum())

    @pyqtSlot(int, dict)
    def _on_worker_file_finished(self, row_idx: int, res: dict) -> None:
        """Update row status in table."""
        status_item = self.table_files.item(row_idx, 5)
        if not status_item:
            return

        if res.get("success"):
            del_r = res.get("deleted_rows", 0)
            status_item.setText(f"✓ Đã cắt {del_r} dòng")
            status_item.setForeground(QColor("#10B981"))
        else:
            status_item.setText("✕ Lỗi")
            status_item.setForeground(QColor("#EF4444"))
            status_item.setToolTip(str(res.get("error", "Lỗi không xác định")))

    @pyqtSlot(dict)
    def _on_worker_all_finished(self, summary: dict) -> None:
        """Re-enable controls and display completion alert."""
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_browse.setEnabled(True)
        self.btn_rescan.setEnabled(True)
        self.btn_close.setEnabled(True)

        succ = summary.get("success", 0)
        err = summary.get("error", 0)
        tot_del = summary.get("total_deleted_rows", 0)
        cancelled = summary.get("cancelled", False)

        status_text = "Đã dừng bởi người dùng" if cancelled else "Hoàn thành lọc BOM!"
        self.lbl_progress_status.setText(f"{status_text} ({succ} thành công, {err} lỗi)")

        msg = (
            f"<h3>{status_text}</h3>"
            f"<p><b>Kết quả xử lý:</b></p>"
            f"<ul>"
            f"<li>Số tệp thành công: <b>{succ}</b></li>"
            f"<li>Số tệp lỗi: <b>{err}</b></li>"
            f"<li>Tổng số dòng linh kiện đã cắt: <b>{tot_del:,} dòng</b></li>"
            f"</ul>"
        )
        if self.chk_backup.isChecked():
            msg += "<p>📁 <i>Các bản sao lưu gốc được lưu an toàn trong thư mục <b>backup_before_filter/</b>.</i></p>"

        if err == 0:
            QMessageBox.information(self, "Hoàn Tất Lọc BOM", msg)
        else:
            QMessageBox.warning(self, "Hoàn Tất Với Cảnh Báo", msg)

    def _on_open_folder(self) -> None:
        """Open the target directory or first file's folder in file explorer."""
        target_dir: Optional[Path] = None
        if self.detected_files:
            target_dir = self.detected_files[0].parent
        elif self.txt_source_path.text().strip():
            p = Path(self.txt_source_path.text().strip())
            target_dir = p if p.is_dir() else p.parent

        if not target_dir or not target_dir.exists():
            target_dir = Path.cwd()

        try:
            if sys.platform == "win32":
                os.startfile(str(target_dir))
            else:
                subprocess.Popen(["xdg-open", str(target_dir)])
        except Exception as ex:
            logger.error("Could not open directory: %s", ex)

    def reject(self) -> None:
        """Handle dialog close/escape request."""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Tiến trình đang chạy",
                "Quá trình lọc BOM đang được thực hiện. Bạn có muốn hủy bỏ và đóng không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            self.worker.cancel()
            self.worker.wait(3000)
        super().reject()

    def closeEvent(self, event: Any) -> None:
        """Prevent thread crash on dialog close."""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Tiến trình đang chạy",
                "Quá trình lọc BOM đang được thực hiện. Bạn có muốn hủy bỏ và đóng không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.worker.cancel()
                self.worker.wait(3000)
                event.accept()
            else:
                event.ignore()
                return
        super().closeEvent(event)
