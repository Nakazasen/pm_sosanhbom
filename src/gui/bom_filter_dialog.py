"""Dialog for Viewing, Adding, Editing, and Deleting BOM Pruning Rules in Shared SQLite.

Adheres to Enterprise Desktop Standard:
- Dual-theme aware (Light & Dark mode).
- Model selection with search and dynamic rule counts.
- Instant table search by item name or part code.
- Full CRUD operations: Add, Edit, Delete rules.
- Excel Import / Export supporting canonical 'BolocBom' sheet format.
- Factory default restoration for any model.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QCompleter,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.bom_filter_manager import BOMFilterManager, BOMRuleRecord
from src.core.default_rules import DEFAULT_MODEL_RULES
from src.gui.styles import get_theme_manager
from src.services.machine_dict_service import MachineDictService

logger = logging.getLogger(__name__)


class BOMFilterConfigDialog(QDialog):
    """Configuration Dialog for PLM BOM decomposition / pruning rules (BolocBom)."""

    def __init__(
        self,
        filter_manager: BOMFilterManager | None = None,
        initial_model: str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("⚙️ Quản Lý & Cấu Hình Bộ Lọc BOM PLM (BolocBom)")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setAutoFillBackground(True)

        screen = QApplication.primaryScreen()
        if screen:
            avail = screen.availableGeometry()
            avail_w = avail.width()
            avail_h = avail.height()
        else:
            avail_w, avail_h = 1280, 720

        target_w = min(1040, max(820, int(avail_w * 0.88)))
        target_h = min(560, max(400, int(avail_h * 0.82)))
        self.resize(target_w, target_h)
        self.setMinimumSize(780, 380)

        self.filter_manager = filter_manager or BOMFilterManager()
        self.current_model: str = initial_model or ""
        self.current_rules: list[BOMRuleRecord] = []
        self.editing_rule_id: int | None = None

        self._init_ui()
        self._apply_theme()
        self._load_models(select_model=self.current_model)

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # 1. Header Banner Card
        self.header_frame = QFrame()
        self.header_frame.setObjectName("dialog_header_frame")
        header_layout = QHBoxLayout(self.header_frame)
        header_layout.setContentsMargins(12, 8, 12, 8)
        header_layout.setSpacing(10)

        icon_label = QLabel()
        icon_label.setPixmap(get_theme_manager().get_styled_icon("filter").pixmap(26, 26))
        header_layout.addWidget(icon_label)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        self.title_lbl = QLabel("Cấu Hình Bộ Lọc BOM PLM (Sheet BolocBom)")
        self.title_lbl.setFont(QFont("Calibri", 12, QFont.Weight.Bold))
        title_vbox.addWidget(self.title_lbl)

        self.sub_lbl = QLabel(
            "Định nghĩa các cụm/linh kiện con cần loại bỏ khi phân rã BOM PLM cho từng dòng máy. "
            "Dữ liệu được lưu trữ tập trung trên CSDL dùng chung LAN."
        )
        self.sub_lbl.setFont(QFont("Calibri", 9))
        title_vbox.addWidget(self.sub_lbl)
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()

        # Connection status badge
        _, is_remote = self.filter_manager.get_active_db_path()
        status_text = "🟢 LAN Server (UNC)" if is_remote else "🟡 Offline Cache (Local)"
        status_color = "#10B981" if is_remote else "#F59E0B"
        self.status_badge = QLabel(status_text)
        self.status_badge.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.status_badge.setStyleSheet(
            f"background-color: {status_color}22; color: {status_color}; "
            f"border: 1px solid {status_color}; border-radius: 4px; padding: 4px 10px;"
        )
        header_layout.addWidget(self.status_badge)
        main_layout.addWidget(self.header_frame, 0)

        # 2. Main Body Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- Left Panel: Models ---
        left_group = QGroupBox("1. Dòng Máy (Model)")
        left_layout = QVBoxLayout(left_group)
        left_layout.setContentsMargins(8, 8, 8, 8)
        left_layout.setSpacing(6)

        self.model_search_box = QLineEdit()
        self.model_search_box.setPlaceholderText("🔍 Lọc dòng máy...")
        self.model_search_box.textChanged.connect(self._filter_model_list)
        left_layout.addWidget(self.model_search_box)

        self.model_list = QListWidget()
        self.model_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.model_list.itemSelectionChanged.connect(self._on_model_selected)
        left_layout.addWidget(self.model_list)

        model_btn_bar = QHBoxLayout()
        self.btn_add_model = QPushButton("➕ Thêm Model")
        self.btn_add_model.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_add_model.setStyleSheet(
            "background-color: #2563EB; color: white; border-radius: 4px; padding: 5px 10px;"
        )
        self.btn_add_model.clicked.connect(self._on_add_model)

        self.btn_del_model = QPushButton("🗑️ Xóa Model")
        self.btn_del_model.setFont(QFont("Calibri", 9))
        self.btn_del_model.setStyleSheet(
            "background-color: #EF4444; color: white; border-radius: 4px; padding: 5px 10px;"
        )
        self.btn_del_model.clicked.connect(self._on_delete_model)

        model_btn_bar.addWidget(self.btn_add_model)
        model_btn_bar.addWidget(self.btn_del_model)
        left_layout.addLayout(model_btn_bar)

        left_widget = QWidget()
        left_widget_layout = QVBoxLayout(left_widget)
        left_widget_layout.setContentsMargins(0, 0, 0, 0)
        left_widget_layout.addWidget(left_group)
        splitter.addWidget(left_widget)

        # --- Right Panel: Rules Table & Edit Form ---
        right_group = QGroupBox("2. Danh Sách Quy Tắc Lọc")
        right_layout = QVBoxLayout(right_group)
        right_layout.setContentsMargins(8, 8, 8, 8)
        right_layout.setSpacing(8)

        # Search bar for rules
        search_row = QHBoxLayout()
        self.rule_search_box = QLineEdit()
        self.rule_search_box.setPlaceholderText("🔍 Tìm theo tên linh kiện hoặc mã part...")
        self.rule_search_box.textChanged.connect(self._filter_rules_table)
        search_row.addWidget(self.rule_search_box)

        self.rule_count_lbl = QLabel("Tổng số: 0 quy tắc")
        self.rule_count_lbl.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        search_row.addWidget(self.rule_count_lbl)
        right_layout.addLayout(search_row)

        # Rules table
        self.rules_table = QTableWidget()
        self.rules_table.setColumnCount(6)
        self.rules_table.setHorizontalHeaderLabels(
            ["STT", "Tên Linh Kiện / Cụm Cần Lọc", "Chế Độ Khớp", "Mã Part", "Ghi Chú", "Thao Tác"]
        )
        self.rules_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.rules_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.rules_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.rules_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.rules_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.rules_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.rules_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.rules_table.setAlternatingRowColors(True)
        self.rules_table.cellDoubleClicked.connect(self._on_table_double_clicked)
        right_layout.addWidget(self.rules_table)

        # --- Sub-Form: Add / Edit Rule ---
        self.form_group = QGroupBox("3. Thêm Mới / Chỉnh Sửa Quy Tắc Lọc")
        form_layout = QVBoxLayout(self.form_group)
        form_layout.setContentsMargins(8, 8, 8, 8)
        form_layout.setSpacing(6)

        grid_row1 = QHBoxLayout()
        grid_row1.addWidget(QLabel("Tên linh kiện/cụm:"))
        self.txt_item_name = QLineEdit()
        self.txt_item_name.setPlaceholderText("Ví dụ: BOTTLE WASTE, PWB PANEL MAIN ASSY...")
        grid_row1.addWidget(self.txt_item_name, 3)

        grid_row1.addWidget(QLabel("Chế độ so khớp:"))
        self.combo_match_mode = QComboBox()
        self.combo_match_mode.addItem("Khớp chính xác (Full_name)", "Full_name")
        self.combo_match_mode.addItem("Khớp một phần / Chứa từ (Part_name)", "Part_name")
        grid_row1.addWidget(self.combo_match_mode, 2)
        form_layout.addLayout(grid_row1)

        grid_row2 = QHBoxLayout()
        grid_row2.addWidget(QLabel("Mã linh kiện (Part Code):"))
        self.txt_part_code = QLineEdit()
        self.txt_part_code.setPlaceholderText("Để trống nếu áp dụng cho mọi mã...")
        grid_row2.addWidget(self.txt_part_code, 2)

        grid_row2.addWidget(QLabel("Ghi chú:"))
        self.txt_notes = QLineEdit()
        self.txt_notes.setPlaceholderText("Lý do loại bỏ...")
        grid_row2.addWidget(self.txt_notes, 3)
        form_layout.addLayout(grid_row2)

        form_btn_row = QHBoxLayout()
        form_btn_row.addStretch()

        self.btn_save_rule = QPushButton("➕ Thêm Quy Tắc")
        self.btn_save_rule.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_save_rule.setStyleSheet(
            "background-color: #2563EB; color: white; border-radius: 4px; padding: 6px 14px;"
        )
        self.btn_save_rule.clicked.connect(self._on_save_rule)
        form_btn_row.addWidget(self.btn_save_rule)

        self.btn_cancel_edit = QPushButton("❌ Hủy Sửa")
        self.btn_cancel_edit.setFont(QFont("Calibri", 9))
        self.btn_cancel_edit.setVisible(False)
        self.btn_cancel_edit.clicked.connect(self._reset_form)
        form_btn_row.addWidget(self.btn_cancel_edit)

        form_layout.addLayout(form_btn_row)
        right_layout.addWidget(self.form_group)

        right_widget = QWidget()
        right_widget_layout = QVBoxLayout(right_widget)
        right_widget_layout.setContentsMargins(0, 0, 0, 0)
        right_widget_layout.addWidget(right_group)
        splitter.addWidget(right_widget)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([260, 740])
        main_layout.addWidget(splitter, 1)

        # 3. Bottom Action Bar Card
        self.bottom_frame = QFrame()
        self.bottom_frame.setObjectName("dialog_bottom_frame")
        bottom_bar = QHBoxLayout(self.bottom_frame)
        bottom_bar.setContentsMargins(8, 6, 8, 6)
        bottom_bar.setSpacing(8)

        self.btn_visual_builder = QPushButton("🌳 Tạo Lọc Trực Quan Từ Cây BOM...")
        self.btn_visual_builder.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_visual_builder.setStyleSheet(
            "background-color: #059669; color: white; border-radius: 4px; padding: 6px 12px;"
        )
        self.btn_visual_builder.clicked.connect(self._on_open_visual_builder)
        bottom_bar.addWidget(self.btn_visual_builder)

        self.btn_import = QPushButton("📥 Nhập từ Excel (Sheet BolocBom)")
        self.btn_import.setFont(QFont("Calibri", 9))
        self.btn_import.clicked.connect(self._on_import_excel)
        bottom_bar.addWidget(self.btn_import)

        self.btn_export = QPushButton("📤 Xuất ra Excel")
        self.btn_export.setFont(QFont("Calibri", 9))
        self.btn_export.clicked.connect(self._on_export_excel)
        bottom_bar.addWidget(self.btn_export)

        self.btn_reset_default = QPushButton("🔄 Khôi Phục Mặc Định")
        self.btn_reset_default.setFont(QFont("Calibri", 9))
        self.btn_reset_default.setStyleSheet("color: #D97706;")
        self.btn_reset_default.clicked.connect(self._on_reset_default)
        bottom_bar.addWidget(self.btn_reset_default)

        bottom_bar.addStretch()

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_close.setStyleSheet("padding: 6px 18px; border-radius: 4px;")
        self.btn_close.clicked.connect(self.accept)
        bottom_bar.addWidget(self.btn_close)

        main_layout.addWidget(self.bottom_frame, 0)

    def _apply_theme(self) -> None:
        """Apply theme styling consistent with Kyocera Dark/Light standard."""
        is_dark = get_theme_manager().is_dark()

        palette = self.palette()
        bg_col = QColor("#0B0F17" if is_dark else "#F8FAFC")
        fg_col = QColor("#F1F5F9" if is_dark else "#0F172A")
        palette.setColor(self.backgroundRole(), bg_col)
        palette.setColor(self.foregroundRole(), fg_col)
        self.setPalette(palette)

        if is_dark:
            self.title_lbl.setStyleSheet("color: #F1F5F9;")
            self.sub_lbl.setStyleSheet("color: #94A3B8;")
            self.setStyleSheet(
                """
                BOMFilterConfigDialog, QDialog { background-color: #0B0F17; color: #F1F5F9; }
                QFrame#dialog_header_frame {
                    background-color: #151D2A;
                    border: 1px solid #2A374A;
                    border-radius: 6px;
                }
                QFrame#dialog_bottom_frame {
                    background-color: #151D2A;
                    border: 1px solid #2A374A;
                    border-radius: 6px;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #2A374A;
                    border-radius: 6px;
                    margin-top: 6px;
                    padding-top: 10px;
                    color: #E2E8F0;
                    background-color: #111722;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 8px;
                    padding: 0 4px;
                    color: #38BDF8;
                }
                QLineEdit, QComboBox {
                    background-color: #151D2A;
                    color: #F1F5F9;
                    border: 1px solid #2A374A;
                    border-radius: 4px;
                    padding: 5px 8px;
                }
                QLineEdit:focus, QComboBox:focus {
                    border: 1px solid #3B82F6;
                }
                QListWidget {
                    background-color: #151D2A;
                    color: #F1F5F9;
                    border: 1px solid #2A374A;
                    border-radius: 4px;
                }
                QListWidget::item {
                    color: #F1F5F9;
                    padding: 4px 6px;
                }
                QListWidget::item:selected {
                    background-color: #1E3A5F;
                    color: #FFFFFF;
                    font-weight: bold;
                }
                QTableWidget {
                    background-color: #151D2A;
                    alternate-background-color: #1A2436;
                    color: #F1F5F9;
                    border: 1px solid #2A374A;
                    gridline-color: #2A374A;
                    border-radius: 4px;
                }
                QTableWidget::item {
                    color: #F1F5F9;
                    padding: 3px 6px;
                }
                QTableWidget::item:selected {
                    background-color: #1E3A5F;
                    color: #FFFFFF;
                }
                QHeaderView::section {
                    background-color: #1A2436;
                    color: #94A3B8;
                    font-weight: bold;
                    border: 1px solid #2A374A;
                    padding: 4px;
                }
                QPushButton {
                    background-color: #1E293B;
                    color: #F1F5F9;
                    border: 1px solid #334155;
                    border-radius: 4px;
                    padding: 5px 10px;
                }
                QPushButton:hover {
                    background-color: #293548;
                    border-color: #475569;
                }
                """
            )
        else:
            self.title_lbl.setStyleSheet("color: #0F172A;")
            self.sub_lbl.setStyleSheet("color: #64748B;")
            self.setStyleSheet(
                """
                BOMFilterConfigDialog, QDialog { background-color: #F8FAFC; color: #0F172A; }
                QFrame#dialog_header_frame {
                    background-color: #F1F5F9;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                }
                QFrame#dialog_bottom_frame {
                    background-color: #F1F5F9;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                    margin-top: 6px;
                    padding-top: 10px;
                    color: #1E293B;
                    background-color: #FFFFFF;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 8px;
                    padding: 0 4px;
                    color: #0284C7;
                }
                QLineEdit, QComboBox {
                    background-color: #FFFFFF;
                    color: #0F172A;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                    padding: 5px 8px;
                }
                QLineEdit:focus, QComboBox:focus {
                    border: 1px solid #2563EB;
                }
                QListWidget {
                    background-color: #FFFFFF;
                    color: #0F172A;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                }
                QListWidget::item {
                    color: #0F172A;
                    padding: 4px 6px;
                }
                QListWidget::item:selected {
                    background-color: #DBEAFE;
                    color: #1E40AF;
                    font-weight: bold;
                }
                QTableWidget {
                    background-color: #FFFFFF;
                    alternate-background-color: #F8FAFC;
                    color: #0F172A;
                    border: 1px solid #CBD5E1;
                    gridline-color: #F1F5F9;
                    border-radius: 4px;
                }
                QTableWidget::item {
                    color: #0F172A;
                    padding: 3px 6px;
                }
                QTableWidget::item:selected {
                    background-color: #DBEAFE;
                    color: #1E40AF;
                }
                QHeaderView::section {
                    background-color: #F1F5F9;
                    color: #475569;
                    font-weight: bold;
                    border: 1px solid #CBD5E1;
                    padding: 4px;
                }
                QPushButton {
                    background-color: #F1F5F9;
                    color: #0F172A;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                    padding: 5px 10px;
                }
                QPushButton:hover {
                    background-color: #E2E8F0;
                    border-color: #94A3B8;
                }
                """
            )

    def _load_models(self, select_model: str | None = None) -> None:
        """Load distinct machine models from DB into left list widget."""
        self.model_list.clear()
        models = self.filter_manager.get_all_models()

        selected_row = -1
        for idx, model_name in enumerate(models):
            rules = self.filter_manager.get_rules_for_model(model_name)
            item = QListWidgetItem(f"{model_name}  ({len(rules)} quy tắc)")
            item.setData(Qt.ItemDataRole.UserRole, model_name)
            self.model_list.addItem(item)

            if select_model and model_name.strip().lower() == select_model.strip().lower():
                selected_row = idx

        if selected_row >= 0:
            self.model_list.setCurrentRow(selected_row)
        elif self.model_list.count() > 0:
            self.model_list.setCurrentRow(0)

    def _filter_model_list(self, query: str) -> None:
        """Filter models in the left list widget."""
        q = query.strip().lower()
        for i in range(self.model_list.count()):
            item = self.model_list.item(i)
            model_name = item.data(Qt.ItemDataRole.UserRole) or ""
            item.setHidden(q not in model_name.lower())

    def _on_model_selected(self) -> None:
        """Handle selection change in the model list."""
        current_item = self.model_list.currentItem()
        if not current_item:
            self.current_model = ""
            self.current_rules = []
            self.rules_table.setRowCount(0)
            self.rule_count_lbl.setText("Tổng số: 0 quy tắc")
            return

        self.current_model = current_item.data(Qt.ItemDataRole.UserRole) or ""
        self._reset_form()
        self._refresh_rules_table()

    def _refresh_rules_table(self) -> None:
        """Load and display rules for the currently selected model."""
        if not self.current_model:
            return

        self.current_rules = self.filter_manager.get_rules_for_model(self.current_model)
        self._filter_rules_table(self.rule_search_box.text())

    def _filter_rules_table(self, query: str = "") -> None:
        """Filter table rows by search keyword in item name or part code."""
        q = query.strip().lower()
        filtered = [
            r for r in self.current_rules
            if not q or (r.item_name and q in r.item_name.lower()) or (r.part_code and q in r.part_code.lower())
        ]

        self.rules_table.setRowCount(len(filtered))
        self.rule_count_lbl.setText(f"Hiển thị: {len(filtered)} / {len(self.current_rules)} quy tắc")

        is_dark = get_theme_manager().is_dark()
        text_color = QColor("#F1F5F9" if is_dark else "#0F172A")
        dim_color = QColor("#94A3B8" if is_dark else "#64748B")

        for row_idx, rule in enumerate(filtered):
            # 0: STT
            stt_item = QTableWidgetItem(str(row_idx + 1))
            stt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            stt_item.setData(Qt.ItemDataRole.UserRole, rule.id)
            stt_item.setForeground(text_color)
            self.rules_table.setItem(row_idx, 0, stt_item)

            # 1: Item Name
            item_name_item = QTableWidgetItem(rule.item_name or "(Bất kỳ)")
            item_name_item.setForeground(text_color if rule.item_name else dim_color)
            self.rules_table.setItem(row_idx, 1, item_name_item)

            # 2: Match Mode
            mode_desc = "Chính xác (Full)" if rule.match_mode == "Full_name" else "Chứa từ (Part)"
            mode_item = QTableWidgetItem(mode_desc)
            mode_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if rule.match_mode == "Full_name":
                mode_item.setForeground(QColor("#60A5FA" if is_dark else "#2563EB"))
            else:
                mode_item.setForeground(QColor("#FBBF24" if is_dark else "#D97706"))
            self.rules_table.setItem(row_idx, 2, mode_item)

            # 3: Part Code
            part_item = QTableWidgetItem(rule.part_code or "-")
            part_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            part_item.setForeground(text_color if rule.part_code else dim_color)
            self.rules_table.setItem(row_idx, 3, part_item)

            # 4: Notes
            notes_item = QTableWidgetItem(rule.notes or "")
            notes_item.setForeground(text_color)
            self.rules_table.setItem(row_idx, 4, notes_item)

            # 5: Actions
            btn_container = QWidget()
            btn_layout = QHBoxLayout(btn_container)
            btn_layout.setContentsMargins(4, 2, 4, 2)
            btn_layout.setSpacing(4)

            btn_edit = QPushButton("✏️")
            btn_edit.setToolTip("Chỉnh sửa quy tắc này")
            btn_edit.setStyleSheet("padding: 2px 6px; font-size: 11px;")
            btn_edit.clicked.connect(lambda checked, r_id=rule.id: self._start_edit_rule(r_id))

            btn_del = QPushButton("🗑️")
            btn_del.setToolTip("Xóa quy tắc này")
            btn_del.setStyleSheet("padding: 2px 6px; font-size: 11px; color: #EF4444;")
            btn_del.clicked.connect(lambda checked, r_id=rule.id, r_name=rule.item_name: self._on_delete_rule(r_id, r_name))

            btn_layout.addWidget(btn_edit)
            btn_layout.addWidget(btn_del)
            self.rules_table.setCellWidget(row_idx, 5, btn_container)

    def _on_table_double_clicked(self, row: int, col: int) -> None:
        """Double clicking a row triggers edit mode."""
        stt_item = self.rules_table.item(row, 0)
        if stt_item:
            rule_id = stt_item.data(Qt.ItemDataRole.UserRole)
            if rule_id:
                self._start_edit_rule(rule_id)

    def _start_edit_rule(self, rule_id: int) -> None:
        """Populate bottom form with rule data for editing."""
        matched = [r for r in self.current_rules if r.id == rule_id]
        if not matched:
            return

        rule = matched[0]
        self.editing_rule_id = rule.id
        self.txt_item_name.setText(rule.item_name or "")
        idx = 0 if rule.match_mode == "Full_name" else 1
        self.combo_match_mode.setCurrentIndex(idx)
        self.txt_part_code.setText(rule.part_code or "")
        self.txt_notes.setText(rule.notes or "")

        self.form_group.setTitle(f"3. Đang Sửa Quy Tắc ID #{rule.id} (Model: {self.current_model})")
        self.btn_save_rule.setText("💾 Lưu Cập Nhật")
        self.btn_save_rule.setStyleSheet(
            "background-color: #059669; color: white; border-radius: 4px; padding: 6px 14px;"
        )
        self.btn_cancel_edit.setVisible(True)
        self.txt_item_name.setFocus()

    def _reset_form(self) -> None:
        """Reset form back to Add mode."""
        self.editing_rule_id = None
        self.txt_item_name.clear()
        self.combo_match_mode.setCurrentIndex(0)
        self.txt_part_code.clear()
        self.txt_notes.clear()

        self.form_group.setTitle(f"3. Thêm Mới Quy Tắc Lọc (Model: {self.current_model or 'Chưa chọn'})")
        self.btn_save_rule.setText("➕ Thêm Quy Tắc")
        self.btn_save_rule.setStyleSheet(
            "background-color: #2563EB; color: white; border-radius: 4px; padding: 6px 14px;"
        )
        self.btn_cancel_edit.setVisible(False)

    def _on_save_rule(self) -> None:
        """Save a new rule or update an existing rule."""
        if not self.current_model:
            QMessageBox.warning(self, "Chưa chọn dòng máy", "Vui lòng chọn một dòng máy ở danh sách bên trái trước.")
            return

        item_name = self.txt_item_name.text().strip() or None
        part_code = self.txt_part_code.text().strip() or None
        match_mode = self.combo_match_mode.currentData() or "Full_name"
        notes = self.txt_notes.text().strip()

        if not item_name and not part_code:
            QMessageBox.warning(
                self,
                "Thiếu thông tin",
                "Bạn cần nhập ít nhất 'Tên linh kiện/cụm' hoặc 'Mã linh kiện (Part Code)'.",
            )
            self.txt_item_name.setFocus()
            return

        if self.editing_rule_id:
            # Update
            success = self.filter_manager.update_rule(
                rule_id=self.editing_rule_id,
                item_name=item_name,
                match_mode=match_mode,
                part_code=part_code,
                notes=notes,
            )
            if success:
                self._reset_form()
                self._refresh_rules_table()
                self._update_model_count_in_list(self.current_model)
            else:
                QMessageBox.critical(self, "Lỗi", "Không thể cập nhật quy tắc vào CSDL.")
        else:
            # Add new
            rule_id = self.filter_manager.add_rule(
                model_name=self.current_model,
                item_name=item_name,
                match_mode=match_mode,
                part_code=part_code,
                notes=notes,
            )
            if rule_id > 0:
                self._reset_form()
                self._refresh_rules_table()
                self._update_model_count_in_list(self.current_model)
            else:
                QMessageBox.critical(self, "Lỗi", "Không thể thêm quy tắc vào CSDL.")

    def _on_delete_rule(self, rule_id: int, item_name: str | None) -> None:
        """Confirm and delete a rule."""
        name_display = item_name or f"ID #{rule_id}"
        reply = QMessageBox.question(
            self,
            "Xác nhận xóa quy tắc",
            f"Bạn có chắc chắn muốn xóa quy tắc lọc '{name_display}' khỏi model '{self.current_model}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            success = self.filter_manager.delete_rule(rule_id)
            if success:
                if self.editing_rule_id == rule_id:
                    self._reset_form()
                self._refresh_rules_table()
                self._update_model_count_in_list(self.current_model)
            else:
                QMessageBox.critical(self, "Lỗi", "Không thể xóa quy tắc khỏi CSDL.")

    def _update_model_count_in_list(self, model_name: str) -> None:
        """Update badge count on the model item in the left list."""
        rules = self.filter_manager.get_rules_for_model(model_name)
        for i in range(self.model_list.count()):
            item = self.model_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == model_name:
                item.setText(f"{model_name}  ({len(rules)} quy tắc)")
                break

    def _on_add_model(self) -> None:
        """Add a new model to the list."""
        try:
            known_models = MachineDictService().get_model_names()
        except Exception:
            known_models = []
        text, ok = QInputDialog.getText(
            self,
            "Thêm Dòng Máy Mới",
            "Nhập tên dòng máy cần thêm bộ lọc BOM:\n(Ví dụ: 6th Next, Mercury, TASKalfa 4054ci...)",
        )
        if ok and text and text.strip():
            model_name = text.strip()
            # Check if model already exists
            existing_models = [m.lower() for m in self.filter_manager.get_all_models()]
            if model_name.lower() in existing_models:
                QMessageBox.information(self, "Đã tồn tại", f"Dòng máy '{model_name}' đã có trong danh sách.")
                self._load_models(select_model=model_name)
                return

            # Add a placeholder/first rule or seed
            self.filter_manager.add_rule(
                model_name=model_name,
                item_name="BOTTLE WASTE",
                match_mode="Full_name",
                notes="Quy tắc khởi tạo ban đầu",
            )
            self._load_models(select_model=model_name)
            QMessageBox.information(
                self,
                "Thành công",
                f"Đã thêm dòng máy '{model_name}'. Bạn có thể tiếp tục thêm các quy tắc linh kiện cho dòng máy này.",
            )

    def _on_delete_model(self) -> None:
        """Confirm and delete all rules for current model."""
        if not self.current_model:
            return

        reply = QMessageBox.warning(
            self,
            "Xác nhận xóa Model",
            f"CẢNH BÁO: Thao tác này sẽ XÓA TOÀN BỘ ({len(self.current_rules)} quy tắc) của dòng máy '{self.current_model}'.\n\n"
            "Bạn có chắc chắn muốn xóa không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.filter_manager.delete_model(self.current_model)
            self._load_models()

    def _on_reset_default(self) -> None:
        """Reset current model back to factory defaults."""
        if not self.current_model:
            return

        # Check if default rules exist for this model
        clean = self.current_model.strip().lower()
        has_defaults = any(m.lower() == clean for m in DEFAULT_MODEL_RULES)
        if not has_defaults:
            QMessageBox.information(
                self,
                "Không có cấu hình gốc",
                f"Dòng máy '{self.current_model}' là model tự tạo, không có trong 6 bộ quy tắc mặc định ban đầu.",
            )
            return

        reply = QMessageBox.question(
            self,
            "Khôi phục quy tắc mặc định",
            f"Bạn có chắc chắn muốn đặt lại các quy tắc của dòng máy '{self.current_model}' "
            f"về mặc định ban đầu ({len(DEFAULT_MODEL_RULES.get(self.current_model, []))} quy tắc gốc)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            reset_count = self.filter_manager.reset_model_to_default(self.current_model)
            self._refresh_rules_table()
            self._update_model_count_in_list(self.current_model)
            QMessageBox.information(self, "Đã khôi phục", f"Đã khôi phục {reset_count} quy tắc mặc định cho {self.current_model}.")

    def _on_import_excel(self) -> None:
        """Import rules from Excel sheet 'BolocBom'."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn Tệp Excel Chứa Sheet 'BolocBom'",
            "",
            "Excel Files (*.xlsx *.xlsm);;All Files (*.*)",
        )
        if not file_path:
            return

        try:
            imported_count = self.filter_manager.import_from_excel(file_path)
            if imported_count > 0:
                self._load_models(select_model=self.current_model)
                QMessageBox.information(
                    self,
                    "Nhập Excel thành công",
                    f"Đã nhập thành công {imported_count} quy tắc từ tệp:\n{file_path}",
                )
            else:
                QMessageBox.warning(
                    self,
                    "Không tìm thấy dữ liệu",
                    "Không tìm thấy dòng quy tắc hợp lệ trong sheet 'BolocBom'. "
                    "Vui lòng kiểm tra định dạng cột (Model, Item Name, Match Mode, Part Code).",
                )
        except Exception as ex:
            logger.error("Lỗi nhập Excel: %s", ex)
            QMessageBox.critical(self, "Lỗi nhập Excel", f"Không thể đọc file: {ex}")

    def _on_export_excel(self) -> None:
        """Export current model or all models to Excel."""
        out_file, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu Danh Sách Quy Tắc Ra Excel",
            f"BolocBom_{self.current_model or 'All'}.xlsx",
            "Excel Files (*.xlsx);;All Files (*.*)",
        )
        if not out_file:
            return

        try:
            exported_path = self.filter_manager.export_to_excel(out_file, model_name=self.current_model or None)
            QMessageBox.information(
                self,
                "Xuất Excel thành công",
                f"Đã lưu danh sách quy tắc ra tệp:\n{exported_path}",
            )
        except Exception as ex:
            logger.error("Lỗi xuất Excel: %s", ex)
            QMessageBox.critical(self, "Lỗi xuất Excel", f"Không thể xuất file: {ex}")

    def _on_open_visual_builder(self) -> None:
        """Open the visual tree simulation dialog to create rules interactively from a real BOM."""
        from src.gui.bom_visual_builder_dialog import BOMVisualRuleBuilderDialog
        dlg = BOMVisualRuleBuilderDialog(
            filter_manager=self.filter_manager,
            initial_model=self.current_model,
            parent=self,
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._load_models(select_model=dlg.current_model)
            self._refresh_rules_table()
