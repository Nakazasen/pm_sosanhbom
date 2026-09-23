"""Visual Tree BOM Filter Builder Dialog.

Enables engineers and leaders to build PLM BOM decomposition rules visually:
1. Load a Full PLM BOM Excel file (14-column standard).
2. Render an interactive, multi-level hierarchical tree (QTreeWidget).
3. Expand / collapse by levels, search parts, and inspect sub-assemblies.
4. Directly select Phantom assemblies or units on the tree to:
   - Keep assembly, prune child components (Rule 2)
   - Prune assembly and child components entirely (Rule 1/4)
5. Live Simulation Preview: See which sub-assemblies are dimmed or hidden in real-time,
   along with row count reduction KPIs.
6. Extract and persist rules directly into the shared SQLite database (bolocbom_rules).
"""

from __future__ import annotations

import logging
from pathlib import Path
import re
from typing import Any

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
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
    QPushButton,
    QSplitter,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.bom_filter_manager import BOMFilterManager
from src.core.models import BOMNode, BOMTree
from src.core.tree_parser import BOMTreeParser
from src.gui.styles import get_theme_manager
from src.services.machine_dict_service import MachineDictService

logger = logging.getLogger(__name__)


class BOMVisualRuleBuilderDialog(QDialog):
    """Interactive visual tree dialog for creating BOM pruning rules from a real PLM BOM."""

    rules_saved = pyqtSignal(str, int)  # (model_name, rule_count)

    def __init__(
        self,
        filter_manager: BOMFilterManager | None = None,
        initial_model: str | None = None,
        initial_file: Path | str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("🌳 Trình Tạo Bộ Lọc BOM Trực Quan Từ Cây BOM Thực Tế")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setAutoFillBackground(True)

        screen = QApplication.primaryScreen()
        if screen:
            avail = screen.availableGeometry()
            avail_w = avail.width()
            avail_h = avail.height()
        else:
            avail_w, avail_h = 1280, 720

        target_w = min(1120, max(820, int(avail_w * 0.90)))
        target_h = min(560, max(380, int(avail_h * 0.82)))
        self.resize(target_w, target_h)
        self.setMinimumSize(780, 380)

        self.filter_manager = filter_manager or BOMFilterManager()
        self.current_model: str = initial_model or ""
        self.raw_tree: BOMTree | None = None
        self.file_path: Path | None = Path(initial_file) if initial_file else None

        # Specific branch/node action state: {node_key: 'prune_children' | 'prune_node'}
        self.node_rule_actions: dict[str, str] = {}
        # Detailed configured node data: {node_key: {...}}
        self.configured_nodes: dict[str, dict[str, Any]] = {}
        # Existing model rules from database
        self.existing_model_rules: list[Any] = []
        # Rule action state per item_name (for backward compatibility): {item_name: action}
        self.rule_actions: dict[str, str] = {}
        # Part code mapping: {item_name: part_code}
        self.item_part_codes: dict[str, str | None] = {}

        self._init_ui()
        self._apply_theme()

        if self.file_path and self.file_path.exists():
            self._load_file(self.file_path)

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

        icon_lbl = QLabel()
        icon_lbl.setPixmap(get_theme_manager().get_styled_icon("layers").pixmap(26, 26))
        header_layout.addWidget(icon_lbl)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        self.title_lbl = QLabel("Trình Tạo Bộ Lọc BOM Trực Quan Dựa Trên Cây BOM Mô Phỏng")
        self.title_lbl.setFont(QFont("Calibri", 12, QFont.Weight.Bold))
        title_vbox.addWidget(self.title_lbl)

        self.sub_lbl = QLabel(
            "Nạp file BOM PLM Full lớn nhất của dòng máy. Đóng/mở các cấp Level và tick chọn trực tiếp các cụm "
            "Phantom / Cụm lắp ráp cần loại bỏ linh kiện con. Xem trước kết quả cắt tỉa theo thời gian thực."
        )
        self.sub_lbl.setFont(QFont("Calibri", 9))
        title_vbox.addWidget(self.sub_lbl)
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()

        # Model selector in header
        model_box = QHBoxLayout()
        self.model_lbl = QLabel("Dòng máy (Model):")
        self.model_lbl.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        model_box.addWidget(self.model_lbl)

        self.combo_model = QComboBox()
        self.combo_model.setEditable(True)
        self.combo_model.setMinimumWidth(180)
        self._populate_models()
        self.combo_model.currentTextChanged.connect(self._on_model_changed)
        model_box.addWidget(self.combo_model)
        header_layout.addLayout(model_box)

        main_layout.addWidget(self.header_frame, 0)

        # 2. Control Toolbar
        toolbar_group = QGroupBox("1. Thao Tác Nạp Tệp & Điều Khiển Cây BOM")
        tb_layout = QHBoxLayout(toolbar_group)
        tb_layout.setContentsMargins(8, 8, 8, 8)
        tb_layout.setSpacing(8)

        self.btn_load_file = QPushButton("📂 Nạp Tệp BOM PLM Full (.xlsx / .xlsm)")
        self.btn_load_file.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_load_file.setStyleSheet(
            "background-color: #2563EB; color: white; padding: 6px 14px; border-radius: 4px;"
        )
        self.btn_load_file.clicked.connect(self._on_browse_file)
        tb_layout.addWidget(self.btn_load_file)

        self.lbl_loaded_file = QLabel("Chưa nạp tệp BOM nào")
        self.lbl_loaded_file.setFont(QFont("Calibri", 9, QFont.Weight.Medium))
        self.lbl_loaded_file.setStyleSheet("color: #64748B;")
        tb_layout.addWidget(self.lbl_loaded_file)

        tb_layout.addStretch()

        # Tree navigation buttons
        self.btn_expand_l2 = QPushButton("🔽 Cấp 1-2")
        self.btn_expand_l2.setToolTip("Mở rộng các cụm đến Level 2")
        self.btn_expand_l2.clicked.connect(self._expand_to_level_2)
        tb_layout.addWidget(self.btn_expand_l2)

        self.btn_expand_all = QPushButton("🔽 Mở rộng tất cả")
        self.btn_expand_all.clicked.connect(lambda: self.tree_widget.expandAll())
        tb_layout.addWidget(self.btn_expand_all)

        self.btn_collapse_all = QPushButton("🔼 Thu gọn tất cả")
        self.btn_collapse_all.clicked.connect(lambda: self.tree_widget.collapseAll())
        tb_layout.addWidget(self.btn_collapse_all)

        # Search bar
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("🔍 Tìm tên linh kiện hoặc mã part...")
        self.txt_search.setMinimumWidth(220)
        self.txt_search.textChanged.connect(self._on_search_tree)
        tb_layout.addWidget(self.txt_search)

        main_layout.addWidget(toolbar_group, 0)

        # 3. Main Tree View
        tree_group = QGroupBox("2. Mô Phỏng Cấu Trúc Cây BOM & Thiết Lập Quy Tắc Cắt Tỉa")
        tree_layout = QVBoxLayout(tree_group)
        tree_layout.setContentsMargins(8, 8, 8, 8)
        tree_layout.setSpacing(6)

        self.tree_widget = QTreeWidget()
        self.tree_widget.setColumnCount(7)
        self.tree_widget.setHeaderLabels([
            "Cấu Trúc BOM & Tên Linh Kiện (Item Name)",
            "Cấp (Level)",
            "Mã Linh Kiện (Item ID)",
            "Phiên Bản (Rev)",
            "Số Lượng",
            "Có Con?",
            "Quy Tắc Lọc Áp Dụng",
        ])
        self.tree_widget.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree_widget.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_widget.header().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_widget.header().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_widget.header().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_widget.header().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_widget.header().setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        self.tree_widget.setAlternatingRowColors(True)
        self.tree_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        tree_layout.addWidget(self.tree_widget)
        main_layout.addWidget(tree_group, 1)

        # 4. Simulation Stats & Action Bar Card
        self.bottom_frame = QFrame()
        self.bottom_frame.setObjectName("dialog_bottom_frame")
        bottom_box = QHBoxLayout(self.bottom_frame)
        bottom_box.setContentsMargins(10, 6, 10, 6)
        bottom_box.setSpacing(10)

        # Stats KPIs
        self.lbl_stats_total = QLabel("Tổng linh kiện: 0")
        self.lbl_stats_total.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        bottom_box.addWidget(self.lbl_stats_total)

        self.lbl_stats_filtered = QLabel("Sau khi lọc: 0")
        self.lbl_stats_filtered.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.lbl_stats_filtered.setStyleSheet("color: #059669;")
        bottom_box.addWidget(self.lbl_stats_filtered)

        self.lbl_stats_rules = QLabel("Quy tắc đã chọn: 0")
        self.lbl_stats_rules.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.lbl_stats_rules.setStyleSheet("color: #D97706;")
        bottom_box.addWidget(self.lbl_stats_rules)

        bottom_box.addSpacing(15)

        # Preview checkboxes
        self.chk_preview = QCheckBox("👁️ Bật xem trước cắt tỉa")
        self.chk_preview.setChecked(True)
        self.chk_preview.toggled.connect(self._apply_simulation_preview)
        bottom_box.addWidget(self.chk_preview)

        self.chk_hide_pruned = QCheckBox("Ẩn hoàn toàn nhánh bị cắt (không chỉ làm mờ)")
        self.chk_hide_pruned.setChecked(False)
        self.chk_hide_pruned.toggled.connect(self._apply_simulation_preview)
        bottom_box.addWidget(self.chk_hide_pruned)

        bottom_box.addStretch()

        self.btn_save = QPushButton("💾 Lưu Vào Bộ Lọc Của Model")
        self.btn_save.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_save.setStyleSheet(
            "background-color: #059669; color: white; padding: 7px 18px; border-radius: 4px;"
        )
        self.btn_save.clicked.connect(self._on_save_rules)
        bottom_box.addWidget(self.btn_save)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_close.setStyleSheet("padding: 7px 14px; border-radius: 4px;")
        self.btn_close.clicked.connect(self.reject)
        bottom_box.addWidget(self.btn_close)

        main_layout.addWidget(self.bottom_frame, 0)

    def _apply_theme(self) -> None:
        """Apply dark/light styling consistent with Kyocera enterprise theme."""
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
            self.model_lbl.setStyleSheet("color: #F1F5F9;")
            self.setStyleSheet(
                """
                BOMVisualRuleBuilderDialog, QDialog { background-color: #0B0F17; color: #F1F5F9; }
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
                    padding: 4px 8px;
                }
                QLineEdit:focus, QComboBox:focus {
                    border: 1px solid #3B82F6;
                }
                QTreeWidget {
                    background-color: #151D2A;
                    alternate-background-color: #1A2436;
                    color: #F1F5F9;
                    border: 1px solid #2A374A;
                    border-radius: 4px;
                }
                QTreeWidget::item {
                    color: #F1F5F9;
                    padding: 3px 0px;
                }
                QTreeWidget::item:selected {
                    background-color: #1E3A5F;
                    color: #FFFFFF;
                    font-weight: bold;
                }
                QTreeWidget::item:hover {
                    background-color: #1E293B;
                }
                QHeaderView::section {
                    background-color: #1A2436;
                    color: #94A3B8;
                    font-weight: 600;
                    border: 1px solid #2A374A;
                    padding: 4px 6px;
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
                QCheckBox {
                    color: #F1F5F9;
                }
                """
            )
        else:
            self.title_lbl.setStyleSheet("color: #0F172A;")
            self.sub_lbl.setStyleSheet("color: #64748B;")
            self.model_lbl.setStyleSheet("color: #0F172A;")
            self.setStyleSheet(
                """
                BOMVisualRuleBuilderDialog, QDialog { background-color: #F8FAFC; color: #0F172A; }
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
                    padding: 4px 8px;
                }
                QLineEdit:focus, QComboBox:focus {
                    border: 1px solid #2563EB;
                }
                QTreeWidget {
                    background-color: #FFFFFF;
                    alternate-background-color: #F8FAFC;
                    color: #0F172A;
                    border: 1px solid #CBD5E1;
                    border-radius: 4px;
                }
                QTreeWidget::item {
                    color: #0F172A;
                    padding: 3px 0px;
                }
                QTreeWidget::item:selected {
                    background-color: #DBEAFE;
                    color: #1E40AF;
                    font-weight: bold;
                }
                QTreeWidget::item:hover {
                    background-color: #F1F5F9;
                }
                QHeaderView::section {
                    background-color: #F1F5F9;
                    color: #475569;
                    font-weight: 600;
                    border: 1px solid #CBD5E1;
                    padding: 4px 6px;
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
                QCheckBox {
                    color: #0F172A;
                }
                """
            )

    def _populate_models(self) -> None:
        """Populate model combobox from database and dictionary service."""
        models = self.filter_manager.get_all_models()
        try:
            known = MachineDictService().get_model_names()
        except Exception:
            known = []
        all_models = sorted(list(set(models + known)), key=lambda s: s.lower())

        self.combo_model.clear()
        self.combo_model.addItems(all_models)
        if self.current_model:
            idx = self.combo_model.findText(self.current_model, Qt.MatchFlag.MatchFixedString)
            if idx >= 0:
                self.combo_model.setCurrentIndex(idx)
            else:
                self.combo_model.setEditText(self.current_model)

    def _on_model_changed(self, model_name: str) -> None:
        self.current_model = model_name.strip()
        self.btn_save.setText(f"💾 Lưu Vào Bộ Lọc Cho Model: {self.current_model or 'Chưa chọn'}")

    def _on_browse_file(self) -> None:
        """Open file dialog to choose PLM BOM file."""
        fpath, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn Tệp BOM PLM Full Của Dòng Máy",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)",
        )
        if not fpath:
            return

        self._load_file(Path(fpath))

    def _load_file(self, file_path: Path) -> None:
        """Parse BOM file and populate the tree widget."""
        try:
            parser = BOMTreeParser()
            self.raw_tree = parser.parse_file(file_path)
            self.file_path = file_path
            self.lbl_loaded_file.setText(f"📁 {file_path.name}")
            self.lbl_loaded_file.setToolTip(str(file_path))

            # Auto-detect model if empty
            if not self.current_model:
                fname = file_path.name.upper()
                for m in self.filter_manager.get_all_models():
                    if m.upper() in fname:
                        self.combo_model.setCurrentText(m)
                        break

            # Pre-load existing rules for this model into action dictionary
            self._load_existing_model_rules()

            self._build_tree_widget()
            self._expand_to_level_2()
            self._apply_simulation_preview()
        except Exception as ex:
            logger.error("Error loading BOM tree from %s: %s", file_path, ex)
            QMessageBox.critical(self, "Lỗi nạp BOM", f"Không thể phân tích cây BOM từ tệp:\n{ex}")

    def _load_existing_model_rules(self) -> None:
        """Load any existing rules for current model to pre-select them on the tree."""
        self.node_rule_actions.clear()
        self.configured_nodes.clear()
        self.rule_actions.clear()
        self.existing_model_rules = []
        if not self.current_model:
            return

        self.existing_model_rules = self.filter_manager.get_rules_for_model(self.current_model)
        for r in self.existing_model_rules:
            if r.item_name:
                clean_name = r.item_name.strip().upper()
                r_notes = (r.notes or "").lower()
                act = "prune_node" if ("loại bỏ cả cụm" in r_notes or "prune_node" in r_notes) else "prune_children"
                self.rule_actions[clean_name] = act
                if r.part_code:
                    self.item_part_codes[clean_name] = r.part_code

    def _find_matching_model_rule_action(self, node: BOMNode, parent_item: QTreeWidgetItem | None = None) -> str | None:
        """Check if an existing model rule from DB matches this specific node."""
        if not getattr(self, "existing_model_rules", None):
            return None

        node_code = (node.item_id or "").strip().upper()
        node_name = (node.item_name or "").strip().upper()
        parent_code = None
        if parent_item:
            p_data = parent_item.data(0, Qt.ItemDataRole.UserRole) or {}
            parent_code = (p_data.get("item_id") or "").strip().upper()

        for r in self.existing_model_rules:
            r_notes = (r.notes or "").lower()
            m = re.search(r"\[branch:\s*([^\]]+)\]", r.notes or "")
            if m:
                expected_parent = m.group(1).strip().upper()
                if parent_code != expected_parent:
                    continue

            # 1. Match by part_code if specified
            if r.part_code and r.part_code.strip():
                if node_code == r.part_code.strip().upper():
                    return "prune_node" if "loại bỏ cả cụm" in r_notes else "prune_children"

            # 2. Match by item_name (only if valid, not generic "RELEASED" or "(CHƯA ĐẶT TÊN)")
            elif r.item_name and r.item_name.strip():
                r_name = r.item_name.strip().upper()
                if r_name not in ("RELEASED", "(CHƯA ĐẶT TÊN)") and node_name not in ("RELEASED", "(CHƯA ĐẶT TÊN)"):
                    mode = (r.match_mode or "Full_name").strip().lower()
                    if mode == "part_name" and r_name in node_name:
                        return "prune_node" if "loại bỏ cả cụm" in r_notes else "prune_children"
                    elif r_name == node_name:
                        return "prune_node" if "loại bỏ cả cụm" in r_notes else "prune_children"

        return None

    def _build_tree_widget(self) -> None:
        """Construct QTreeWidgetItem nodes recursively from self.raw_tree."""
        self.tree_widget.clear()
        if not self.raw_tree:
            return

        self.tree_widget.setUpdatesEnabled(False)
        try:
            for root_node in self.raw_tree.roots:
                self._add_node_to_tree(root_node, parent_item=None, parent_path="")
        finally:
            self.tree_widget.setUpdatesEnabled(True)

    def _add_node_to_tree(
        self,
        node: BOMNode,
        parent_item: QTreeWidgetItem | None = None,
        parent_path: str = "",
    ) -> QTreeWidgetItem:
        """Create a tree item for node and recurse for children."""
        item = QTreeWidgetItem()
        item.setText(0, node.item_name or "(Chưa đặt tên)")
        item.setText(1, f"L{node.level}")
        item.setText(2, node.item_id or "")
        item.setText(3, getattr(node, "revision", "") or "")
        item.setText(4, str(getattr(node, "quantity", 1)))
        item.setText(5, "Có" if node.has_children else "Không")

        # Alignment
        item.setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
        item.setTextAlignment(2, Qt.AlignmentFlag.AlignCenter)
        item.setTextAlignment(3, Qt.AlignmentFlag.AlignCenter)
        item.setTextAlignment(4, Qt.AlignmentFlag.AlignCenter)
        item.setTextAlignment(5, Qt.AlignmentFlag.AlignCenter)

        # Unique branch/node identifier
        parent_part_code = None
        if parent_item:
            p_data = parent_item.data(0, Qt.ItemDataRole.UserRole) or {}
            parent_part_code = p_data.get("item_id")
        node_key = f"{parent_path}/{node.item_id}_r{node.row_index or id(node)}"
        current_path = f"{parent_path}/{node.item_id}" if parent_path else (node.item_id or "")

        # Store domain data
        item.setData(0, Qt.ItemDataRole.UserRole, {
            "node_key": node_key,
            "level": node.level,
            "item_name": node.item_name,
            "item_id": node.item_id,
            "has_children": node.has_children,
            "row_index": node.row_index,
            "parent_part_code": parent_part_code,
            "parent_path": parent_path,
        })

        # Record part code
        if node.item_name:
            clean_name = node.item_name.strip().upper()
            if clean_name not in self.item_part_codes or not self.item_part_codes[clean_name]:
                self.item_part_codes[clean_name] = node.item_id

        if parent_item:
            parent_item.addChild(item)
        else:
            self.tree_widget.addTopLevelItem(item)

        # Check existing action
        current_action = self.node_rule_actions.get(node_key)
        if not current_action:
            current_action = self._find_matching_model_rule_action(node, parent_item)
            if current_action:
                self.node_rule_actions[node_key] = current_action
                self.configured_nodes[node_key] = {
                    "action": current_action,
                    "item_name": node.item_name or "",
                    "part_code": node.item_id,
                    "level": node.level,
                    "parent_part_code": parent_part_code,
                    "tree_item": item,
                }
                if node.item_name:
                    clean_name = node.item_name.strip().upper()
                    self.rule_actions[clean_name] = current_action
                    self.item_part_codes[clean_name] = node.item_id

        # Cell Widget: Filter Action ComboBox
        action_combo = QComboBox()
        action_combo.addItem("— Bình thường (Giữ)", "none")
        action_combo.addItem("✂️ Cắt con (Rule 2)", "prune_children")
        action_combo.addItem("❌ Bỏ cả cụm (Rule 1/4)", "prune_node")

        idx = 0
        if current_action == "prune_children":
            idx = 1
        elif current_action == "prune_node":
            idx = 2
        action_combo.setCurrentIndex(idx)

        # Connect action change targeting ONLY this specific tree item!
        action_combo.currentIndexChanged.connect(
            lambda index, tree_item=item, combo=action_combo: self._on_tree_item_action_changed(tree_item, combo.itemData(index))
        )
        self.tree_widget.setItemWidget(item, 6, action_combo)

        # Recurse for children
        for child in node.children:
            self._add_node_to_tree(child, parent_item=item, parent_path=current_path)

        return item

    def _on_tree_item_action_changed(self, item: QTreeWidgetItem, new_action: str) -> None:
        """Update action rule for THIS specific tree node/branch and refresh preview."""
        data = item.data(0, Qt.ItemDataRole.UserRole) or {}
        node_key = data.get("node_key")
        if not node_key:
            return

        item_name = data.get("item_name") or ""
        clean_name = item_name.strip().upper()

        if new_action == "none":
            self.node_rule_actions.pop(node_key, None)
            self.configured_nodes.pop(node_key, None)
            still_used = any(
                n.get("item_name", "").strip().upper() == clean_name
                for n in self.configured_nodes.values()
            )
            if not still_used and clean_name:
                self.rule_actions.pop(clean_name, None)
        else:
            self.node_rule_actions[node_key] = new_action
            self.configured_nodes[node_key] = {
                "action": new_action,
                "item_name": item_name,
                "part_code": data.get("item_id"),
                "level": data.get("level"),
                "parent_part_code": data.get("parent_part_code"),
                "tree_item": item,
            }
            if clean_name:
                self.rule_actions[clean_name] = new_action
                self.item_part_codes[clean_name] = data.get("item_id")

        # Re-run simulation preview & stats (without affecting other branches)
        self._apply_simulation_preview()

    def _on_action_changed(self, item_name: str | None, new_action: str) -> None:
        """Programmatic helper to set action for nodes matching item_name (backward compatibility)."""
        if not item_name:
            return

        clean_name = item_name.strip().upper()
        if new_action == "none":
            self.rule_actions.pop(clean_name, None)
        else:
            self.rule_actions[clean_name] = new_action

        target_idx = 0
        if new_action == "prune_children":
            target_idx = 1
        elif new_action == "prune_node":
            target_idx = 2

        def visit(item: QTreeWidgetItem) -> None:
            data = item.data(0, Qt.ItemDataRole.UserRole) or {}
            n_name = (data.get("item_name") or "").strip().upper()
            if n_name == clean_name:
                node_key = data.get("node_key")
                if new_action == "none":
                    self.node_rule_actions.pop(node_key, None)
                    self.configured_nodes.pop(node_key, None)
                else:
                    self.node_rule_actions[node_key] = new_action
                    self.configured_nodes[node_key] = {
                        "action": new_action,
                        "item_name": data.get("item_name") or "",
                        "part_code": data.get("item_id"),
                        "level": data.get("level"),
                        "parent_part_code": data.get("parent_part_code"),
                        "tree_item": item,
                    }
                    self.item_part_codes[clean_name] = data.get("item_id")
                combo = self.tree_widget.itemWidget(item, 6)
                if isinstance(combo, QComboBox) and combo.currentIndex() != target_idx:
                    combo.blockSignals(True)
                    combo.setCurrentIndex(target_idx)
                    combo.blockSignals(False)
            for i in range(item.childCount()):
                visit(item.child(i))

        for i in range(self.tree_widget.topLevelItemCount()):
            visit(self.tree_widget.topLevelItem(i))

        # Re-run simulation preview & stats
        self._apply_simulation_preview()

    def _apply_simulation_preview(self) -> None:
        """Apply visual styling (highlighting, dimming, or hiding) according to current rule actions."""
        preview_enabled = self.chk_preview.isChecked()
        hide_pruned = self.chk_hide_pruned.isChecked()

        total_nodes = 0
        surviving_nodes = 0

        def traverse(item: QTreeWidgetItem, ancestor_action: str | None) -> None:
            nonlocal total_nodes, surviving_nodes
            total_nodes += 1

            data = item.data(0, Qt.ItemDataRole.UserRole) or {}
            node_key = data.get("node_key", "")
            node_action = self.node_rule_actions.get(node_key, "none")
            if node_action == "none" and not self.node_rule_actions:
                # Fallback to rule_actions if node_rule_actions is empty (e.g. programmatic test call before tree populated)
                clean_name = (data.get("item_name") or "").strip().upper()
                node_action = self.rule_actions.get(clean_name, "none")

            # Determine effective action on this node
            # If an ancestor in THIS branch is 'prune_node' or 'prune_children', this child is pruned!
            is_pruned = False
            if ancestor_action in ("prune_node", "prune_children"):
                is_pruned = True
            elif node_action == "prune_node":
                is_pruned = True

            if not is_pruned:
                surviving_nodes += 1

            # Visual representation
            if not preview_enabled:
                item.setHidden(False)
                self._reset_item_visuals(item)
            elif is_pruned:
                if hide_pruned:
                    item.setHidden(True)
                else:
                    item.setHidden(False)
                    self._set_item_dimmed(item)
            else:
                item.setHidden(False)
                if node_action == "prune_children":
                    self._set_item_rule2_highlight(item)
                elif node_action == "prune_node":
                    self._set_item_rule1_highlight(item)
                else:
                    self._reset_item_visuals(item)

            # Recurse to children with updated ancestor state
            # ONLY children of this specific node inherit this node's pruning action!
            next_ancestor = node_action if node_action in ("prune_node", "prune_children") else ancestor_action
            for c_idx in range(item.childCount()):
                traverse(item.child(c_idx), next_ancestor)

        for top_idx in range(self.tree_widget.topLevelItemCount()):
            traverse(self.tree_widget.topLevelItem(top_idx), None)

        # Update KPI statistics
        self.lbl_stats_total.setText(f"Tổng linh kiện: {total_nodes:,}")
        pct_reduction = 0.0
        if total_nodes > 0:
            pct_reduction = ((total_nodes - surviving_nodes) / total_nodes) * 100
        self.lbl_stats_filtered.setText(
            f"Sau khi lọc: {surviving_nodes:,} ({pct_reduction:.1f}% giảm)"
        )
        rule_count = len(self.node_rule_actions) if self.node_rule_actions else len(self.rule_actions)
        self.lbl_stats_rules.setText(f"Quy tắc đã chọn: {rule_count} quy tắc")

    def _reset_item_visuals(self, item: QTreeWidgetItem) -> None:
        """Reset font and background colors to normal."""
        is_dark = get_theme_manager().is_dark()
        font = item.font(0)
        font.setStrikeOut(False)
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, QColor(0, 0, 0, 0))
            item.setForeground(col, QColor("#F1F5F9" if is_dark else "#0F172A"))

    def _set_item_dimmed(self, item: QTreeWidgetItem) -> None:
        """Dim node and strikethrough font to show it is eliminated in the output."""
        is_dark = get_theme_manager().is_dark()
        font = item.font(0)
        font.setStrikeOut(True)
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, QColor(0, 0, 0, 0))
            item.setForeground(col, QColor("#64748B" if is_dark else "#94A3B8"))

    def _set_item_rule2_highlight(self, item: QTreeWidgetItem) -> None:
        """Highlight Rule 2 assembly (kept, but its children pruned)."""
        is_dark = get_theme_manager().is_dark()
        font = item.font(0)
        font.setStrikeOut(False)
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, QColor("#451A03" if is_dark else "#FEF3C7"))
            item.setForeground(col, QColor("#FBBF24" if is_dark else "#D97706"))

    def _set_item_rule1_highlight(self, item: QTreeWidgetItem) -> None:
        """Highlight Rule 1/4 assembly (node and children completely deleted)."""
        is_dark = get_theme_manager().is_dark()
        font = item.font(0)
        font.setStrikeOut(True)
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, QColor("#450A0A" if is_dark else "#FEE2E2"))
            item.setForeground(col, QColor("#F87171" if is_dark else "#EF4444"))

    def _expand_to_level_2(self) -> None:
        """Expand tree nodes up to level 2, collapsing deeper nodes."""
        def visit(item: QTreeWidgetItem) -> None:
            data = item.data(0, Qt.ItemDataRole.UserRole)
            level = data.get("level", 1) if data else 1
            if level <= 2:
                item.setExpanded(True)
            else:
                item.setExpanded(False)
            for i in range(item.childCount()):
                visit(item.child(i))

        for i in range(self.tree_widget.topLevelItemCount()):
            visit(self.tree_widget.topLevelItem(i))

    def _on_search_tree(self, query: str) -> None:
        """Filter / highlight matching nodes on tree."""
        q = query.strip().lower()
        if not q:
            self._apply_simulation_preview()
            return

        def visit(item: QTreeWidgetItem) -> bool:
            name = item.text(0).lower()
            code = item.text(2).lower()
            matches = q in name or q in code

            child_matches = False
            for i in range(item.childCount()):
                if visit(item.child(i)):
                    child_matches = True

            should_show = matches or child_matches
            item.setHidden(not should_show)
            if should_show and child_matches:
                item.setExpanded(True)

            if matches:
                item.setBackground(0, QColor("#38BDF844"))
            else:
                item.setBackground(0, QColor(0, 0, 0, 0))

            return should_show

        for i in range(self.tree_widget.topLevelItemCount()):
            visit(self.tree_widget.topLevelItem(i))

    def _on_save_rules(self) -> None:
        """Extract all chosen pruning rules and persist them into SQLite bolocbom_rules."""
        model_name = self.combo_model.currentText().strip()
        if not model_name:
            QMessageBox.warning(self, "Thiếu dòng máy", "Vui lòng chọn hoặc nhập tên Dòng máy (Model) cần lưu bộ lọc.")
            self.combo_model.setFocus()
            return

        # Collect rules to save from configured nodes or rule_actions
        rules_to_save: list[dict[str, Any]] = []
        if self.configured_nodes:
            for node_key, n_info in self.configured_nodes.items():
                act = n_info.get("action")
                if act and act != "none":
                    rules_to_save.append(n_info)
        elif self.rule_actions:
            for item_name, act in self.rule_actions.items():
                if act and act != "none":
                    rules_to_save.append({
                        "item_name": item_name,
                        "part_code": self.item_part_codes.get(item_name),
                        "action": act,
                        "parent_part_code": None,
                    })

        if not rules_to_save:
            QMessageBox.warning(
                self,
                "Chưa chọn quy tắc",
                "Bạn chưa tick chọn bất kỳ cụm nào để cắt bỏ linh kiện con hoặc loại bỏ. "
                "Vui lòng chọn '✂️ Cắt con' hoặc '❌ Bỏ cả cụm' trên cây trước khi lưu.",
            )
            return

        reply = QMessageBox.question(
            self,
            "Xác nhận lưu bộ lọc",
            f"Bạn có muốn lưu {len(rules_to_save)} quy tắc lọc vừa chọn vào CSDL cho dòng máy '{model_name}'?\n\n"
            "Các quy tắc này sẽ được tự động áp dụng khi thực hiện lọc BOM cho model này.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        saved_count = 0
        try:
            # First, delete previous rules for this model to prevent duplicates
            self.filter_manager.delete_model(model_name)

            for r_info in rules_to_save:
                item_name = r_info.get("item_name") or ""
                part_code = r_info.get("part_code")
                action = r_info.get("action", "prune_children")
                parent_part_code = r_info.get("parent_part_code")
                match_mode = "Full_name"
                note_str = "Tạo tự động từ Cây BOM Mô phỏng"
                if action == "prune_node":
                    note_str += " (Loại bỏ cả cụm)"
                else:
                    note_str += " (Cắt con)"
                if parent_part_code:
                    note_str += f" [branch: {parent_part_code}]"

                rule_id = self.filter_manager.add_rule(
                    model_name=model_name,
                    item_name=item_name if item_name else None,
                    match_mode=match_mode,
                    part_code=part_code,
                    notes=note_str,
                )
                if rule_id > 0:
                    saved_count += 1

            self.rules_saved.emit(model_name, saved_count)
            QMessageBox.information(
                self,
                "Lưu thành công",
                f"Đã lưu thành công {saved_count} quy tắc lọc vào CSDL cho dòng máy '{model_name}'.\n\n"
                "Bạn có thể mở cửa sổ 'Quản lý Bộ lọc BOM' để xem bảng danh sách hoặc tinh chỉnh bất kỳ lúc nào.",
            )
            self.accept()
        except Exception as ex:
            logger.error("Lỗi khi lưu quy tắc vào CSDL: %s", ex)
            QMessageBox.critical(self, "Lỗi lưu quy tắc", f"Không thể lưu vào CSDL:\n{ex}")
