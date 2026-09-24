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
import unicodedata

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QFont, QKeyEvent
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

ACTION_DISPLAY_TEXT: dict[str, str] = {
    "none": "— Bình thường (Giữ)",
    "prune_children": "✂️ Cắt con (Rule 2)",
    "prune_node": "❌ Bỏ cả cụm (Rule 1/4)",
}


LEVEL_REGEX = re.compile(
    r"(?<![a-zA-Z0-9_-])(?:(?:level|lvl|cấp|cap)\s*[:=]?|[lL]\s*[:=]?)\s*([0-9]|1[0-9]|20)(?![a-zA-Z0-9_-])",
    re.IGNORECASE,
)


def _normalize_vietnamese(text: str) -> str:
    """Normalize Vietnamese accented text to lowercase unaccented text."""
    nfkd = unicodedata.normalize("NFKD", text.lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c)).replace("đ", "d")


def _node_matches_query(
    name: str,
    code: str,
    level: int,
    has_children: bool,
    item_type: str,
    action: str,
    raw_query: str,
) -> bool:
    """Evaluate whether a BOM node matches the multi-criteria search query."""
    q = raw_query.strip().lower()
    if not q:
        return True

    # Fast-path 1: Exact part code match
    clean_code = code.strip().lower()
    if clean_code and q == clean_code:
        return True

    # 1. Level matching: e.g. "l1", "l2", "level 2", "lvl 2", "cấp 2", "cap 2", "level: 2"
    level_match = LEVEL_REGEX.search(q)
    if level_match:
        req_level = int(level_match.group(1))
        # If the level doesn't match, verify that the query isn't an exact part code or name match
        if level != req_level:
            if q not in clean_code and q not in name.lower():
                return False
        # Remove level pattern from query for subsequent keyword matching
        q = (q[:level_match.start()] + " " + q[level_match.end():]).strip()
        if not q:
            return level == req_level

    # 2. Phantom / Has Children matching
    is_phantom = (
        has_children
        or (item_type.strip().lower() == "phantom")
        or ("phantom" in name.lower())
        or ("phantom" in item_type.lower())
    )
    phantom_keywords = (
        "#phantom",
        "phantom",
        "#cụm ảo",
        "cụm ảo",
        "#cum ao",
        "cum ao",
        "#has_children",
        "has_children",
        "#có con",
        "#co con",
        "có con",
        "co con",
        "#cụm",
        "#cum",
    )
    for kw in phantom_keywords:
        if kw in q:
            if not is_phantom:
                return False
            q = q.replace(kw, " ").strip()
            if not q:
                return True
            break

    # 3. No children matching
    no_child_keywords = (
        "#không con",
        "không con",
        "#khong con",
        "khong con",
        "không có con",
        "khong co con",
        "#leaf",
        "leaf",
    )
    for kw in no_child_keywords:
        if kw in q:
            if has_children:
                return False
            q = q.replace(kw, " ").strip()
            if not q:
                return True
            break

    # 4. Action / Rule status matching
    rule2_keywords = (
        "#cắt con",
        "#cat con",
        "#cắt",
        "#cat",
        "#prune_children",
        "cắt con",
        "cat con",
        "prune_children",
    )
    for kw in rule2_keywords:
        if kw in q:
            if action != "prune_children":
                return False
            q = q.replace(kw, " ").strip()
            if not q:
                return True
            break

    rule1_keywords = (
        "#bỏ cả cụm",
        "#bo ca cum",
        "#bỏ cụm",
        "#bo cum",
        "#prune_node",
        "bỏ cả cụm",
        "bo ca cum",
        "bỏ cụm",
        "bo cum",
        "prune_node",
    )
    for kw in rule1_keywords:
        if kw in q:
            if action != "prune_node":
                return False
            q = q.replace(kw, " ").strip()
            if not q:
                return True
            break

    any_rule_keywords = (
        "#rule",
        "#rules",
        "#quy tắc",
        "#quy tac",
        "#đã chọn",
        "#da chon",
        "đã chọn",
        "da chon",
        "quy tắc",
        "quy tac",
    )
    for kw in any_rule_keywords:
        if kw in q:
            if action not in ("prune_children", "prune_node"):
                return False
            q = q.replace(kw, " ").strip()
            if not q:
                return True
            break

    # 5. Remaining keywords matching against (name, code, item_type)
    name_norm = name.lower()
    code_norm = clean_code
    type_norm = item_type.lower()
    name_no_acc = _normalize_vietnamese(name)
    tokens = q.split()
    for tok in tokens:
        tok_no_acc = _normalize_vietnamese(tok)
        match_tok = (
            (tok in name_norm)
            or (tok in code_norm)
            or (tok in type_norm)
            or (tok_no_acc in name_no_acc)
        )
        if not match_tok:
            return False

    return True


class FastBOMTreeWidget(QTreeWidget):
    """High-performance QTreeWidget supporting on-demand QComboBox creation.

    Avoids allocating thousands of heavyweight QWidget handles upfront, preventing
    catastrophic GUI freeze and sluggishness when rendering thousands of BOM nodes.
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.dialog: Any | None = None
        self.itemClicked.connect(self._on_item_clicked)
        self.itemDoubleClicked.connect(self._on_item_clicked)

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key.Key_Space, Qt.Key.Key_Return, Qt.Key.Key_Enter):
            curr = self.currentItem()
            if curr:
                combo = self.get_or_create_action_combo(curr, 6)
                combo.showPopup()
                return
        super().keyPressEvent(event)

    def get_or_create_action_combo(self, item: QTreeWidgetItem, col: int = 6) -> QComboBox:
        existing = super().itemWidget(item, col)
        if isinstance(existing, QComboBox):
            return existing

        combo = QComboBox()
        combo.addItem("— Bình thường (Giữ)", "none")
        combo.addItem("✂️ Cắt con (Rule 2)", "prune_children")
        combo.addItem("❌ Bỏ cả cụm (Rule 1/4)", "prune_node")

        dialog = getattr(self, "dialog", None) or self.window()
        act = "none"
        data = item.data(0, Qt.ItemDataRole.UserRole) or {}
        node_key = data.get("node_key")
        if hasattr(dialog, "node_rule_actions") and node_key:
            act = dialog.node_rule_actions.get(node_key, "none")
        elif hasattr(dialog, "rule_actions"):
            clean_name = (data.get("item_name") or "").strip().upper()
            act = dialog.rule_actions.get(clean_name, "none")

        idx = 0
        if act == "prune_children":
            idx = 1
        elif act == "prune_node":
            idx = 2
        combo.setCurrentIndex(idx)

        if hasattr(dialog, "_on_tree_item_action_changed"):
            combo.currentIndexChanged.connect(
                lambda index, tree_item=item, c=combo: dialog._on_tree_item_action_changed(tree_item, c.itemData(index))
            )

        super().setItemWidget(item, col, combo)
        return combo

    def itemWidget(self, item: QTreeWidgetItem, column: int) -> QWidget | None:
        if column == 6:
            return self.get_or_create_action_combo(item, column)
        return super().itemWidget(item, column)

    def _on_item_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        if column == 6:
            combo = self.get_or_create_action_combo(item, column)
            combo.showPopup()


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

        # Search navigation state
        self._search_matching_items: list[QTreeWidgetItem] = []
        self._search_match_index: int = -1

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
        toolbar_group = QGroupBox("1. Thao Tác Nạp Tệp, Tìm Kiếm & Kiểm Tra Quy Tắc")
        tb_main_vbox = QVBoxLayout(toolbar_group)
        tb_main_vbox.setContentsMargins(8, 8, 8, 8)
        tb_main_vbox.setSpacing(6)

        # Row 1: File loading & level navigation
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(8)

        self.btn_load_file = QPushButton("📂 Nạp Tệp BOM PLM Full (.xlsx / .xlsm)")
        self.btn_load_file.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_load_file.setStyleSheet(
            "background-color: #2563EB; color: white; padding: 6px 14px; border-radius: 4px;"
        )
        self.btn_load_file.setAutoDefault(False)
        self.btn_load_file.setDefault(False)
        self.btn_load_file.clicked.connect(self._on_browse_file)
        row1_layout.addWidget(self.btn_load_file)

        self.lbl_loaded_file = QLabel("Chưa nạp tệp BOM nào")
        self.lbl_loaded_file.setFont(QFont("Calibri", 9, QFont.Weight.Medium))
        self.lbl_loaded_file.setStyleSheet("color: #64748B;")
        row1_layout.addWidget(self.lbl_loaded_file)

        row1_layout.addStretch()

        self.btn_expand_l2 = QPushButton("🔽 Cấp 1-2")
        self.btn_expand_l2.setToolTip("Mở rộng các cụm đến Level 2")
        self.btn_expand_l2.setAutoDefault(False)
        self.btn_expand_l2.setDefault(False)
        self.btn_expand_l2.clicked.connect(self._expand_to_level_2)
        row1_layout.addWidget(self.btn_expand_l2)

        self.btn_expand_all = QPushButton("🔽 Mở rộng tất cả")
        self.btn_expand_all.setAutoDefault(False)
        self.btn_expand_all.setDefault(False)
        self.btn_expand_all.clicked.connect(lambda: self.tree_widget.expandAll())
        row1_layout.addWidget(self.btn_expand_all)

        self.btn_collapse_all = QPushButton("🔼 Thu gọn tất cả")
        self.btn_collapse_all.setAutoDefault(False)
        self.btn_collapse_all.setDefault(False)
        self.btn_collapse_all.clicked.connect(lambda: self.tree_widget.collapseAll())
        row1_layout.addWidget(self.btn_collapse_all)

        tb_main_vbox.addLayout(row1_layout)

        # Row 2: Audit Review Toggle + Search Bar + Search Navigation
        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(8)

        self.chk_review_rules = QCheckBox("📋 Chỉ xem cụm đã chọn quy tắc (0)")
        self.chk_review_rules.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.chk_review_rules.setStyleSheet("color: #D97706; font-weight: bold;")
        self.chk_review_rules.setToolTip(
            "Bật chế độ rà soát: Chỉ xem các cụm đã chọn quy tắc (Cắt con / Bỏ cả cụm) "
            "để kiểm tra xem có nhầm lẫn không hoặc điều chỉnh nhanh."
        )
        self.chk_review_rules.toggled.connect(self._on_review_mode_toggled)
        row2_layout.addWidget(self.chk_review_rules)

        self.combo_rule_filter = QComboBox()
        self.combo_rule_filter.addItem("Tất cả quy tắc (Cắt con & Bỏ cụm)", "all")
        self.combo_rule_filter.addItem("✂️ Chỉ Cắt con (Rule 2)", "prune_children")
        self.combo_rule_filter.addItem("❌ Chỉ Bỏ cả cụm (Rule 1/4)", "prune_node")
        self.combo_rule_filter.setEnabled(False)
        self.combo_rule_filter.setToolTip("Lọc theo loại quy tắc khi đang ở chế độ xem lại")
        self.combo_rule_filter.currentIndexChanged.connect(lambda: self._apply_tree_filters())
        row2_layout.addWidget(self.combo_rule_filter)

        row2_layout.addSpacing(6)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText(
            "🔍 Tìm tên, mã part, cấp (L1, L2, cấp 1...), #phantom, có con, cắt con... (Nhấn Enter để chuyển)"
        )
        self.txt_search.setMinimumWidth(320)
        self.txt_search.textChanged.connect(self._on_search_text_changed)
        self.txt_search.returnPressed.connect(self._on_search_next)
        self.txt_search.installEventFilter(self)
        row2_layout.addWidget(self.txt_search, 1)

        self.btn_search_prev = QPushButton("◀")
        self.btn_search_prev.setFixedWidth(30)
        self.btn_search_prev.setToolTip("Kết quả trước (Shift+Enter)")
        self.btn_search_prev.setAutoDefault(False)
        self.btn_search_prev.setDefault(False)
        self.btn_search_prev.clicked.connect(self._on_search_prev)
        row2_layout.addWidget(self.btn_search_prev)

        self.btn_search_next = QPushButton("▶")
        self.btn_search_next.setFixedWidth(30)
        self.btn_search_next.setToolTip("Kết quả kế tiếp (Enter)")
        self.btn_search_next.setAutoDefault(False)
        self.btn_search_next.setDefault(False)
        self.btn_search_next.clicked.connect(self._on_search_next)
        row2_layout.addWidget(self.btn_search_next)

        self.btn_search_clear = QPushButton("✖")
        self.btn_search_clear.setFixedWidth(30)
        self.btn_search_clear.setToolTip("Xóa tìm kiếm")
        self.btn_search_clear.setAutoDefault(False)
        self.btn_search_clear.setDefault(False)
        self.btn_search_clear.clicked.connect(lambda: self.txt_search.clear())
        row2_layout.addWidget(self.btn_search_clear)

        self.lbl_search_status = QLabel("")
        self.lbl_search_status.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.lbl_search_status.setStyleSheet("color: #0284C7;")
        row2_layout.addWidget(self.lbl_search_status)

        tb_main_vbox.addLayout(row2_layout)
        main_layout.addWidget(toolbar_group, 0)

        # 3. Main Tree View
        tree_group = QGroupBox("2. Mô Phỏng Cấu Trúc Cây BOM & Thiết Lập Quy Tắc Cắt Tỉa")
        tree_layout = QVBoxLayout(tree_group)
        tree_layout.setContentsMargins(8, 8, 8, 8)
        tree_layout.setSpacing(6)

        self.tree_widget = FastBOMTreeWidget(self)
        self.tree_widget.dialog = self
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
        self.btn_save.setToolTip(
            "Lưu các cụm đã chọn thành các quy tắc lọc trong 'Danh Sách Quy Tắc Lọc' của Model"
        )
        self.btn_save.setStyleSheet(
            "background-color: #059669; color: white; padding: 7px 18px; border-radius: 4px;"
        )
        self.btn_save.setAutoDefault(False)
        self.btn_save.setDefault(False)
        self.btn_save.clicked.connect(self._on_save_rules)
        bottom_box.addWidget(self.btn_save)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setFont(QFont("Calibri", 9, QFont.Weight.Bold))
        self.btn_close.setStyleSheet("padding: 7px 14px; border-radius: 4px;")
        self.btn_close.setAutoDefault(False)
        self.btn_close.setDefault(False)
        self.btn_close.clicked.connect(self.reject)
        bottom_box.addWidget(self.btn_close)

        main_layout.addWidget(self.bottom_frame, 0)

        # Prevent all QPushButtons from being triggered automatically on Enter key
        for btn in self.findChildren(QPushButton):
            btn.setAutoDefault(False)
            btn.setDefault(False)

    def eventFilter(self, watched: Any, event: Any) -> bool:
        """Filter events on search bar to handle Enter / Shift+Enter navigation."""
        if watched == self.txt_search and event.type() == QKeyEvent.Type.KeyPress:
            if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                    self._on_search_prev()
                else:
                    self._on_search_next()
                return True
            elif event.key() == Qt.Key.Key_Escape:
                if self.txt_search.text():
                    self.txt_search.clear()
                    return True
        return super().eventFilter(watched, event)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        """Handle dialog-level key press events to avoid unwanted default button triggers."""
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            focus = self.focusWidget()
            if focus == self.txt_search or self.txt_search.hasFocus():
                if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                    self._on_search_prev()
                else:
                    self._on_search_next()
                event.accept()
                return
            if isinstance(focus, QPushButton):
                focus.click()
                event.accept()
                return
            if isinstance(focus, QCheckBox):
                focus.toggle()
                event.accept()
                return
            # Prevent default QDialog accept() on Enter
            event.accept()
            return
        elif event.key() == Qt.Key.Key_Escape:
            if self.txt_search.text():
                self.txt_search.clear()
                event.accept()
                return
        super().keyPressEvent(event)

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
        if self.raw_tree is not None:
            self._load_existing_model_rules()
            self._build_tree_widget()
            self._expand_to_level_2()
            self._update_review_mode_ui()
            self._apply_simulation_preview()
            self._apply_tree_filters()

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
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self.lbl_loaded_file.setText("⏳ Đang nạp và phân tích tệp BOM...")
        QApplication.processEvents()
        try:
            parser = BOMTreeParser()
            self.raw_tree = parser.parse_file(file_path)
            self.file_path = file_path

            # Check if BOM has part names
            all_nodes = self.raw_tree.get_all_nodes() if hasattr(self.raw_tree, "get_all_nodes") else []
            has_names = any(
                bool(n.item_name and n.item_name.strip().upper() != (n.item_id or "").strip().upper())
                for n in all_nodes
            )
            if not has_names and all_nodes:
                self.lbl_loaded_file.setText(f"📁 {file_path.name} ⚠️ (Thiếu cột Parts Text)")
                self.lbl_loaded_file.setToolTip(
                    f"{file_path}\n"
                    f"⚠️ Tệp BOM này chưa có cột 'Parts Text' (Tên linh kiện) từ Teamcenter PLM.\n"
                    f"Tên linh kiện đang tạm lấy theo mã. Để hiển thị đầy đủ tên mô tả tiếng Anh,\n"
                    f"vui lòng xuất lại trên Teamcenter với cấu hình cột 'KTCT_Trong'."
                )
            else:
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
            self._update_review_mode_ui()
            self._apply_simulation_preview()
            self._apply_tree_filters()
        except Exception as ex:
            logger.error("Error loading BOM tree from %s: %s", file_path, ex)
            QMessageBox.critical(self, "Lỗi nạp BOM", f"Không thể phân tích cây BOM từ tệp:\n{ex}")
        finally:
            QApplication.restoreOverrideCursor()

    def _load_existing_model_rules(self) -> None:
        """Load any existing rules for current model to pre-select them on the tree."""
        self.node_rule_actions.clear()
        self.configured_nodes.clear()
        self.rule_actions.clear()
        self.existing_model_rules = []
        self._rule_lookup_branch_code: dict[tuple[str, str], str] = {}
        self._rule_lookup_code: dict[str, str] = {}
        self._rule_lookup_branch_name: dict[tuple[str, str], str] = {}
        self._rule_lookup_name: dict[str, str] = {}
        self._rule_substring_rules: list[tuple[str | None, str, str]] = []

        if not self.current_model:
            return

        self.existing_model_rules = self.filter_manager.get_rules_for_model(self.current_model)
        for r in self.existing_model_rules:
            r_notes = (r.notes or "").lower()
            act = "prune_node" if ("loại bỏ cả cụm" in r_notes or "prune_node" in r_notes) else "prune_children"

            expected_parent = None
            m = re.search(r"\[branch:\s*([^\]]+)\]", r.notes or "")
            if m:
                expected_parent = m.group(1).strip().upper()

            if r.part_code and r.part_code.strip():
                clean_code = r.part_code.strip().upper()
                if expected_parent:
                    self._rule_lookup_branch_code[(expected_parent, clean_code)] = act
                else:
                    self._rule_lookup_code[clean_code] = act

            if r.item_name and r.item_name.strip():
                clean_name = r.item_name.strip().upper()
                if clean_name not in ("RELEASED", "(CHƯA ĐẶT TÊN)"):
                    self.rule_actions[clean_name] = act
                    if r.part_code:
                        self.item_part_codes[clean_name] = r.part_code
                    mode = (r.match_mode or "Full_name").strip().lower()
                    if mode == "part_name":
                        self._rule_substring_rules.append((expected_parent, clean_name, act))
                    else:
                        if expected_parent:
                            self._rule_lookup_branch_name[(expected_parent, clean_name)] = act
                        else:
                            self._rule_lookup_name[clean_name] = act

        self._update_review_mode_ui()

    def _find_matching_model_rule_action(
        self,
        node: BOMNode,
        parent_item: QTreeWidgetItem | None = None,
        parent_part_code: str | None = None,
    ) -> str | None:
        """Check if an existing model rule from DB matches this specific node (O(1) indexed lookup)."""
        if not getattr(self, "existing_model_rules", None):
            return None

        if parent_part_code is None and parent_item is not None:
            p_data = parent_item.data(0, Qt.ItemDataRole.UserRole) or {}
            parent_part_code = p_data.get("item_id")

        node_code = (node.item_id or "").strip().upper()
        node_name = (node.item_name or "").strip().upper()
        p_code = parent_part_code.strip().upper() if parent_part_code else None

        # 1. Match by part_code with branch
        if p_code and node_code and (p_code, node_code) in getattr(self, "_rule_lookup_branch_code", {}):
            return self._rule_lookup_branch_code[(p_code, node_code)]

        # 2. Match by part_code without branch
        if node_code and node_code in getattr(self, "_rule_lookup_code", {}):
            return self._rule_lookup_code[node_code]

        # 3. Match by item_name
        if node_name and node_name not in ("RELEASED", "(CHƯA ĐẶT TÊN)"):
            if p_code and (p_code, node_name) in getattr(self, "_rule_lookup_branch_name", {}):
                return self._rule_lookup_branch_name[(p_code, node_name)]
            if node_name in getattr(self, "_rule_lookup_name", {}):
                return self._rule_lookup_name[node_name]

            # 4. Substring match (part_name mode)
            if getattr(self, "_rule_substring_rules", None):
                for exp_p, substr, act in self._rule_substring_rules:
                    if exp_p and p_code != exp_p:
                        continue
                    if substr in node_name:
                        return act

        return None

    def _build_tree_widget(self) -> None:
        """Construct QTreeWidgetItem nodes recursively from self.raw_tree in high-performance batch mode."""
        self.tree_widget.clear()
        if not self.raw_tree:
            return

        self.tree_widget.setUpdatesEnabled(False)
        try:
            top_items: list[QTreeWidgetItem] = []
            for root_node in self.raw_tree.roots:
                top_items.append(self._create_tree_item(root_node, parent_part_code=None, parent_path=""))
            self.tree_widget.addTopLevelItems(top_items)
        finally:
            self.tree_widget.setUpdatesEnabled(True)

    def _create_tree_item(
        self,
        node: BOMNode,
        parent_part_code: str | None = None,
        parent_path: str = "",
    ) -> QTreeWidgetItem:
        """Create a tree item for node, match rules, set text, and recurse for children."""
        item = QTreeWidgetItem()
        item_name = (node.item_name or "").strip()
        item_id = (node.item_id or "").strip()

        # If item_name exists and is distinct from item_id, use it
        if item_name and item_name.upper() != item_id.upper():
            item_display_name = item_name
        else:
            # Fallback when BOM export lacks 'Parts Text' description in PLM export
            mach_info = (
                self._machine_dict.extract_and_lookup_material(item_id)
                if (node.level == 0 and hasattr(self, "_machine_dict"))
                else None
            )
            if mach_info and mach_info.machine_name:
                item_display_name = f"{item_id} [{mach_info.machine_name}] (Chưa có tên PLM)"
            elif item_id:
                item_display_name = f"{item_id} (Chưa có tên PLM)"
            else:
                item_display_name = "(Chưa đặt tên)"
            item.setToolTip(
                0,
                f"Mã linh kiện: {item_id or 'N/A'}\n"
                f"⚠️ Tệp BOM này chưa có cột 'Parts Text' (Tên linh kiện) từ Teamcenter PLM.\n"
                f"Vui lòng xuất trên Teamcenter với cấu hình cột 'KTCT_Trong' để có đầy đủ tên mô tả tiếng Anh.",
            )

        item.setText(0, item_display_name)
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
        node_key = f"{parent_path}/{node.item_id}_r{node.row_index or id(node)}"
        current_path = f"{parent_path}/{node.item_id}" if parent_path else (node.item_id or "")

        # Store domain data
        item.setData(0, Qt.ItemDataRole.UserRole, {
            "node_key": node_key,
            "level": node.level,
            "item_name": node.item_name or node.item_id or "",
            "item_id": node.item_id,
            "has_children": node.has_children,
            "item_type": getattr(node, "item_type", "") or "",
            "row_index": node.row_index,
            "parent_part_code": parent_part_code,
            "parent_path": parent_path,
        })

        # Record part code
        if node.item_name:
            clean_name = node.item_name.strip().upper()
            if clean_name not in self.item_part_codes or not self.item_part_codes[clean_name]:
                self.item_part_codes[clean_name] = node.item_id
        elif node.item_id:
            clean_id = node.item_id.strip().upper()
            if clean_id not in self.item_part_codes or not self.item_part_codes[clean_id]:
                self.item_part_codes[clean_id] = node.item_id

        # Check existing action
        current_action = self.node_rule_actions.get(node_key)
        if not current_action:
            current_action = self._find_matching_model_rule_action(node, parent_part_code=parent_part_code)
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

        # Display filter action text in column 6
        act_text = ACTION_DISPLAY_TEXT.get(current_action or "none", ACTION_DISPLAY_TEXT["none"])
        item.setText(6, act_text)
        item.setTextAlignment(6, Qt.AlignmentFlag.AlignCenter)

        # Recurse for children
        my_part_code = node.item_id
        for child in node.children:
            child_item = self._create_tree_item(child, parent_part_code=my_part_code, parent_path=current_path)
            item.addChild(child_item)

        return item

    def _add_node_to_tree(
        self,
        node: BOMNode,
        parent_item: QTreeWidgetItem | None = None,
        parent_path: str = "",
    ) -> QTreeWidgetItem:
        """Create a tree item for node and recurse for children (backward-compatible API)."""
        parent_part_code = None
        if parent_item:
            p_data = parent_item.data(0, Qt.ItemDataRole.UserRole) or {}
            parent_part_code = p_data.get("item_id")

        item = self._create_tree_item(node, parent_part_code=parent_part_code, parent_path=parent_path)
        if parent_item:
            parent_item.addChild(item)
        else:
            self.tree_widget.addTopLevelItem(item)
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

        # Update text in column 6
        act_text = ACTION_DISPLAY_TEXT.get(new_action, ACTION_DISPLAY_TEXT["none"])
        item.setText(6, act_text)

        # Update review mode UI count
        self._update_review_mode_ui()

        # Re-run simulation preview & tree filters
        self._apply_simulation_preview()
        self._apply_tree_filters()

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

        act_text = ACTION_DISPLAY_TEXT.get(new_action, ACTION_DISPLAY_TEXT["none"])

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
                item.setText(6, act_text)
                raw_combo = super(FastBOMTreeWidget, self.tree_widget).itemWidget(item, 6)
                if isinstance(raw_combo, QComboBox) and raw_combo.currentIndex() != target_idx:
                    raw_combo.blockSignals(True)
                    raw_combo.setCurrentIndex(target_idx)
                    raw_combo.blockSignals(False)
            for i in range(item.childCount()):
                visit(item.child(i))

        for i in range(self.tree_widget.topLevelItemCount()):
            visit(self.tree_widget.topLevelItem(i))

        # Update review mode UI count
        self._update_review_mode_ui()

        # Re-run simulation preview & tree filters
        self._apply_simulation_preview()
        self._apply_tree_filters()

    def _apply_simulation_preview(self, preserve_hidden: bool = False) -> None:
        """Apply visual styling (highlighting, dimming, or hiding) according to current rule actions."""
        preview_enabled = self.chk_preview.isChecked()
        hide_pruned = self.chk_hide_pruned.isChecked()

        total_nodes = 0
        surviving_nodes = 0

        self.tree_widget.setUpdatesEnabled(False)
        try:
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

                # Visual state determination
                if not preview_enabled:
                    target_state = "normal"
                elif node_action == "prune_node":
                    target_state = "hidden" if hide_pruned else "rule1"
                elif node_action == "prune_children":
                    target_state = "rule2"
                elif is_pruned:
                    target_state = "hidden" if hide_pruned else "dimmed"
                else:
                    target_state = "normal"

                current_state = getattr(item, "_visual_state", None)
                if current_state != target_state:
                    item._visual_state = target_state
                    if target_state == "hidden":
                        item.setHidden(True)
                    elif target_state == "dimmed":
                        if not preserve_hidden or not item.isHidden():
                            item.setHidden(False)
                        self._set_item_dimmed(item)
                    elif target_state == "rule2":
                        if not preserve_hidden or not item.isHidden():
                            item.setHidden(False)
                        self._set_item_rule2_highlight(item)
                    elif target_state == "rule1":
                        if not preserve_hidden or not item.isHidden():
                            item.setHidden(False)
                        self._set_item_rule1_highlight(item)
                    else:  # normal
                        if not preserve_hidden or not item.isHidden():
                            item.setHidden(False)
                        self._reset_item_visuals(item)

                # Recurse to children with updated ancestor state
                next_ancestor = node_action if node_action in ("prune_node", "prune_children") else ancestor_action
                for c_idx in range(item.childCount()):
                    traverse(item.child(c_idx), next_ancestor)

            for top_idx in range(self.tree_widget.topLevelItemCount()):
                traverse(self.tree_widget.topLevelItem(top_idx), None)
        finally:
            self.tree_widget.setUpdatesEnabled(True)

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

    def _get_style_resources(self) -> dict[str, Any]:
        """Cache QBrush and QFont objects to avoid tens of thousands of allocations per styling pass."""
        is_dark = get_theme_manager().is_dark()
        cache = getattr(self, "_style_cache", None)
        if cache and cache.get("is_dark") == is_dark:
            return cache

        font_normal = self.font()
        font_strike = QFont(font_normal)
        font_strike.setStrikeOut(True)

        res = {
            "is_dark": is_dark,
            "font_normal": font_normal,
            "font_strike": font_strike,
            "trans_bg": QBrush(QColor(0, 0, 0, 0)),
            "normal_fg": QBrush(QColor("#F1F5F9" if is_dark else "#0F172A")),
            "dimmed_fg": QBrush(QColor("#64748B" if is_dark else "#94A3B8")),
            "rule2_bg": QBrush(QColor("#451A03" if is_dark else "#FEF3C7")),
            "rule2_fg": QBrush(QColor("#FBBF24" if is_dark else "#D97706")),
            "rule1_bg": QBrush(QColor("#450A0A" if is_dark else "#FEE2E2")),
            "rule1_fg": QBrush(QColor("#F87171" if is_dark else "#EF4444")),
        }
        self._style_cache = res
        return res

    def _reset_item_visuals(self, item: QTreeWidgetItem) -> None:
        """Reset font and background colors to normal."""
        res = self._get_style_resources()
        font = res["font_normal"]
        bg = res["trans_bg"]
        fg = res["normal_fg"]
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, bg)
            item.setForeground(col, fg)

    def _set_item_dimmed(self, item: QTreeWidgetItem) -> None:
        """Dim node and strikethrough font to show it is eliminated in the output."""
        res = self._get_style_resources()
        font = res["font_strike"]
        bg = res["trans_bg"]
        fg = res["dimmed_fg"]
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, bg)
            item.setForeground(col, fg)

    def _set_item_rule2_highlight(self, item: QTreeWidgetItem) -> None:
        """Highlight Rule 2 assembly (kept, but its children pruned)."""
        res = self._get_style_resources()
        font = res["font_normal"]
        bg = res["rule2_bg"]
        fg = res["rule2_fg"]
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, bg)
            item.setForeground(col, fg)

    def _set_item_rule1_highlight(self, item: QTreeWidgetItem) -> None:
        """Highlight Rule 1/4 assembly (node and children completely deleted)."""
        res = self._get_style_resources()
        font = res["font_strike"]
        bg = res["rule1_bg"]
        fg = res["rule1_fg"]
        for col in range(6):
            item.setFont(col, font)
            item.setBackground(col, bg)
            item.setForeground(col, fg)

    def _expand_to_level_2(self) -> None:
        """Expand tree nodes up to level 2, collapsing deeper nodes."""
        self.tree_widget.setUpdatesEnabled(False)
        try:
            def visit(item: QTreeWidgetItem) -> None:
                data = item.data(0, Qt.ItemDataRole.UserRole)
                level = data.get("level", 1) if data else 1
                if level <= 2:
                    item.setExpanded(True)
                    for i in range(item.childCount()):
                        visit(item.child(i))
                else:
                    item.setExpanded(False)

            for i in range(self.tree_widget.topLevelItemCount()):
                visit(self.tree_widget.topLevelItem(i))
        finally:
            self.tree_widget.setUpdatesEnabled(True)

    def _update_review_mode_ui(self) -> None:
        """Update review mode checkbox text with current rule count."""
        count = len(self.configured_nodes) if self.configured_nodes else len(self.rule_actions)
        self.chk_review_rules.setText(f"📋 Chỉ xem cụm đã chọn quy tắc ({count})")

    def _on_review_mode_toggled(self, checked: bool) -> None:
        """Handle toggle of review / audit mode."""
        self.combo_rule_filter.setEnabled(checked)
        self._apply_tree_filters()

    def _on_search_text_changed(self, text: str) -> None:
        """Triggered when search input text changes."""
        self._apply_tree_filters()

    def _on_search_next(self) -> None:
        """Cycle to next search match in tree."""
        if not self._search_matching_items:
            return
        total = len(self._search_matching_items)
        self._search_match_index = (self._search_match_index + 1) % total
        item = self._search_matching_items[self._search_match_index]
        self.tree_widget.setCurrentItem(item)
        self.tree_widget.scrollToItem(item)
        if self.txt_search.text().strip():
            self.lbl_search_status.setText(f"🔍 {self._search_match_index + 1}/{total} kết quả")

    def _on_search_prev(self) -> None:
        """Cycle to previous search match in tree."""
        if not self._search_matching_items:
            return
        total = len(self._search_matching_items)
        self._search_match_index = (self._search_match_index - 1) % total
        item = self._search_matching_items[self._search_match_index]
        self.tree_widget.setCurrentItem(item)
        self.tree_widget.scrollToItem(item)
        if self.txt_search.text().strip():
            self.lbl_search_status.setText(f"🔍 {self._search_match_index + 1}/{total} kết quả")

    def _on_search_tree(self, query: str) -> None:
        """Filter / highlight matching nodes on tree (backward-compatible API)."""
        if self.txt_search.text() != query:
            self.txt_search.blockSignals(True)
            self.txt_search.setText(query)
            self.txt_search.blockSignals(False)
        self._apply_tree_filters()

    def _apply_tree_filters(self) -> None:
        """Filter and highlight tree nodes based on search query and review mode toggle."""
        query = self.txt_search.text().strip()
        review_mode = self.chk_review_rules.isChecked()
        rule_filter_type = self.combo_rule_filter.currentData() if review_mode else "all"

        self._search_matching_items = []

        if not query and not review_mode:
            # Full normal view
            self.lbl_search_status.setText("")
            self._search_match_index = -1
            self.tree_widget.setUpdatesEnabled(False)
            try:
                def restore(item: QTreeWidgetItem) -> None:
                    item._visual_state = None
                    item.setBackground(0, QColor(0, 0, 0, 0))
                    item.setHidden(False)
                    for i in range(item.childCount()):
                        restore(item.child(i))

                for i in range(self.tree_widget.topLevelItemCount()):
                    restore(self.tree_widget.topLevelItem(i))
            finally:
                self.tree_widget.setUpdatesEnabled(True)

            self._apply_simulation_preview(preserve_hidden=False)
            return

        self.tree_widget.setUpdatesEnabled(False)
        try:
            def visit(item: QTreeWidgetItem, under_rule_node: bool = False) -> tuple[bool, bool]:
                data = item.data(0, Qt.ItemDataRole.UserRole) or {}
                node_key = data.get("node_key", "")
                name = item.text(0)
                code = item.text(2)
                level = data.get("level", 1)
                has_children = data.get("has_children", False)
                item_type = data.get("item_type", "")
                action = self.node_rule_actions.get(node_key, "none")
                if action == "none" and not self.node_rule_actions:
                    clean_name = (data.get("item_name") or "").strip().upper()
                    action = self.rule_actions.get(clean_name, "none")

                # Review mode check
                is_rule_node = False
                if action != "none":
                    if rule_filter_type == "all":
                        is_rule_node = True
                    elif rule_filter_type == "prune_children" and action == "prune_children":
                        is_rule_node = True
                    elif rule_filter_type == "prune_node" and action == "prune_node":
                        is_rule_node = True

                # Search query check
                matches_search = True
                if query:
                    matches_search = _node_matches_query(
                        name=name,
                        code=code,
                        level=level,
                        has_children=has_children,
                        item_type=item_type,
                        action=action,
                        raw_query=query,
                    )

                if review_mode and query:
                    is_direct_match = is_rule_node and matches_search
                elif review_mode:
                    is_direct_match = is_rule_node
                elif query:
                    is_direct_match = matches_search
                else:
                    is_direct_match = True

                if is_direct_match:
                    self._search_matching_items.append(item)

                child_matches = False
                next_under_rule = under_rule_node or is_direct_match
                for i in range(item.childCount()):
                    c_matches, _ = visit(item.child(i), under_rule_node=next_under_rule)
                    if c_matches:
                        child_matches = True

                has_or_contains_rule = is_direct_match or child_matches

                if review_mode and not query:
                    should_show = has_or_contains_rule or under_rule_node
                    item.setHidden(not should_show)
                    if child_matches:
                        item.setExpanded(True)
                    elif is_direct_match and not child_matches:
                        item.setExpanded(False)
                else:
                    should_show = is_direct_match or child_matches
                    item.setHidden(not should_show)
                    if should_show and child_matches:
                        item.setExpanded(True)

                if is_direct_match and query:
                    item.setBackground(0, QColor("#38BDF855"))
                else:
                    item.setBackground(0, QColor(0, 0, 0, 0))

                return has_or_contains_rule, should_show

            for i in range(self.tree_widget.topLevelItemCount()):
                visit(self.tree_widget.topLevelItem(i))
        finally:
            self.tree_widget.setUpdatesEnabled(True)

        total_matches = len(self._search_matching_items)
        if query:
            if total_matches > 0:
                self._search_match_index = 0
                curr = self._search_matching_items[0]
                self.tree_widget.setCurrentItem(curr)
                self.tree_widget.scrollToItem(curr)
                self.lbl_search_status.setText(f"🔍 1/{total_matches} kết quả")
            else:
                self._search_match_index = -1
                self.lbl_search_status.setText("🔍 0 kết quả")
        elif review_mode:
            self.lbl_search_status.setText(f"📋 {total_matches} cụm đã chọn quy tắc")
            if total_matches > 0:
                self._search_match_index = 0
                curr = self._search_matching_items[0]
                self.tree_widget.setCurrentItem(curr)
                self.tree_widget.scrollToItem(curr)
            else:
                self._search_match_index = -1
        else:
            self.lbl_search_status.setText("")
            self._search_match_index = -1

        self._apply_simulation_preview(preserve_hidden=True)
        if query:
            for m_item in self._search_matching_items:
                m_item.setBackground(0, QColor("#38BDF855"))

    def _on_save_rules(self) -> None:
        """Extract all chosen pruning rules and persist them into SQLite bolocbom_rules."""
        model_name = self.combo_model.currentText().strip()
        if not model_name:
            QMessageBox.warning(self, "Thiếu dòng máy", "Vui lòng chọn hoặc nhập tên Dòng máy (Model) cần lưu bộ lọc.")
            self.combo_model.setFocus()
            return

        # Collect rules to save from configured nodes or rule_actions (deduplicating identical rules)
        rules_to_save: list[dict[str, Any]] = []
        seen_keys: set[tuple[str, str, str, str]] = set()

        if self.configured_nodes:
            for node_key, n_info in self.configured_nodes.items():
                act = n_info.get("action")
                if act and act != "none":
                    item_name = (n_info.get("item_name") or "").strip().upper()
                    part_code = (n_info.get("part_code") or "").strip().upper()
                    parent_part_code = (n_info.get("parent_part_code") or "").strip().upper()
                    dedup_key = (item_name, part_code, act, parent_part_code)
                    if dedup_key not in seen_keys:
                        seen_keys.add(dedup_key)
                        rules_to_save.append(n_info)
        elif self.rule_actions:
            for item_name, act in self.rule_actions.items():
                if act and act != "none":
                    clean_name = item_name.strip().upper()
                    part_code = (self.item_part_codes.get(item_name) or "").strip().upper()
                    dedup_key = (clean_name, part_code, act, "")
                    if dedup_key not in seen_keys:
                        seen_keys.add(dedup_key)
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
            "Xác nhận lưu quy tắc vào bộ lọc Model",
            f"Bạn có muốn tạo và lưu {len(rules_to_save)} quy tắc lọc vừa chọn vào 'Danh Sách Quy Tắc Lọc' của Model '{model_name}'?\n\n"
            "Các quy tắc này sẽ trở thành từ khóa bộ lọc chuẩn trong CSDL và tự động áp dụng khi so sánh / lọc BOM của model này.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # Collect all part codes and names present in the current tree
        all_tree_codes: set[str] = set()
        all_tree_names: set[str] = set()

        def collect_tree_parts(item: QTreeWidgetItem) -> None:
            c = item.text(2).strip().upper()
            if c:
                all_tree_codes.add(c)
            d = item.data(0, Qt.ItemDataRole.UserRole) or {}
            nm = (d.get("item_name") or item.text(0) or "").strip().upper()
            if nm:
                all_tree_names.add(nm)
            for i in range(item.childCount()):
                collect_tree_parts(item.child(i))

        for i in range(self.tree_widget.topLevelItemCount()):
            collect_tree_parts(self.tree_widget.topLevelItem(i))

        # Preserve existing rules from DB that belong to other units not present in this loaded BOM
        preserved_external_rules: list[Any] = []
        if getattr(self, "existing_model_rules", None):
            for r in self.existing_model_rules:
                r_code = (r.part_code or "").strip().upper()
                r_name = (r.item_name or "").strip().upper()
                in_this_tree = (r_code and r_code in all_tree_codes) or (r_name and r_name in all_tree_names)
                if not in_this_tree:
                    preserved_external_rules.append(r)

        saved_count = 0
        try:
            # Delete previous rules for this model and save the consolidated set
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

            # Re-save preserved rules from other units/modules of this model
            for prev_r in preserved_external_rules:
                mm = prev_r.match_mode.value if hasattr(prev_r.match_mode, "value") else str(prev_r.match_mode or "Full_name")
                self.filter_manager.add_rule(
                    model_name=model_name,
                    item_name=prev_r.item_name,
                    match_mode=mm,
                    part_code=prev_r.part_code,
                    notes=prev_r.notes or "Quy tắc lưu trước đó",
                )

            self.rules_saved.emit(model_name, saved_count)
            QMessageBox.information(
                self,
                "Lưu thành công",
                f"Đã lưu thành công {saved_count} quy tắc lọc vào CSDL cho dòng máy '{model_name}'.\n\n"
                "Danh sách từ khóa quy tắc đã được đồng bộ vào '2. Danh Sách Quy Tắc Lọc' của Model.",
            )
            self.accept()
        except Exception as ex:
            logger.error("Lỗi khi lưu quy tắc vào CSDL: %s", ex)
            QMessageBox.critical(self, "Lỗi lưu quy tắc", f"Không thể lưu vào CSDL:\n{ex}")
