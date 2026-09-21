"""Dialog for Viewing, Adding, Editing, and Deleting Members in Shared SQLite Database.

Adheres to Data-Dense Enterprise Dashboard standard:
- Dual-theme aware (Light & Dark mode).
- Displays live LAN connection indicator.
- Table listing with instant search and department filtering.
- Complete form for adding, updating, and deleting team members.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
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
    QVBoxLayout,
    QWidget,
)

from src.core.member_database import MemberDatabaseManager, MemberRecord
from src.gui.styles import get_theme_manager
from src.reporting.excel_generator import STANDARD_SUB_UNITS


class MemberRosterDialog(QDialog):
    """Modern enterprise management dialog for engineering roster in shared SQLite."""

    roster_changed = pyqtSignal()

    def __init__(
        self,
        parent: QWidget | None = None,
        base_dir: Path | None = None,
        remote_db_path: Path | str | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Quản lý Danh sách Thành viên (CSDL SQLite Dùng Chung)")
        self.resize(980, 620)
        self.setMinimumSize(850, 520)

        self.db_manager = MemberDatabaseManager(base_dir=base_dir, remote_db_path=remote_db_path)
        self.theme_mgr = get_theme_manager()
        self._selected_member: MemberRecord | None = None

        self._init_ui()
        self._load_roster_data()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # 1. Top Bar: Header Title & Connection Status Pill
        # ---------------------------------------------------------------------
        top_bar = QHBoxLayout()
        title_lbl = QLabel("Danh mục Nhân sự Kỹ thuật Sản xuất (Roster Management)")
        title_lbl.setFont(QFont("Calibri", 12, QFont.Weight.Bold))
        top_bar.addWidget(title_lbl)

        top_bar.addStretch()

        _, is_remote = self.db_manager.get_active_db_path()
        self.lbl_conn_status = QLabel()
        if is_remote:
            self.lbl_conn_status.setText("● CSDL Mạng: \\\\fstvn01\\...\\ssbom_master.db")
            self.lbl_conn_status.setStyleSheet(
                "background-color: #DCFCE7; color: #15803D; font-weight: bold; padding: 4px 10px; border-radius: 4px; border: 1px solid #86EFAC;"
            )
        else:
            self.lbl_conn_status.setText("⚠️ Ngoại tuyến (Local Cache): Đang dùng bản sao cục bộ")
            self.lbl_conn_status.setStyleSheet(
                "background-color: #FEF3C7; color: #B45309; font-weight: bold; padding: 4px 10px; border-radius: 4px; border: 1px solid #FCD34D;"
            )
        top_bar.addWidget(self.lbl_conn_status)
        main_layout.addLayout(top_bar)

        # ---------------------------------------------------------------------
        # 2. Main Content Splitter (Left: Table + Filter, Right: Form)
        # ---------------------------------------------------------------------
        splitter = QSplitter(Qt.Orientation.Horizontal, self)

        # --- Left: Table & Filters ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.setSpacing(8)

        # Filter bar
        filter_bar = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 Tìm theo tên hoặc mã tài khoản...")
        self.search_edit.textChanged.connect(self._apply_filters)
        filter_bar.addWidget(self.search_edit, stretch=2)

        filter_bar.addWidget(QLabel("Phòng ban:"))
        self.combo_dept_filter = QComboBox()
        self.combo_dept_filter.addItem("Tất cả")
        self.combo_dept_filter.addItems(self.db_manager.get_departments())
        self.combo_dept_filter.currentTextChanged.connect(self._apply_filters)
        filter_bar.addWidget(self.combo_dept_filter, stretch=1)

        self.chk_active_only = QCheckBox("Chỉ đang hoạt động")
        self.chk_active_only.setChecked(False)
        self.chk_active_only.toggled.connect(self._apply_filters)
        filter_bar.addWidget(self.chk_active_only)

        left_layout.addLayout(filter_bar)

        # Member table
        self.member_table = QTableWidget(0, 7)
        self.member_table.setHorizontalHeaderLabels([
            "STT", "Mã thành viên", "Họ và tên", "Phòng ban", "Công đoạn", "Trạng thái", "Ghi chú",
        ])
        self.member_table.verticalHeader().setDefaultSectionSize(30)
        self.member_table.verticalHeader().setMinimumSectionSize(26)
        self.member_table.setShowGrid(True)
        self.member_table.setAlternatingRowColors(True)
        self.member_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.member_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.member_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        m_header = self.member_table.horizontalHeader()
        m_header.setDefaultSectionSize(110)
        self.member_table.setColumnWidth(0, 40)   # STT
        self.member_table.setColumnWidth(1, 115)  # Mã thành viên
        self.member_table.setColumnWidth(2, 125)  # Họ tên
        self.member_table.setColumnWidth(3, 75)   # Phòng ban
        self.member_table.setColumnWidth(4, 120)  # Công đoạn (hỗ trợ nhiều công đoạn)
        self.member_table.setColumnWidth(5, 95)   # Trạng thái
        m_header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)  # Ghi chú

        self.member_table.itemSelectionChanged.connect(self._on_table_selection_changed)
        left_layout.addWidget(self.member_table)
        splitter.addWidget(left_widget)

        # --- Right: Form Panel ---
        right_group = QGroupBox("Chi tiết & Thao tác Thành viên")
        right_layout = QVBoxLayout(right_group)
        right_layout.setSpacing(10)

        form_layout = QFormLayout()
        form_layout.setSpacing(8)

        self.txt_account_id = QLineEdit()
        self.txt_account_id.setPlaceholderText("VD: Son_mecha1 (duy nhất)")
        form_layout.addRow("Mã thành viên (*):", self.txt_account_id)

        self.txt_full_name = QLineEdit()
        self.txt_full_name.setPlaceholderText("VD: Nguyễn Văn Sơn")
        form_layout.addRow("Họ và tên:", self.txt_full_name)

        self.combo_dept = QComboBox()
        self.combo_dept.setEditable(True)
        self.combo_dept.addItems(["Cơ 1", "Cơ 2", "Cơ 3"])
        form_layout.addRow("Phòng ban (*):", self.combo_dept)

        self.combo_sub_unit = QComboBox()
        self.combo_sub_unit.setEditable(True)
        self.combo_sub_unit.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        if self.combo_sub_unit.lineEdit():
            self.combo_sub_unit.lineEdit().setPlaceholderText("VD: LSU hoặc LSU, DRUM")
        self.combo_sub_unit.addItem("")
        self.combo_sub_unit.addItems(STANDARD_SUB_UNITS)
        form_layout.addRow("Công đoạn mặc định:", self.combo_sub_unit)

        self.chk_is_active = QCheckBox("Đang tham gia dự án (Active)")
        self.chk_is_active.setChecked(True)
        form_layout.addRow("Trạng thái:", self.chk_is_active)

        self.txt_notes = QLineEdit()
        self.txt_notes.setPlaceholderText("Ghi chú bổ sung...")
        form_layout.addRow("Ghi chú:", self.txt_notes)

        right_layout.addLayout(form_layout)
        right_layout.addSpacing(6)

        # Action buttons
        self.btn_add = QPushButton("+ Thêm thành viên mới")
        self.btn_add.setIcon(self.theme_mgr.get_styled_icon("user-check"))
        self.btn_add.setStyleSheet(
            "background-color: #2563EB; color: white; font-weight: bold; padding: 6px 12px; border-radius: 4px;"
        )
        self.btn_add.clicked.connect(self._on_add_member)
        right_layout.addWidget(self.btn_add)

        self.btn_update = QPushButton("✏️ Cập nhật thông tin")
        self.btn_update.setIcon(self.theme_mgr.get_styled_icon("save"))
        self.btn_update.setStyleSheet(
            "background-color: #059669; color: white; font-weight: bold; padding: 6px 12px; border-radius: 4px;"
        )
        self.btn_update.clicked.connect(self._on_update_member)
        right_layout.addWidget(self.btn_update)

        self.btn_delete = QPushButton("🗑️ Xóa thành viên")
        self.btn_delete.setIcon(self.theme_mgr.get_styled_icon("trash-2"))
        self.btn_delete.setStyleSheet(
            "background-color: #DC2626; color: white; font-weight: bold; padding: 6px 12px; border-radius: 4px;"
        )
        self.btn_delete.clicked.connect(self._on_delete_member)
        right_layout.addWidget(self.btn_delete)

        self.btn_clear_form = QPushButton("Làm mới ô nhập")
        self.btn_clear_form.setIcon(self.theme_mgr.get_styled_icon("rotate-ccw"))
        self.btn_clear_form.clicked.connect(self._clear_form)
        right_layout.addWidget(self.btn_clear_form)

        right_layout.addStretch()
        splitter.addWidget(right_group)

        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([630, 320])
        main_layout.addWidget(splitter)

        # ---------------------------------------------------------------------
        # 3. Bottom Action Bar
        # ---------------------------------------------------------------------
        bottom_bar = QHBoxLayout()

        self.btn_sync = QPushButton("Đồng bộ & Nạp lại từ CSDL")
        self.btn_sync.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        self.btn_sync.clicked.connect(self._load_roster_data)
        bottom_bar.addWidget(self.btn_sync)

        self.lbl_stats = QLabel("Tổng số: 0 thành viên")
        self.lbl_stats.setStyleSheet("color: #64748B; font-weight: 500;")
        bottom_bar.addWidget(self.lbl_stats)

        bottom_bar.addStretch()

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setMinimumWidth(90)
        self.btn_close.clicked.connect(self.accept)
        bottom_bar.addWidget(self.btn_close)

        main_layout.addLayout(bottom_bar)

    def _load_roster_data(self) -> None:
        """Fetch records from SQLite and render onto table."""
        self._all_records = self.db_manager.get_members()
        self._apply_filters()

        # Update department filter choices
        current_filter = self.combo_dept_filter.currentText()
        depts = ["Tất cả"] + self.db_manager.get_departments()
        self.combo_dept_filter.blockSignals(True)
        self.combo_dept_filter.clear()
        self.combo_dept_filter.addItems(depts)
        idx = self.combo_dept_filter.findText(current_filter)
        if idx >= 0:
            self.combo_dept_filter.setCurrentIndex(idx)
        self.combo_dept_filter.blockSignals(False)

        # Update connection badge
        _, is_remote = self.db_manager.get_active_db_path()
        if is_remote:
            self.lbl_conn_status.setText("● CSDL Mạng: \\\\fstvn01\\...\\ssbom_master.db")
            self.lbl_conn_status.setStyleSheet(
                "background-color: #DCFCE7; color: #15803D; font-weight: bold; padding: 4px 10px; border-radius: 4px; border: 1px solid #86EFAC;"
            )
        else:
            self.lbl_conn_status.setText("⚠️ Ngoại tuyến (Local Cache): Đang dùng bản sao cục bộ")
            self.lbl_conn_status.setStyleSheet(
                "background-color: #FEF3C7; color: #B45309; font-weight: bold; padding: 4px 10px; border-radius: 4px; border: 1px solid #FCD34D;"
            )

    def _apply_filters(self) -> None:
        """Filter table rows by search keyword, department, and active status."""
        search_kw = self.search_edit.text().strip().lower()
        dept = self.combo_dept_filter.currentText()
        active_only = self.chk_active_only.isChecked()

        filtered: list[MemberRecord] = []
        for m in getattr(self, "_all_records", []):
            if dept and dept != "Tất cả" and m.department != dept:
                continue
            if active_only and not m.is_active:
                continue
            if search_kw:
                match_id = search_kw in m.account_id.lower()
                match_name = search_kw in m.full_name.lower()
                match_dept = search_kw in m.department.lower()
                if not (match_id or match_name or match_dept):
                    continue
            filtered.append(m)

        self._render_table_rows(filtered)

    def _render_table_rows(self, records: list[MemberRecord]) -> None:
        """Display records in table."""
        self.member_table.setRowCount(0)
        self.lbl_stats.setText(f"Hiển thị: {len(records)}/{len(getattr(self, '_all_records', []))} thành viên")

        for idx, rec in enumerate(records):
            r = self.member_table.rowCount()
            self.member_table.insertRow(r)

            # STT
            item_stt = QTableWidgetItem(str(idx + 1))
            item_stt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.member_table.setItem(r, 0, item_stt)

            # Account ID
            item_id = QTableWidgetItem(rec.account_id)
            item_id.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
            self.member_table.setItem(r, 1, item_id)

            # Full Name
            self.member_table.setItem(r, 2, QTableWidgetItem(rec.full_name))

            # Department
            item_dept = QTableWidgetItem(rec.department)
            item_dept.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.member_table.setItem(r, 3, item_dept)

            # Default Sub-unit
            item_unit = QTableWidgetItem(rec.default_sub_unit)
            item_unit.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.member_table.setItem(r, 4, item_unit)

            # Active status
            status_str = "● Hoạt động" if rec.is_active else "○ Tạm ngừng"
            item_status = QTableWidgetItem(status_str)
            item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if rec.is_active:
                item_status.setForeground(QColor("#15803D"))
            else:
                item_status.setForeground(QColor("#94A3B8"))
            self.member_table.setItem(r, 5, item_status)

            # Notes
            self.member_table.setItem(r, 6, QTableWidgetItem(rec.notes))

    def _on_table_selection_changed(self) -> None:
        """Populate form when user selects a row in the table."""
        selected_rows = self.member_table.selectionModel().selectedRows()
        if not selected_rows:
            return
        row = selected_rows[0].row()

        acc_id = self.member_table.item(row, 1).text().strip()
        rec = next((m for m in getattr(self, "_all_records", []) if m.account_id == acc_id), None)
        if rec:
            self._selected_member = rec
            self.txt_account_id.setText(rec.account_id)
            self.txt_account_id.setReadOnly(False)  # Allow editing account ID
            self.txt_full_name.setText(rec.full_name)
            self.combo_dept.setCurrentText(rec.department)
            self.combo_sub_unit.setEditText(rec.default_sub_unit)
            self.chk_is_active.setChecked(rec.is_active)
            self.txt_notes.setText(rec.notes)

    def _clear_form(self) -> None:
        """Reset form inputs for adding a new member."""
        self.member_table.clearSelection()
        self._selected_member = None
        self.txt_account_id.setReadOnly(False)
        self.txt_account_id.clear()
        self.txt_full_name.clear()
        self.combo_dept.setCurrentIndex(0)
        self.combo_sub_unit.setCurrentIndex(0)
        self.combo_sub_unit.setEditText("")
        self.chk_is_active.setChecked(True)
        self.txt_notes.clear()
        self.txt_account_id.setFocus()

    def _on_add_member(self) -> None:
        """Handle '+ Thêm thành viên mới' action."""
        acc_id = self.txt_account_id.text().strip()
        dept = self.combo_dept.currentText().strip()
        if not acc_id:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng nhập Mã thành viên / Tài khoản.")
            self.txt_account_id.setFocus()
            return
        if not dept:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn hoặc nhập Phòng ban.")
            self.combo_dept.setFocus()
            return

        rec = MemberRecord(
            account_id=acc_id,
            full_name=self.txt_full_name.text().strip() or acc_id,
            department=dept,
            default_sub_unit=self.combo_sub_unit.currentText().strip(),
            is_active=self.chk_is_active.isChecked(),
            notes=self.txt_notes.text().strip(),
        )

        ok, msg = self.db_manager.add_member(rec)
        if ok:
            QMessageBox.information(self, "Thành công", msg)
            self._load_roster_data()
            self._clear_form()
            self.roster_changed.emit()
        else:
            QMessageBox.critical(self, "Lỗi", msg)

    def _on_update_member(self) -> None:
        """Handle 'Cập nhật thông tin' action, supporting account_id modification."""
        if not self._selected_member:
            selected_rows = self.member_table.selectionModel().selectedRows()
            if selected_rows:
                row = selected_rows[0].row()
                acc_id = self.member_table.item(row, 1).text().strip()
                self._selected_member = next(
                    (m for m in getattr(self, "_all_records", []) if m.account_id == acc_id),
                    None,
                )

        if not self._selected_member:
            QMessageBox.warning(self, "Chưa chọn", "Vui lòng chọn thành viên trong bảng để cập nhật.")
            return

        acc_id = self.txt_account_id.text().strip()
        dept = self.combo_dept.currentText().strip()
        if not acc_id:
            QMessageBox.warning(self, "Thiếu thông tin", "Mã thành viên không được để trống.")
            self.txt_account_id.setFocus()
            return
        if not dept:
            QMessageBox.warning(self, "Thiếu thông tin", "Phòng ban không được để trống.")
            self.combo_dept.setFocus()
            return

        rec = MemberRecord(
            id=self._selected_member.id,
            account_id=acc_id,
            full_name=self.txt_full_name.text().strip() or acc_id,
            department=dept,
            default_sub_unit=self.combo_sub_unit.currentText().strip(),
            is_active=self.chk_is_active.isChecked(),
            notes=self.txt_notes.text().strip(),
        )

        ok, msg = self.db_manager.update_member(rec)
        if ok:
            QMessageBox.information(self, "Thành công", msg)
            self._selected_member = rec
            self._load_roster_data()
            self.roster_changed.emit()
        else:
            QMessageBox.critical(self, "Lỗi", msg)

    def _on_delete_member(self) -> None:
        """Handle 'Xóa thành viên' action."""
        acc_id = self.txt_account_id.text().strip()
        member_id = self._selected_member.id if self._selected_member else None
        if not acc_id and member_id is None:
            QMessageBox.warning(self, "Chưa chọn", "Vui lòng chọn thành viên trong bảng cần xóa.")
            return

        confirm = QMessageBox.question(
            self,
            "Xác nhận xóa",
            f"Bạn có chắc chắn muốn xóa thành viên '{acc_id}' khỏi cơ sở dữ liệu chung không?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        ok, msg = self.db_manager.delete_member(acc_id, member_id=member_id)
        if ok:
            QMessageBox.information(self, "Thành công", msg)
            self._clear_form()
            self._load_roster_data()
            self.roster_changed.emit()
        else:
            QMessageBox.critical(self, "Lỗi", msg)
