"""Main PyQt6 Desktop GUI Application for Kyocera SSBOM Automation.

Provides:
- Main QMainWindow container with tabbed/role-based navigation between:
  * Leader Management Workspace (LeaderView)
  * Member Input Workspace (MemberView)
- Status bar with connection indicators and environment status.
- Real-time logging console dock capturing system events.
- Menu bar, toolbar, settings dialog launcher, and non-tech error dialogs.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QObject, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QApplication,
    QDockWidget,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QStatusBar,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from src.gui.leader_view import LeaderWorkspaceView
from src.gui.member_view import MemberWorkspaceView
from src.gui.settings_dialog import SettingsDialog

logger = logging.getLogger(__name__)


class LogEmitter(QObject):
    """QObject wrapper emitting string signals safely across threads."""

    new_record = pyqtSignal(str)


class QLogHandler(logging.Handler):
    """Custom logging handler emitting formatted messages to PyQt widgets safely."""

    def __init__(self) -> None:
        super().__init__()
        self.emitter = LogEmitter()
        self.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))

    @property
    def new_record(self) -> Any:
        return self.emitter.new_record

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self.emitter.new_record.emit(msg)
        except Exception:
            pass


class SSBOMMainWindow(QMainWindow):
    """Main application window for Kyocera automated BOM comparison."""

    def closeEvent(self, event: Any) -> None:
        """Ensure logging handler is detached cleanly on window close."""
        if hasattr(self, "log_handler") and self.log_handler in logging.getLogger().handlers:
            logging.getLogger().removeHandler(self.log_handler)
        super().closeEvent(event)

    def __init__(
        self,
        base_dir: Path | None = None,
        config_path: Path | None = None,
    ) -> None:
        super().__init__()
        self.base_dir = Path(base_dir or r"D:\Sandbox\pm_sosanhbom")
        self.config_path = config_path

        self.setWindowTitle("Kyocera SSBOM - Chương Trình So Sánh BOM Tự Động")
        self.resize(1200, 800)
        self.setMinimumSize(950, 650)

        self._setup_logging_dock()
        self._init_ui()
        self._setup_menus_and_toolbars()
        self._setup_status_bar()

        logger.info("SSBOM Desktop Application initialized.")

    def _init_ui(self) -> None:
        """Create central tab widget hosting Leader and Member views."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # Tab Widget for Role-based Navigation
        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("Calibri", 10, QFont.Weight.Bold))

        # View 1: Leader Workspace
        self.leader_view = LeaderWorkspaceView(parent=self, base_dir=self.base_dir)
        self.tabs.addTab(self.leader_view, "👔 Trưởng Nhóm")

        # View 2: Member Workspace
        self.member_view = MemberWorkspaceView(parent=self, base_dir=self.base_dir)
        self.tabs.addTab(self.member_view, "👷 Thành Viên Công Đoạn")

        # Ensure Leader Workspace is active by default
        self.tabs.setCurrentIndex(0)

        main_layout.addWidget(self.tabs)

        # Wire cross-view events
        self.member_view.submission_completed.connect(self._on_member_submitted)
        self.leader_view.batch_finished.connect(self._on_batch_finished)

    def _setup_menus_and_toolbars(self) -> None:
        """Construct top MenuBar and Quick ToolBar."""
        menu_bar = self.menuBar()

        # Menu: Hệ thống (File)
        file_menu = menu_bar.addMenu("Hệ thống")

        # Action: Tải BOM từ Teamcenter PLM / SAP R3
        action_download_plm = QAction("📥 Tải BOM Tự Động (PLM / SAP R3)...", self)
        action_download_plm.setShortcut("Ctrl+D")
        action_download_plm.triggered.connect(self.leader_view._open_plm_download_dialog)
        file_menu.addAction(action_download_plm)

        action_settings = QAction("⚙ Cấu hình kết nối hệ thống...", self)
        action_settings.setShortcut("Ctrl+,")
        action_settings.triggered.connect(self.open_settings_dialog)
        file_menu.addAction(action_settings)

        file_menu.addSeparator()
        action_exit = QAction("Đóng ứng dụng", self)
        action_exit.setShortcut("Alt+F4")
        action_exit.triggered.connect(self.close)
        file_menu.addAction(action_exit)

        # Menu: Không gian làm việc (Workspaces)
        mode_menu = menu_bar.addMenu("Không gian làm việc")
        action_goto_leader = QAction("Chuyển sang Quản lý Trưởng nhóm", self)
        action_goto_leader.triggered.connect(lambda: self.switch_view(0))
        mode_menu.addAction(action_goto_leader)

        action_goto_member = QAction("Chuyển sang Nhập liệu Thành viên", self)
        action_goto_member.triggered.connect(lambda: self.switch_view(1))
        mode_menu.addAction(action_goto_member)

        # Menu: Trợ giúp (Help)
        help_menu = menu_bar.addMenu("Trợ giúp")
        action_about = QAction("Về chương trình...", self)
        action_about.triggered.connect(self._show_about_dialog)
        help_menu.addAction(action_about)

        # Toolbar
        toolbar = QToolBar("Thanh công cụ")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.addAction(action_download_plm)
        toolbar.addSeparator()
        toolbar.addAction(action_goto_leader)
        toolbar.addAction(action_goto_member)
        toolbar.addSeparator()
        toolbar.addAction(action_settings)

    def _setup_status_bar(self) -> None:
        """Create informative status bar with author credit and connection pills."""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.lbl_status_msg = QLabel("Sẵn sàng làm việc")
        self.status_bar.addWidget(self.lbl_status_msg, 1)

        # Author Credit
        self.lbl_author = QLabel("Người viết: Bùi Đức Vinh - Phòng PTHT Chế Tạo")
        self.lbl_author.setStyleSheet("color: #0078D4; font-weight: bold; margin-right: 16px;")
        self.status_bar.addPermanentWidget(self.lbl_author)

        # Connection indicators
        self.lbl_tc_indicator = QLabel("● TC24: Kết nối Web")
        self.lbl_tc_indicator.setStyleSheet("color: #10B981; font-weight: bold; margin-right: 10px;")
        self.status_bar.addPermanentWidget(self.lbl_tc_indicator)

        self.lbl_sap_indicator = QLabel("● SAP R3: Sẵn sàng")
        self.lbl_sap_indicator.setStyleSheet("color: #10B981; font-weight: bold; margin-right: 10px;")
        self.status_bar.addPermanentWidget(self.lbl_sap_indicator)

        self.lbl_version = QLabel("v2.0")
        self.lbl_version.setStyleSheet("color: #6c757d; margin-right: 8px;")
        self.status_bar.addPermanentWidget(self.lbl_version)

    def _setup_logging_dock(self) -> None:
        """Setup dockable real-time logging console at bottom."""
        self.log_dock = QDockWidget("Nhật ký hoạt động", self)
        self.log_dock.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea | Qt.DockWidgetArea.TopDockWidgetArea)

        dock_content = QWidget()
        dock_layout = QVBoxLayout(dock_content)
        dock_layout.setContentsMargins(4, 4, 4, 4)

        self.log_text_edit = QPlainTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.log_text_edit.setMaximumBlockCount(1000)
        self.log_text_edit.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: Consolas, monospace; font-size: 11px;")
        dock_layout.addWidget(self.log_text_edit)

        self.log_dock.setWidget(dock_content)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, self.log_dock)

        # Attach handler to root logger
        self.log_handler = QLogHandler()
        self.log_handler.new_record.connect(self.append_log)
        logging.getLogger().addHandler(self.log_handler)

    @pyqtSlot(str)
    def append_log(self, text: str) -> None:
        """Append log message to the console widget."""
        self.log_text_edit.appendPlainText(text)

    def switch_view(self, tab_index: int) -> None:
        """Switch active role workspace."""
        if 0 <= tab_index < self.tabs.count():
            self.tabs.setCurrentIndex(tab_index)

    def open_settings_dialog(self) -> None:
        """Display configuration dialog."""
        dlg = SettingsDialog(parent=self, config_path=self.config_path)
        dlg.exec()

    def _on_member_submitted(self, info: dict[str, Any]) -> None:
        """Handle event when a member submits their CTTT data."""
        unit = info.get("sub_unit", "")
        author = info.get("author", "")
        count = info.get("item_count", 0)
        self.lbl_status_msg.setText(f"Đã nhận nộp bài từ công đoạn '{unit}' ({author}, {count} linh kiện)")
        logger.info("Member submitted CTTT data: %s", info)

        # Refresh leader's submission board
        self.leader_view.scan_member_submissions()

    def _on_batch_finished(self, result: Any) -> None:
        """Handle event when leader finishes batch reconciliation."""
        self.lbl_status_msg.setText(f"Đối soát hàng loạt hoàn thành! Trạng thái: {result.overall_status}")

    def _show_about_dialog(self) -> None:
        """Show about software information."""
        QMessageBox.about(
            self,
            "Giới thiệu Phần mềm",
            "<h3>Hệ Thống Đối Soát BOM Tự Động (SSBOM v2.0)</h3>"
            "<p>Hiện đại hóa toàn diện quy trình so sánh BOM CTTT vs PLM vs R3.</p>"
            "<ul>"
            "<li><strong>Người viết:</strong> Bùi Đức Vinh - Phòng PTHT Chế Tạo</li>"
            "<li><strong>Đơn vị:</strong> Kyocera Document Solutions Vietnam</li>"
            "<li><strong>Công nghệ:</strong> Python 3.13, PyQt6, OpenXML, Selenium TC24, SAP GUI Scripting</li>"
            "<li><strong>Kiến trúc:</strong> Modular Adapter Pattern, Multi-level BOM Tree O(N) Resolver</li>"
            "</ul>"
            "<p>Bản quyền © 2026 Kyocera Document Solutions Vietnam Co., Ltd.</p>",
        )


def main() -> None:
    """Application entrypoint."""
    logging.basicConfig(level=logging.INFO, format="[%(asctime)s] [%(levelname)s] %(message)s")
    if "--health-check" in sys.argv:
        print("SSBOM Health Check: OK")
        sys.exit(0)
    app = QApplication(sys.argv)
    window = SSBOMMainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
