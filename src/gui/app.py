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

from PyQt6.QtCore import QObject, Qt, QTimer, pyqtSignal, pyqtSlot
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
from src.gui.styles import get_theme_manager
from src.gui.update_dialog import UpdateDialog
from src.services.app_updates import AppUpdateManager
from src.services.update_delivery import UpdateDeliveryService

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
        self.resize(1280, 860)
        self.setMinimumSize(1000, 680)

        self.theme_mgr = get_theme_manager()
        self.theme_mgr.apply_theme_to_app()
        self.setWindowIcon(self.theme_mgr.get_styled_icon("layers"))

        self._setup_logging_dock()
        self._init_ui()
        self._setup_menus_and_toolbars()
        self._setup_status_bar()
        # Schedule background update check if enabled in sources config
        QTimer.singleShot(2500, self._check_updates_silent)

        logger.info("Hệ thống So Sánh BOM (SSBOM) đã khởi tạo thành công.")

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
        self.tabs.addTab(self.leader_view, self.theme_mgr.get_styled_icon("users"), "Trưởng Nhóm")

        # View 2: Member Workspace
        self.member_view = MemberWorkspaceView(parent=self, base_dir=self.base_dir)
        self.tabs.addTab(self.member_view, self.theme_mgr.get_styled_icon("user-check"), "Thành Viên Công Đoạn")

        # Ensure Leader Workspace is active by default
        self.tabs.setCurrentIndex(0)

        main_layout.addWidget(self.tabs)

        # Wire cross-view events
        self.member_view.submission_completed.connect(self._on_member_submitted)
        self.leader_view.batch_finished.connect(self._on_batch_finished)

    def _setup_menus_and_toolbars(self) -> None:
        """Construct top MenuBar and Quick ToolBar with SVG icons."""
        menu_bar = self.menuBar()

        # Menu: Hệ thống (File)
        file_menu = menu_bar.addMenu("Hệ thống")

        # Action: Tải BOM từ Teamcenter PLM / SAP R3
        self.action_download_plm = QAction(" Tải BOM (PLM/SAP)", self)
        self.action_download_plm.setIcon(self.theme_mgr.get_styled_icon("download"))
        self.action_download_plm.setShortcut("Ctrl+D")
        self.action_download_plm.setToolTip("Tải BOM tự động từ PLM Teamcenter TC24 và SAP R3 [Ctrl+D]")
        self.action_download_plm.triggered.connect(self.leader_view._open_plm_download_dialog)
        file_menu.addAction(self.action_download_plm)

        self.action_settings = QAction(" Cấu hình hệ thống", self)
        self.action_settings.setIcon(self.theme_mgr.get_styled_icon("settings"))
        self.action_settings.setShortcut("Ctrl+,")
        self.action_settings.setToolTip("Cấu hình tài khoản, thư mục và giao diện [Ctrl+,]")
        self.action_settings.triggered.connect(self.open_settings_dialog)
        file_menu.addAction(self.action_settings)

        self.action_theme = QAction(" Đổi Giao Diện Sáng/Tối", self)
        self.action_theme.setShortcut("Ctrl+T")
        self.action_theme.triggered.connect(self._toggle_theme)
        self._update_theme_action(self.theme_mgr.current_theme)
        file_menu.addAction(self.action_theme)

        file_menu.addSeparator()
        action_exit = QAction("Đóng ứng dụng", self)
        action_exit.setShortcut("Alt+F4")
        action_exit.triggered.connect(self.close)
        file_menu.addAction(action_exit)

        # Menu: Không gian làm việc (Workspaces)
        mode_menu = menu_bar.addMenu("Không gian làm việc")
        self.action_goto_leader = QAction(" Quản lý Trưởng nhóm", self)
        self.action_goto_leader.setIcon(self.theme_mgr.get_styled_icon("users"))
        self.action_goto_leader.triggered.connect(lambda: self.switch_view(0))
        mode_menu.addAction(self.action_goto_leader)

        self.action_goto_member = QAction(" Nhập liệu Thành viên", self)
        self.action_goto_member.setIcon(self.theme_mgr.get_styled_icon("user-check"))
        self.action_goto_member.triggered.connect(lambda: self.switch_view(1))
        mode_menu.addAction(self.action_goto_member)

        # Menu: Trợ giúp (Help)
        help_menu = menu_bar.addMenu("Trợ giúp")
        self.action_check_updates = QAction(" Kiểm tra cập nhật...", self)
        self.action_check_updates.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        self.action_check_updates.triggered.connect(lambda: self.check_for_updates(interactive=True))
        help_menu.addAction(self.action_check_updates)
        help_menu.addSeparator()
        self.action_about = QAction(" Về chương trình...", self)
        self.action_about.setIcon(self.theme_mgr.get_styled_icon("info"))
        self.action_about.triggered.connect(self._show_about_dialog)
        help_menu.addAction(self.action_about)

        # Toolbar with high contrast icons and visible text labels
        self.toolbar = QToolBar("Thanh công cụ")
        self.toolbar.setMovable(False)
        self.toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.addToolBar(self.toolbar)

        self.toolbar.addAction(self.action_download_plm)
        self.toolbar.addSeparator()
        self.toolbar.addAction(self.action_goto_leader)
        self.toolbar.addAction(self.action_goto_member)
        self.toolbar.addSeparator()
        self.toolbar.addAction(self.action_theme)
        self.toolbar.addAction(self.action_settings)

        # Listen to theme changed signal to re-render all icons dynamically
        self.theme_mgr.theme_changed.connect(self._on_theme_changed)

    def _setup_status_bar(self) -> None:
        """Create informative status bar with author credit and connection pills."""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.lbl_status_msg = QLabel("Sẵn sàng làm việc")
        self.status_bar.addWidget(self.lbl_status_msg, 1)

        # Author Credit
        self.lbl_author = QLabel("Người viết: Bùi Đức Vinh - Phòng PTHT Chế Tạo")
        self.lbl_author.setStyleSheet("font-weight: bold; margin-right: 16px;")
        self.status_bar.addPermanentWidget(self.lbl_author)

        # Connection indicators
        self.lbl_tc_indicator = QLabel("● TC24: Kết nối Web")
        self.lbl_tc_indicator.setStyleSheet("color: #10B981; font-weight: bold; margin-right: 10px;")
        self.status_bar.addPermanentWidget(self.lbl_tc_indicator)

        self.lbl_sap_indicator = QLabel("● SAP R3: Sẵn sàng")
        self.lbl_sap_indicator.setStyleSheet("color: #10B981; font-weight: bold; margin-right: 10px;")
        self.status_bar.addPermanentWidget(self.lbl_sap_indicator)

        try:
            active_ver = AppUpdateManager(self.base_dir).get_active_version()
        except Exception:
            active_ver = "1.0.0"

        self.lbl_version = QLabel(f"v{active_ver}")
        self.lbl_version.setStyleSheet("margin-right: 8px; font-weight: bold;")
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
        self.log_text_edit.setFont(QFont("Consolas", 10))
        dock_layout.addWidget(self.log_text_edit)

        self.log_dock.setWidget(dock_content)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, self.log_dock)
        self.resizeDocks([self.log_dock], [90], Qt.Orientation.Vertical)

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

    def _refresh_theme_icons(self) -> None:
        """Re-apply dynamic styled SVG icons across all actions, tabs, and window."""
        if hasattr(self, "action_download_plm"):
            self.action_download_plm.setIcon(self.theme_mgr.get_styled_icon("download"))
        if hasattr(self, "action_goto_leader"):
            self.action_goto_leader.setIcon(self.theme_mgr.get_styled_icon("users"))
        if hasattr(self, "action_goto_member"):
            self.action_goto_member.setIcon(self.theme_mgr.get_styled_icon("user-check"))
        if hasattr(self, "action_settings"):
            self.action_settings.setIcon(self.theme_mgr.get_styled_icon("settings"))
        if hasattr(self, "action_check_updates"):
            self.action_check_updates.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        if hasattr(self, "action_about"):
            self.action_about.setIcon(self.theme_mgr.get_styled_icon("info"))
        if hasattr(self, "tabs") and self.tabs.count() >= 2:
            self.tabs.setTabIcon(0, self.theme_mgr.get_styled_icon("users"))
            self.tabs.setTabIcon(1, self.theme_mgr.get_styled_icon("user-check"))
        self.setWindowIcon(self.theme_mgr.get_styled_icon("layers"))
        if hasattr(self, "action_theme"):
            self._update_theme_action(self.theme_mgr.current_theme)

    def _on_theme_changed(self, effective_theme: str) -> None:
        """Handle theme change signal, notifying child views and updating icons."""
        self._refresh_theme_icons()
        if hasattr(self, "leader_view") and hasattr(self.leader_view, "on_theme_changed"):
            self.leader_view.on_theme_changed(effective_theme)
        if hasattr(self, "member_view") and hasattr(self.member_view, "on_theme_changed"):
            self.member_view.on_theme_changed(effective_theme)

    def _update_theme_action(self, theme: str) -> None:
        """Update theme toggle action icon and tooltip based on active theme."""
        effective = getattr(self.theme_mgr, "effective_theme", theme)
        if effective == "dark":
            self.action_theme.setIcon(self.theme_mgr.get_styled_icon("sun"))
            self.action_theme.setText(" Giao diện: Tối [Ctrl+T]")
            self.action_theme.setToolTip("Đang ở giao diện Tối. Bấm để chuyển sang Giao diện Sáng (Light) [Ctrl+T]")
        else:
            self.action_theme.setIcon(self.theme_mgr.get_styled_icon("moon"))
            self.action_theme.setText(" Giao diện: Sáng [Ctrl+T]")
            self.action_theme.setToolTip("Đang ở giao diện Sáng. Bấm để chuyển sang Giao diện Tối (Dark) [Ctrl+T]")

    def _toggle_theme(self) -> None:
        """Toggle between Light (Slate Industrial) and Dark (Industrial Dark Mode)."""
        new_theme = "dark" if self.theme_mgr.current_theme == "light" else "light"
        self.theme_mgr.set_theme(new_theme)
        theme_vn = "Sáng" if new_theme == "light" else "Tối"
        self.lbl_status_msg.setText(f"Đã chuyển sang giao diện: {theme_vn.upper()}")
        logger.info("Đã chuyển sang giao diện: %s", theme_vn)

    def _on_member_submitted(self, info: dict[str, Any]) -> None:
        """Handle event when a member submits their CTTT data."""
        unit = info.get("sub_unit", "")
        author = info.get("author", "")
        count = info.get("item_count", 0)
        self.lbl_status_msg.setText(f"Đã nhận nộp bài từ công đoạn '{unit}' ({author}, {count} linh kiện)")
        logger.info("Thành viên công đoạn đã hoàn tất nộp bài: %s (Công đoạn: %s, %d linh kiện)", author, unit, count)

        # Refresh leader's submission board
        self.leader_view.scan_member_submissions()

    def _on_batch_finished(self, result: Any) -> None:
        """Handle event when leader finishes batch reconciliation."""
        self.lbl_status_msg.setText(f"Đối soát hàng loạt hoàn thành! Trạng thái: {result.overall_status}")

    def check_for_updates(self, interactive: bool = True) -> None:
        """Check for updates from LAN sources and present dialog if available."""
        try:
            update_mgr = AppUpdateManager(install_root=self.base_dir)
            current_ver = update_mgr.get_active_version()
            delivery = UpdateDeliveryService(base_dir=self.base_dir, current_version=current_ver)
            candidate = delivery.check_for_updates()

            if candidate:
                dlg = UpdateDialog(
                    parent=self,
                    candidate=candidate,
                    current_version=current_ver,
                    delivery_service=delivery,
                    update_manager=update_mgr,
                )
                dlg.exec()
            elif interactive:
                QMessageBox.information(
                    self,
                    "Cập nhật phần mềm",
                    f"Phiên bản hiện tại v{current_ver} đã là phiên bản mới nhất trên hệ thống mạng!",
                )
        except Exception as exc:
            logger.error("Error during update check: %s", exc)
            if interactive:
                QMessageBox.warning(
                    self,
                    "Kiểm tra cập nhật",
                    f"Không thể kiểm tra bản cập nhật từ mạng LAN:\n{exc}",
                )

    def _check_updates_silent(self) -> None:
        """Silent update check triggered on application startup if enabled in config."""
        try:
            delivery = UpdateDeliveryService(base_dir=self.base_dir)
            cfg = delivery.load_sources_configuration()
            if not cfg.get("startup_check", True):
                return

            self.check_for_updates(interactive=False)
        except Exception as exc:
            logger.debug("Silent startup update check failed: %s", exc)

    def _show_about_dialog(self) -> None:
        """Show about software information."""
        try:
            active_ver = AppUpdateManager(self.base_dir).get_active_version()
        except Exception:
            active_ver = "1.0.0"

        QMessageBox.about(
            self,
            "Giới thiệu Phần mềm",
            f"<h3>Hệ Thống Đối Soát BOM Tự Động (SSBOM v{active_ver})</h3>"
            "<p>Hiện đại hóa toàn diện quy trình so sánh BOM CTTT vs PLM vs R3.</p>"
            "<ul>"
            "<li><strong>Người viết:</strong> Bùi Đức Vinh - Phòng PTHT Chế Tạo</li>"
            "<li><strong>Đơn vị:</strong> Kyocera Document Solutions Vietnam</li>"
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
