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

from PyQt6.QtCore import QObject, QThread, Qt, QTimer, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QApplication,
    QDockWidget,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStatusBar,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from src.core.path_utils import get_app_base_dir, get_asset_path, safe_exists, safe_is_file
from src.gui.leader_view import LeaderWorkspaceView
from src.gui.member_view import MemberWorkspaceView
from src.gui.styles import (
    compute_smart_geometry,
    ensure_window_fits_screen,
    get_theme_manager,
)
from src.gui.update_dialog import UpdateDialog
from src.services.app_updates import AppUpdateManager, get_build_release_notes, get_build_version
from src.services.update_delivery import UpdateDeliveryService

logger = logging.getLogger(__name__)

_global_terminating_threads: list[QThread] = []


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


class SilentUpdateCheckWorker(QThread):
    """Background QThread performing non-blocking update discovery across LAN sources."""

    candidate_found = pyqtSignal(object, str, object, object)

    def __init__(self, base_dir: Path, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.base_dir = base_dir

    def run(self) -> None:
        try:
            update_mgr = AppUpdateManager(install_root=self.base_dir)
            current_ver = update_mgr.get_active_version()
            build_ver = get_build_version()
            if build_ver and build_ver != "unknown":
                current_ver = build_ver

            logger.info("Kiểm tra cập nhật ngầm: phiên bản đang chạy v%s...", current_ver)
            delivery_worker = UpdateDeliveryService(base_dir=self.base_dir, current_version=current_ver)
            cfg = delivery_worker.load_sources_configuration()
            if not cfg.get("startup_check", True):
                logger.info("Tính năng kiểm tra cập nhật tự động đã bị tắt trong cấu hình.")
                return

            candidate = delivery_worker.check_for_updates()
            if candidate:
                logger.info(
                    "Đã tìm thấy bản cập nhật mới v%s (gói %s). Chuẩn bị hiển thị gợi ý cập nhật.",
                    candidate.version,
                    candidate.package_name,
                )
                self.candidate_found.emit(candidate, current_ver, delivery_worker, update_mgr)
            else:
                logger.info("Kiểm tra cập nhật ngầm: hệ thống đang ở phiên bản mới nhất (v%s).", current_ver)
        except Exception as exc:
            logger.warning("Silent background update check failed: %s", exc)


class SSBOMMainWindow(QMainWindow):
    """Main application window for Kyocera automated BOM comparison."""

    def closeEvent(self, event: Any) -> None:
        """Ensure logging handler is detached cleanly and background workers stopped on window close."""
        timer = getattr(self, "_periodic_update_timer", None)
        if timer is not None:
            try:
                timer.stop()
            except Exception:
                pass

        worker = getattr(self, "_update_check_worker", None)
        self._update_check_worker = None
        if worker is not None:
            try:
                worker.candidate_found.disconnect()
            except Exception:
                pass
            is_running = False
            try:
                is_running = worker.isRunning()
            except Exception:
                is_running = False

            if is_running:
                try:
                    worker.quit()
                    worker.wait(1000)
                except Exception:
                    pass
                try:
                    if worker.isRunning():
                        _global_terminating_threads.append(worker)
                        worker.finished.connect(
                            lambda w=worker: _global_terminating_threads.remove(w)
                            if w in _global_terminating_threads
                            else None
                        )
                except Exception:
                    pass

        if hasattr(self, "log_handler") and self.log_handler in logging.getLogger().handlers:
            logging.getLogger().removeHandler(self.log_handler)
        super().closeEvent(event)

    def __init__(
        self,
        base_dir: Path | None = None,
        config_path: Path | None = None,
    ) -> None:
        super().__init__()
        self.base_dir = Path(base_dir) if base_dir else get_app_base_dir()
        self.config_path = config_path
        self._update_check_worker: SilentUpdateCheckWorker | None = None
        self._active_update_dialog: UpdateDialog | None = None
        self._periodic_update_timer: QTimer | None = None

        # Initialize global crash protection and diagnostic logging
        from src.core.crash_handler import setup_crash_protection
        setup_crash_protection(base_dir=self.base_dir)

        self.setWindowTitle("Kyocera SSBOM - Chương Trình So Sánh BOM Tự Động")
        x, y, w, h, is_compact = compute_smart_geometry(
            widget_or_parent=self,
            ideal_w=1280,
            ideal_h=800,
            min_w=880,
            min_h=480,
            max_w_ratio=0.96,
            max_h_ratio=0.92,
            margin_bottom=20,
        )
        self.setMinimumSize(min(880, w), min(480, h))
        self.setGeometry(x, y, w, h)

        self.theme_mgr = get_theme_manager()
        self.theme_mgr.apply_theme_to_app()
        app_icon_path = get_asset_path("icon.png", self.base_dir)
        if app_icon_path:
            from PyQt6.QtGui import QIcon
            self.setWindowIcon(QIcon(str(app_icon_path)))
        else:
            self.setWindowIcon(self.theme_mgr.get_styled_icon("layers"))

        self._init_ui()
        self._setup_logging_dock()
        self._setup_menus_and_toolbars()
        self._setup_status_bar()
        # Schedule background update check if enabled in sources config (skipped during unit testing)
        import sys
        if "pytest" not in sys.modules:
            self._attach_update_log()
            QTimer.singleShot(1500, self._show_whats_new_once)
            QTimer.singleShot(2500, self._check_updates_silent)

            # Periodic background update check every 60 minutes
            self._periodic_update_timer = QTimer(self)
            self._periodic_update_timer.setInterval(60 * 60 * 1000)
            self._periodic_update_timer.timeout.connect(self._check_updates_silent)
            self._periodic_update_timer.start()

        logger.info("Hệ thống So Sánh BOM (SSBOM) đã khởi tạo thành công.")

    def showEvent(self, event: Any) -> None:
        """Ensure main window never clips beneath the Windows taskbar on initial show."""
        super().showEvent(event)
        ensure_window_fits_screen(self, margin_bottom=20)

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

        self.action_bom_filter = QAction(" Quản lý Bộ lọc BOM PLM", self)
        self.action_bom_filter.setIcon(self.theme_mgr.get_styled_icon("filter"))
        self.action_bom_filter.setShortcut("Ctrl+B")
        self.action_bom_filter.setToolTip("Quản lý và cấu hình các quy tắc lọc BOM PLM (Sheet BolocBom) [Ctrl+B]")
        self.action_bom_filter.triggered.connect(self.open_bom_filter_dialog)
        file_menu.addAction(self.action_bom_filter)

        self.action_standalone_filter = QAction(" Áp dụng bộ lọc cho file có sẵn...", self)
        self.action_standalone_filter.setIcon(self.theme_mgr.get_styled_icon("filter"))
        self.action_standalone_filter.setShortcut("Ctrl+Shift+B")
        self.action_standalone_filter.setToolTip("Lọc lại file Excel PLM trên máy tính theo bộ quy tắc mà không cần tải lại từ Teamcenter [Ctrl+Shift+B]")
        self.action_standalone_filter.triggered.connect(self.open_standalone_bom_filter_dialog)
        file_menu.addAction(self.action_standalone_filter)

        self.action_bom_update = QAction(" Cập nhật BOM mới (PLM & R3)...", self)
        self.action_bom_update.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        self.action_bom_update.setShortcut("Ctrl+U")
        self.action_bom_update.setToolTip("Cập nhật file PLM hoặc SAP R3 mới vào workbook BOM (tương đương Nút 7 & 8 VBA) [Ctrl+U]")
        self.action_bom_update.triggered.connect(lambda: self.open_bom_update_dialog())
        file_menu.addAction(self.action_bom_update)

        self.action_theme = QAction(" Đổi Giao Diện Sáng/Tối", self)
        self.action_theme.setShortcut("Ctrl+T")
        self.action_theme.triggered.connect(self._toggle_theme)
        self._update_theme_action(self.theme_mgr.current_theme)
        file_menu.addAction(self.action_theme)

        self.action_toggle_log = QAction(" Nhật ký hoạt động", self)
        self.action_toggle_log.setIcon(self.theme_mgr.get_styled_icon("clipboard"))
        self.action_toggle_log.setShortcut("Ctrl+L")
        self.action_toggle_log.setCheckable(True)
        self.action_toggle_log.setChecked(False)
        self.action_toggle_log.setToolTip("Ẩn / Hiện cửa sổ Nhật ký hoạt động [Ctrl+L]")
        self.action_toggle_log.triggered.connect(self._toggle_logging_dock)
        self.log_dock.visibilityChanged.connect(self._on_log_dock_visibility_changed)
        file_menu.addAction(self.action_toggle_log)

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

        self.action_open_logs = QAction(" Mở thư mục nhật ký lỗi (Crash Logs)...", self)
        self.action_open_logs.setIcon(self.theme_mgr.get_styled_icon("folder"))
        self.action_open_logs.setToolTip("Mở thư mục chứa các tệp nhật ký sự cố và hoạt động hệ thống")
        self.action_open_logs.triggered.connect(self._open_logs_folder)
        help_menu.addAction(self.action_open_logs)

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
        self.toolbar.addAction(self.action_bom_filter)
        self.toolbar.addAction(self.action_standalone_filter)
        self.toolbar.addAction(self.action_bom_update)
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
        self.lbl_status_msg.setStyleSheet("padding-left: 8px; font-weight: 500;")
        self.status_bar.addWidget(self.lbl_status_msg, 1)

        self.btn_toggle_log = QPushButton("📋 Nhật ký (Ctrl+L)")
        self.btn_toggle_log.setToolTip("Bật / Tắt cửa sổ Nhật ký hoạt động [Ctrl+L]")
        self.btn_toggle_log.setStyleSheet(
            "QPushButton { padding: 2px 8px; font-size: 11px; font-weight: 600; border-radius: 3px; }"
        )
        self.btn_toggle_log.clicked.connect(self._toggle_logging_dock)
        self.status_bar.addWidget(self.btn_toggle_log)

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
            active_ver = "unknown"
        build_ver = get_build_version()

        self.lbl_version = QLabel(f"v{build_ver}")
        if build_ver != active_ver and build_ver != "unknown":
            self.lbl_version.setText(f"v{build_ver} ⚠")
            self.lbl_version.setToolTip(
                f"Đang chạy bản v{build_ver} nhưng thư mục cài đặt trỏ tới v{active_ver}.\n"
                "Hãy mở chương trình bằng SSBOM_Launcher hoặc shortcut Start Menu."
            )
            self.lbl_version.setStyleSheet("margin-right: 8px; font-weight: bold; color: #F59E0B;")
        else:
            self.lbl_version.setToolTip(f"Phiên bản đang chạy: v{build_ver}")
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
        self.log_dock.hide()  # Hidden by default to maximize table vertical space

        # Attach handler to root logger
        self.log_handler = QLogHandler()
        self.log_handler.new_record.connect(self.append_log)
        logging.getLogger().addHandler(self.log_handler)

    def _toggle_logging_dock(self) -> None:
        """Toggle visibility of the logging dock widget."""
        visible = not self.log_dock.isVisible()
        self.log_dock.setVisible(visible)
        if hasattr(self, "action_toggle_log"):
            self.action_toggle_log.setChecked(visible)

    def _on_log_dock_visibility_changed(self, visible: bool) -> None:
        """Sync action checked state with dock widget visibility."""
        if hasattr(self, "action_toggle_log"):
            self.action_toggle_log.setChecked(visible)

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
        from src.gui.settings_dialog import SettingsDialog
        dlg = SettingsDialog(parent=self, config_path=self.config_path)
        dlg.exec()

    def open_bom_filter_dialog(self) -> None:
        """Display BOM filter rules configuration dialog."""
        from src.gui.bom_filter_dialog import BOMFilterConfigDialog
        current_model = None
        if hasattr(self, "leader_view") and hasattr(self.leader_view, "state"):
            current_model = getattr(self.leader_view.state, "model_name", None)
        dlg = BOMFilterConfigDialog(initial_model=current_model, parent=self)
        dlg.exec()

    def open_standalone_bom_filter_dialog(self) -> None:
        """Display standalone BOM filtering dialog for existing files."""
        from src.gui.standalone_bom_filter_dialog import StandaloneBOMFilterDialog
        current_model = None
        if hasattr(self, "leader_view") and hasattr(self.leader_view, "state"):
            current_model = getattr(self.leader_view.state, "model_name", None)
        dlg = StandaloneBOMFilterDialog(initial_model=current_model, parent=self)
        dlg.exec()

    def open_bom_update_dialog(
        self,
        initial_dir: Optional[Path] = None,
        initial_bom_file: Optional[Path] = None,
    ) -> None:
        """Display standalone BOM update dialog for new PLM & R3 files (legacy VBA Buttons 7 & 8)."""
        from src.gui.bom_update_dialog import BOMUpdateDialog

        if initial_dir is None and hasattr(self, "leader_view"):
            initial_dir = getattr(self.leader_view, "current_machine_dir", None)
        dlg = BOMUpdateDialog(parent=self, initial_dir=initial_dir, initial_bom_file=initial_bom_file)
        dlg.exec()

    def _open_logs_folder(self) -> None:
        """Open diagnostic crash logs directory."""
        from src.core.crash_handler import open_logs_folder
        if not open_logs_folder():
            QMessageBox.information(
                self,
                "Nhật ký lỗi",
                f"Thư mục nhật ký nằm tại:\n{self.base_dir / 'logs'}",
            )

    def _refresh_theme_icons(self) -> None:
        """Re-apply dynamic styled SVG icons across all actions, tabs, and window."""
        if hasattr(self, "action_download_plm"):
            self.action_download_plm.setIcon(self.theme_mgr.get_styled_icon("download"))
        if hasattr(self, "action_bom_filter"):
            self.action_bom_filter.setIcon(self.theme_mgr.get_styled_icon("filter"))
        if hasattr(self, "action_standalone_filter"):
            self.action_standalone_filter.setIcon(self.theme_mgr.get_styled_icon("filter"))
        if hasattr(self, "action_bom_update"):
            self.action_bom_update.setIcon(self.theme_mgr.get_styled_icon("refresh"))
        if hasattr(self, "action_open_logs"):
            self.action_open_logs.setIcon(self.theme_mgr.get_styled_icon("folder"))
        if hasattr(self, "action_goto_leader"):
            self.action_goto_leader.setIcon(self.theme_mgr.get_styled_icon("users"))
        if hasattr(self, "action_goto_member"):
            self.action_goto_member.setIcon(self.theme_mgr.get_styled_icon("user-check"))
        if hasattr(self, "action_settings"):
            self.action_settings.setIcon(self.theme_mgr.get_styled_icon("settings"))
        if hasattr(self, "action_toggle_log"):
            self.action_toggle_log.setIcon(self.theme_mgr.get_styled_icon("clipboard"))
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
        self.lbl_status_msg.setText(f"Đã nhận file CTTT từ công đoạn '{unit}' ({author}, {count} linh kiện)")
        logger.info("Thành viên công đoạn đã hoàn tất cập nhật file CTTT: %s (Công đoạn: %s, %d linh kiện)", author, unit, count)

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
            build_ver = get_build_version()
            if build_ver and build_ver != "unknown":
                current_ver = build_ver
            delivery = UpdateDeliveryService(base_dir=self.base_dir, current_version=current_ver)
            candidate = delivery.check_for_updates()

            if candidate:
                self._show_update_candidate(candidate, current_ver, delivery, update_mgr)
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

    def _show_update_candidate(self, candidate: Any, current_ver: str, delivery: Any, update_mgr: Any) -> None:
        """Present update dialog on UI thread when candidate is found."""
        if getattr(self, "_active_update_dialog", None) is not None:
            if self._active_update_dialog.isVisible():
                return
        try:
            self._active_update_dialog = UpdateDialog(
                parent=self,
                candidate=candidate,
                current_version=current_ver,
                delivery_service=delivery,
                update_manager=update_mgr,
            )
            self._active_update_dialog.exec()
        except Exception as exc:
            logger.debug("Failed to display update dialog: %s", exc)
        finally:
            self._active_update_dialog = None

    def _on_silent_worker_finished(self, worker: Any) -> None:
        """Slot invoked when silent update worker finishes to clear reference cleanly."""
        if getattr(self, "_update_check_worker", None) is worker:
            self._update_check_worker = None

    def _check_updates_silent(self) -> None:
        """Silent update check triggered on application startup or periodically in background QThread."""
        worker = getattr(self, "_update_check_worker", None)
        if worker is not None:
            is_running = False
            try:
                is_running = worker.isRunning()
            except Exception:
                is_running = False

            if is_running:
                return
            try:
                worker.deleteLater()
            except Exception:
                pass
            self._update_check_worker = None

        worker = SilentUpdateCheckWorker(base_dir=self.base_dir, parent=None)
        worker.candidate_found.connect(self._show_update_candidate)
        worker.finished.connect(lambda w=worker: self._on_silent_worker_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._update_check_worker = worker
        worker.start()

    def _attach_update_log(self) -> None:
        """Write updater activity (check, download, install, failures) to logs/update.log."""
        try:
            log_dir = self.base_dir / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            handler = logging.FileHandler(log_dir / "update.log", encoding="utf-8")
            handler.setLevel(logging.INFO)
            handler.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s"))
            for name in ("app_updates", "update_delivery", "update_dialog", "update_security", __name__):
                lg = logging.getLogger(name)
                lg.setLevel(logging.INFO)
                lg.addHandler(handler)
            logging.getLogger(__name__).info(
                "Khởi động SSBOM build v%s (base_dir=%s)", get_build_version(), self.base_dir
            )
        except Exception as exc:
            logger.warning("Cannot attach update log: %s", exc)

    def _show_whats_new_once(self) -> None:
        """Show bundled release notes the first time a new build is launched."""
        build_ver = get_build_version()
        notes = get_build_release_notes()
        if build_ver == "unknown" or not notes:
            return
        marker = self.base_dir / "logs" / "last_seen_version.txt"
        try:
            if marker.is_file() and marker.read_text(encoding="utf-8").strip() == build_ver:
                return
            marker.parent.mkdir(parents=True, exist_ok=True)
            marker.write_text(build_ver, encoding="utf-8")
        except Exception as exc:
            logger.warning("Cannot persist last seen version: %s", exc)
        QMessageBox.information(self, f"Có gì mới trong SSBOM v{build_ver}", notes)

    def _show_about_dialog(self) -> None:
        """Show about software information."""
        active_ver = get_build_version()

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
    from src.core.crash_handler import setup_crash_protection
    setup_crash_protection()

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
