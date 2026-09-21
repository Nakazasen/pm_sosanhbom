"""Settings and Configuration Dialog for SSBOM Application.

Allows configuration and validation of:
- Siemens Teamcenter TC14 Web Active Workspace credentials, URL, browser type, and timeouts.
- SAP R3 CS12 transaction logon parameters, plant, usage, and saplogon.exe paths.
- Project directories, master Fix Serial file path, and report output paths.
"""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.gui.styles import get_theme_manager

logger = logging.getLogger(__name__)

DEFAULT_CONFIG_PATH = Path("config/settings.json")
FALLBACK_CONFIG_PATH = Path.home() / ".ssbom" / "config.json"

DEFAULT_SETTINGS: dict[str, Any] = {
    "tc14": {
        "base_url": "http://tcmp3gwb:3000/",
        "username": "vn_pe03",
        "password": "vn_pe03",
        "browser": "edge",
        "headless": True,
        "timeout": 30,
    },
    "sap": {
        "system_id": "P1J(ERP60-AWS)-VN",
        "client": "100",
        "plant": "2200",
        "bom_usage": "pp01",
        "alternative": "01",
        "saplogon_path": r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
    },
    "paths": {
        "base_dir": r"D:\Sandbox\pm_sosanhbom",
        "fix_serial_path": r"D:\Sandbox\pm_sosanhbom\FIX_SERIAL_DLTOOL_VER010.xls",
        "reports_dir": r"D:\Sandbox\pm_sosanhbom\Reports",
        "default_model": "Virgo",
        "shared_db_path": r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db",
        "update_dir": r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\release_update",
    },
    "ui": {
        "theme": "light",
    },
}


class SettingsDialog(QDialog):
    """Modal dialog for modifying and testing application settings."""

    settings_saved = pyqtSignal(dict)

    def __init__(
        self,
        parent: QWidget | None = None,
        config_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Cấu hình Hệ thống (System Settings)")
        self.resize(580, 480)
        if config_path is not None:
            self.config_path = Path(config_path)
        elif DEFAULT_CONFIG_PATH.exists():
            self.config_path = DEFAULT_CONFIG_PATH
        elif FALLBACK_CONFIG_PATH.exists():
            self.config_path = FALLBACK_CONFIG_PATH
        else:
            self.config_path = DEFAULT_CONFIG_PATH

        self._current_settings: dict[str, Any] = json.loads(json.dumps(DEFAULT_SETTINGS))
        self._init_ui()
        self.load_from_file()

    def _init_ui(self) -> None:
        """Create tabbed configuration interface."""
        theme_mgr = get_theme_manager()
        self.setWindowIcon(theme_mgr.get_styled_icon("settings"))

        main_layout = QVBoxLayout(self)

        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Tab 1: Teamcenter TC24
        tc_widget = QWidget()
        tc_layout = QVBoxLayout(tc_widget)
        tc_group = QGroupBox("Cổng Web Siemens Teamcenter Active Workspace (TC24)")
        tc_form = QFormLayout(tc_group)

        self.tc_url_edit = QLineEdit()
        self.tc_user_edit = QLineEdit()
        self.tc_pass_edit = QLineEdit()
        self.tc_pass_edit.setEchoMode(QLineEdit.EchoMode.Password)

        self.tc_browser_combo = QComboBox()
        self.tc_browser_combo.addItems(["edge", "chrome"])

        self.tc_headless_check = QCheckBox("Chạy ẩn danh (Headless - khuyến nghị)")
        self.tc_headless_check.setChecked(True)

        self.tc_timeout_spin = QSpinBox()
        self.tc_timeout_spin.setRange(5, 120)
        self.tc_timeout_spin.setValue(30)
        self.tc_timeout_spin.setSuffix(" giây")

        tc_form.addRow("URL Cổng TC24:", self.tc_url_edit)
        tc_form.addRow("Tài khoản (Username):", self.tc_user_edit)
        tc_form.addRow("Mật khẩu (Password):", self.tc_pass_edit)
        tc_form.addRow("Trình duyệt (Browser):", self.tc_browser_combo)
        tc_form.addRow("Chế độ trình duyệt:", self.tc_headless_check)
        tc_form.addRow("Thời gian chờ (Timeout):", self.tc_timeout_spin)

        tc_btn_layout = QHBoxLayout()
        self.btn_test_tc = QPushButton(" Kiểm tra kết nối TC24")
        self.btn_test_tc.setIcon(theme_mgr.get_styled_icon("refresh"))
        self.btn_test_tc.clicked.connect(self._test_tc_connection)
        tc_btn_layout.addWidget(self.btn_test_tc)
        tc_btn_layout.addStretch()

        tc_layout.addWidget(tc_group)
        tc_layout.addLayout(tc_btn_layout)
        tc_layout.addStretch()
        self.tab_widget.addTab(tc_widget, theme_mgr.get_styled_icon("layers"), "Teamcenter TC24")

        # Tab 2: SAP R3
        sap_widget = QWidget()
        sap_layout = QVBoxLayout(sap_widget)
        sap_group = QGroupBox("Thông số kết nối SAP R3 (CS12 Multilevel BOM)")
        sap_form = QFormLayout(sap_group)

        self.sap_system_edit = QLineEdit()
        self.sap_client_edit = QLineEdit()
        self.sap_plant_edit = QLineEdit()
        self.sap_usage_edit = QLineEdit()
        self.sap_alt_edit = QLineEdit()

        sap_path_layout = QHBoxLayout()
        self.sap_path_edit = QLineEdit()
        btn_browse_sap = QPushButton(" Chọn...")
        btn_browse_sap.setIcon(theme_mgr.get_styled_icon("folder"))
        btn_browse_sap.clicked.connect(self._browse_saplogon)
        sap_path_layout.addWidget(self.sap_path_edit)
        sap_path_layout.addWidget(btn_browse_sap)

        sap_form.addRow("System Description (P1J):", self.sap_system_edit)
        sap_form.addRow("SAP Client:", self.sap_client_edit)
        sap_form.addRow("Plant (Mã nhà máy):", self.sap_plant_edit)
        sap_form.addRow("BOM Usage:", self.sap_usage_edit)
        sap_form.addRow("Alternative:", self.sap_alt_edit)
        sap_form.addRow("Đường dẫn saplogon.exe:", sap_path_layout)

        sap_btn_layout = QHBoxLayout()
        self.btn_test_sap = QPushButton(" Kiểm tra kết nối SAP GUI")
        self.btn_test_sap.setIcon(theme_mgr.get_styled_icon("refresh"))
        self.btn_test_sap.clicked.connect(self._test_sap_connection)
        sap_btn_layout.addWidget(self.btn_test_sap)
        sap_btn_layout.addStretch()

        sap_layout.addWidget(sap_group)
        sap_layout.addLayout(sap_btn_layout)
        sap_layout.addStretch()
        self.tab_widget.addTab(sap_widget, theme_mgr.get_styled_icon("download"), "SAP R3")

        # Tab 3: Directories and Paths
        path_widget = QWidget()
        path_layout = QVBoxLayout(path_widget)
        path_group = QGroupBox("Đường dẫn thư mục dự án và tệp mẫu")
        path_form = QFormLayout(path_group)

        # Base Dir
        base_layout = QHBoxLayout()
        self.base_dir_edit = QLineEdit()
        btn_browse_base = QPushButton(" Chọn...")
        btn_browse_base.setIcon(theme_mgr.get_styled_icon("folder"))
        btn_browse_base.clicked.connect(lambda: self._browse_dir(self.base_dir_edit, "Chọn thư mục gốc dự án"))
        base_layout.addWidget(self.base_dir_edit)
        base_layout.addWidget(btn_browse_base)

        # Fix serial path
        fs_layout = QHBoxLayout()
        self.fs_path_edit = QLineEdit()
        btn_browse_fs = QPushButton(" Chọn...")
        btn_browse_fs.setIcon(theme_mgr.get_styled_icon("folder"))
        btn_browse_fs.clicked.connect(self._browse_fix_serial)
        fs_layout.addWidget(self.fs_path_edit)
        fs_layout.addWidget(btn_browse_fs)

        # Reports dir
        rep_layout = QHBoxLayout()
        self.reports_dir_edit = QLineEdit()
        btn_browse_rep = QPushButton(" Chọn...")
        btn_browse_rep.setIcon(theme_mgr.get_styled_icon("folder"))
        btn_browse_rep.clicked.connect(lambda: self._browse_dir(self.reports_dir_edit, "Chọn thư mục lưu báo cáo"))
        rep_layout.addWidget(self.reports_dir_edit)
        rep_layout.addWidget(btn_browse_rep)

        # Shared Database path (SQLite LAN)
        shared_db_layout = QHBoxLayout()
        self.shared_db_edit = QLineEdit()
        btn_browse_db = QPushButton(" Chọn...")
        btn_browse_db.setIcon(theme_mgr.get_styled_icon("folder"))
        btn_browse_db.clicked.connect(self._browse_shared_db)
        shared_db_layout.addWidget(self.shared_db_edit)
        shared_db_layout.addWidget(btn_browse_db)

        # Update LAN Directory (.ssbupdate)
        update_dir_layout = QHBoxLayout()
        self.update_dir_edit = QLineEdit()
        btn_browse_update = QPushButton(" Chọn...")
        btn_browse_update.setIcon(theme_mgr.get_styled_icon("folder"))
        btn_browse_update.clicked.connect(lambda: self._browse_dir(self.update_dir_edit, "Chọn thư mục phát hành cập nhật LAN"))
        update_dir_layout.addWidget(self.update_dir_edit)
        update_dir_layout.addWidget(btn_browse_update)

        self.default_model_combo = QComboBox()
        self.default_model_combo.addItems(["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"])

        path_form.addRow("Thư mục gốc dự án:", base_layout)
        path_form.addRow("Tệp Master Fix Serial Tool:", fs_layout)
        path_form.addRow("Thư mục xuất Báo cáo:", rep_layout)
        path_form.addRow("CSDL Nhân sự chung (SQLite):", shared_db_layout)
        path_form.addRow("Thư mục cập nhật LAN (.ssbupdate):", update_dir_layout)
        path_form.addRow("Model máy mặc định:", self.default_model_combo)

        path_layout.addWidget(path_group)
        path_layout.addStretch()
        self.tab_widget.addTab(path_widget, theme_mgr.get_styled_icon("folder"), "Đường dẫn & Dữ liệu")

        # Tab 4: Giao diện (UI Theme)
        theme_widget = QWidget()
        theme_layout = QVBoxLayout(theme_widget)
        theme_group = QGroupBox("Giao diện & Hiển thị (UI / Theme)")
        theme_form = QFormLayout(theme_group)

        self.theme_combo = QComboBox()
        self.theme_combo.addItem("Giao diện Sáng (Slate Industrial)", "light")
        self.theme_combo.addItem("Giao diện Tối (Industrial Dark Mode)", "dark")
        self.theme_combo.addItem("Theo cài đặt hệ thống Windows", "system")
        self.theme_combo.currentIndexChanged.connect(self._on_theme_selection_changed)

        theme_form.addRow("Chế độ giao diện (Theme):", self.theme_combo)

        theme_note = QLabel(
            "<b>Chuẩn Enterprise Data-Dense:</b> Hỗ trợ thay đổi theme tức thì (hot-reload) mà không cần khởi động lại ứng dụng. "
            "Tối ưu tương phản WCAG AAA (> 7.0:1) cho môi trường nhà xưởng và văn phòng kỹ thuật."
        )
        theme_note.setWordWrap(True)
        theme_note.setStyleSheet("color: #475569; font-size: 11.5px; line-height: 1.4; padding-top: 8px;")
        theme_layout.addWidget(theme_group)
        theme_layout.addWidget(theme_note)
        theme_layout.addStretch()
        self.tab_widget.addTab(theme_widget, theme_mgr.get_styled_icon("sun"), "Giao diện (UI Theme)")

        # Dialog Buttons
        bottom_layout = QHBoxLayout()
        btn_reset = QPushButton(" Khôi phục mặc định")
        btn_reset.setIcon(theme_mgr.get_styled_icon("rotate-ccw"))
        btn_reset.clicked.connect(self.reset_to_defaults)

        btn_save = QPushButton(" Lưu cấu hình")
        btn_save.setIcon(theme_mgr.get_styled_icon("save", color="#FFFFFF"))
        btn_save.setDefault(True)
        btn_save.setStyleSheet("background-color: #059669; color: white; font-weight: bold; padding: 6px 14px;")
        btn_save.clicked.connect(self.save_and_close)

        btn_cancel = QPushButton("Đóng")
        btn_cancel.clicked.connect(self.reject)

        bottom_layout.addWidget(btn_reset)
        bottom_layout.addStretch()
        bottom_layout.addWidget(btn_save)
        bottom_layout.addWidget(btn_cancel)

        main_layout.addLayout(bottom_layout)

    # =========================================================================
    # Data Loading & Persistence
    # =========================================================================

    def load_settings(self, settings_dict: dict[str, Any]) -> None:
        """Bind settings dictionary values into form input fields."""
        tc = settings_dict.get("tc14", {})
        self.tc_url_edit.setText(tc.get("base_url", DEFAULT_SETTINGS["tc14"]["base_url"]))
        self.tc_user_edit.setText(tc.get("username", DEFAULT_SETTINGS["tc14"]["username"]))
        self.tc_pass_edit.setText(tc.get("password", DEFAULT_SETTINGS["tc14"]["password"]))
        self.tc_browser_combo.setCurrentText(tc.get("browser", "edge"))
        self.tc_headless_check.setChecked(bool(tc.get("headless", True)))
        self.tc_timeout_spin.setValue(int(tc.get("timeout", 30)))

        sap = settings_dict.get("sap", {})
        self.sap_system_edit.setText(sap.get("system_id", DEFAULT_SETTINGS["sap"]["system_id"]))
        self.sap_client_edit.setText(sap.get("client", DEFAULT_SETTINGS["sap"]["client"]))
        self.sap_plant_edit.setText(sap.get("plant", DEFAULT_SETTINGS["sap"]["plant"]))
        self.sap_usage_edit.setText(sap.get("bom_usage", DEFAULT_SETTINGS["sap"]["bom_usage"]))
        self.sap_alt_edit.setText(sap.get("alternative", DEFAULT_SETTINGS["sap"]["alternative"]))
        self.sap_path_edit.setText(sap.get("saplogon_path", DEFAULT_SETTINGS["sap"]["saplogon_path"]))

        paths = settings_dict.get("paths", {})
        self.base_dir_edit.setText(paths.get("base_dir", DEFAULT_SETTINGS["paths"]["base_dir"]))
        self.fs_path_edit.setText(paths.get("fix_serial_path", DEFAULT_SETTINGS["paths"]["fix_serial_path"]))
        self.reports_dir_edit.setText(paths.get("reports_dir", DEFAULT_SETTINGS["paths"]["reports_dir"]))
        self.shared_db_edit.setText(paths.get("shared_db_path", DEFAULT_SETTINGS["paths"]["shared_db_path"]))
        self.update_dir_edit.setText(paths.get("update_dir", DEFAULT_SETTINGS["paths"]["update_dir"]))
        model = paths.get("default_model", "Virgo")
        idx = self.default_model_combo.findText(model)
        if idx >= 0:
            self.default_model_combo.setCurrentIndex(idx)

        ui_cfg = settings_dict.get("ui", {})
        theme_code = ui_cfg.get("theme", settings_dict.get("theme", "light"))
        idx_t = self.theme_combo.findData(theme_code)
        if idx_t >= 0:
            self.theme_combo.setCurrentIndex(idx_t)

        self._current_settings = settings_dict

    def _on_theme_selection_changed(self) -> None:
        """Apply theme hot-reload immediately upon selection."""
        theme_code = self.theme_combo.currentData() or "light"
        get_theme_manager().set_theme(theme_code, save_preference=True)

    def get_settings(self) -> dict[str, Any]:
        """Collect current form values into settings dictionary."""
        theme_val = self.theme_combo.currentData() or "light"
        return {
            "tc14": {
                "base_url": self.tc_url_edit.text().strip(),
                "username": self.tc_user_edit.text().strip(),
                "password": self.tc_pass_edit.text().strip(),
                "browser": self.tc_browser_combo.currentText().strip(),
                "headless": self.tc_headless_check.isChecked(),
                "timeout": self.tc_timeout_spin.value(),
            },
            "sap": {
                "system_id": self.sap_system_edit.text().strip(),
                "client": self.sap_client_edit.text().strip(),
                "plant": self.sap_plant_edit.text().strip(),
                "bom_usage": self.sap_usage_edit.text().strip(),
                "alternative": self.sap_alt_edit.text().strip(),
                "saplogon_path": self.sap_path_edit.text().strip(),
            },
            "paths": {
                "base_dir": self.base_dir_edit.text().strip(),
                "fix_serial_path": self.fs_path_edit.text().strip(),
                "reports_dir": self.reports_dir_edit.text().strip(),
                "shared_db_path": self.shared_db_edit.text().strip(),
                "update_dir": self.update_dir_edit.text().strip(),
                "default_model": self.default_model_combo.currentText().strip(),
            },
            "ui": {
                "theme": theme_val,
            },
            "theme": theme_val,
        }

    def load_from_file(self) -> None:
        """Load settings from JSON config file if present, else use defaults."""
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                self.load_settings(data)
                logger.info("Đã nạp cấu hình hệ thống từ: %s", self.config_path)
                return
            except Exception as exc:
                logger.warning("Không thể đọc tệp cấu hình %s: %s", self.config_path, exc)
        self.load_settings(DEFAULT_SETTINGS)

    def save_and_close(self) -> None:
        """Save settings to config file and emit signal."""
        cfg = self.get_settings()
        try:
            theme_val = cfg.get("theme") or cfg.get("ui", {}).get("theme", "light")
            get_theme_manager().set_theme(theme_val, save_preference=True)
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as fp:
                json.dump(cfg, fp, indent=2, ensure_ascii=False)
            logger.info("Đã lưu cấu hình hệ thống vào: %s", self.config_path)
            self._current_settings = cfg
            self.settings_saved.emit(cfg)
            if self.isVisible():
                QMessageBox.information(self, "Thành công", "Đã lưu cấu hình hệ thống thành công.")
            self.accept()
        except Exception as exc:
            logger.error("Failed to write config file %s: %s", self.config_path, exc)
            QMessageBox.critical(self, "Lỗi lưu cấu hình", f"Không thể ghi tệp cấu hình:\n{exc}")

    def reset_to_defaults(self) -> None:
        """Reset form fields to hardcoded defaults."""
        reply = QMessageBox.question(
            self,
            "Xác nhận",
            "Bạn có chắc muốn khôi phục tất cả thông số về mặc định?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.load_settings(DEFAULT_SETTINGS)

    # =========================================================================
    # File Dialog Helpers & Connectivity Tests
    # =========================================================================

    def _browse_saplogon(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn saplogon.exe", "", "Executable (*.exe);;All Files (*)")
        if path:
            self.sap_path_edit.setText(path)

    def _browse_fix_serial(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn FIX_SERIAL_DLTOOL.xls", "", "Excel Files (*.xls *.xlsx);;All Files (*)")
        if path:
            self.fs_path_edit.setText(path)

    def _browse_shared_db(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn tệp CSDL Nhân sự SQLite (ssbom_master.db)",
            self.shared_db_edit.text(),
            "SQLite Database (*.db *.sqlite *.sqlite3);;All Files (*.*)",
        )
        if path:
            self.shared_db_edit.setText(path)

    def _browse_dir(self, line_edit: QLineEdit, title: str) -> None:
        folder = QFileDialog.getExistingDirectory(self, title, line_edit.text() or "")
        if folder:
            line_edit.setText(folder)

    def _test_tc_connection(self) -> None:
        """Check accessibility of TC14 URL."""
        import urllib.request
        url = self.tc_url_edit.text().strip()
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=5) as response:
                status_code = response.getcode()
            if status_code in (200, 301, 302):
                QMessageBox.information(self, "Kết nối TC24", f"Kết nối thành công tới cổng TC24!\nMã phản hồi HTTP: {status_code}")
            else:
                QMessageBox.warning(self, "Kết nối TC24", f"Cổng TC24 phản hồi với mã: {status_code}")
        except Exception as exc:
            QMessageBox.critical(self, "Kết nối TC24 Thất bại", f"Không thể kết nối tới {url}:\n{exc}")

    def _test_sap_connection(self) -> None:
        """Check SAP logon executable existence or COM scripting status."""
        sap_exe = self.sap_path_edit.text().strip()
        if os.path.exists(sap_exe):
            QMessageBox.information(self, "SAP GUI", f"Đã tìm thấy saplogon.exe tại:\n{sap_exe}")
        else:
            QMessageBox.warning(self, "SAP GUI", f"Không tìm thấy saplogon.exe tại đường dẫn đã cấu hình:\n{sap_exe}\nVui lòng kiểm tra lại.")
