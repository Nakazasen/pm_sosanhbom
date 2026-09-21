"""Update Dialog for SSBOM Manager.

Provides a clean, professional PyQt6 interface for checking, reviewing release notes,
downloading, and installing application updates from LAN sources.
"""

from __future__ import annotations

import logging
import subprocess
import sys
from typing import Optional

from PyQt6.QtCore import Qt, QThread, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.gui.styles import get_theme_manager
from src.services.app_updates import AppUpdateManager
from src.services.update_delivery import UpdateCandidate, UpdateDeliveryService

logger = logging.getLogger("update_dialog")


class UpdateWorker(QThread):
    """Background worker thread performing non-blocking download and installation."""

    progress_signal = pyqtSignal(int, int)  # (downloaded_bytes, total_bytes)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(
        self,
        delivery_service: UpdateDeliveryService,
        update_manager: AppUpdateManager,
        candidate: UpdateCandidate,
    ) -> None:
        super().__init__()
        self.delivery_service = delivery_service
        self.update_manager = update_manager
        self.candidate = candidate

    def run(self) -> None:
        try:
            self.status_signal.emit("Đang tải gói cập nhật từ mạng LAN...")

            def on_progress(copied: int, total: int) -> None:
                self.progress_signal.emit(copied, total)

            cached_pkg = self.delivery_service.download_package(
                self.candidate, progress_callback=on_progress
            )

            self.status_signal.emit("Đang giải nén an toàn & kiểm tra toàn vẹn SHA-256...")
            self.update_manager.install_update_package(cached_pkg)

            self.status_signal.emit("Cập nhật thành công!")
            self.finished_signal.emit(
                True,
                f"Đã cập nhật thành công lên phiên bản {self.candidate.version}.\n"
                f"Vui lòng khởi động lại ứng dụng để áp dụng.",
            )
        except Exception as exc:
            logger.error("Update process failed in worker: %s", exc)
            self.status_signal.emit(f"Lỗi: {exc}")
            self.finished_signal.emit(False, str(exc))


class UpdateDialog(QDialog):
    """Dialog displaying release notes and controlling update execution."""

    def __init__(
        self,
        parent: Optional[QWidget],
        candidate: UpdateCandidate,
        current_version: str,
        delivery_service: UpdateDeliveryService,
        update_manager: AppUpdateManager,
    ) -> None:
        super().__init__(parent)
        self.candidate = candidate
        self.current_version = current_version
        self.delivery_service = delivery_service
        self.update_manager = update_manager
        self.worker: Optional[UpdateWorker] = None

        self.setWindowTitle("Cập Nhật Phần Mềm SSBOM Manager")
        self.setMinimumWidth(540)
        self.setMinimumHeight(420)
        self.setModal(True)

        self._init_ui()

    def _init_ui(self) -> None:
        theme_mgr = get_theme_manager()
        self.setWindowIcon(theme_mgr.get_styled_icon("refresh"))

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Header card
        header_card = QFrame()
        header_card.setStyleSheet(
            "background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 6px; padding: 12px;"
        )
        card_layout = QVBoxLayout(header_card)
        card_layout.setContentsMargins(8, 8, 8, 8)

        title_lbl = QLabel("Có Phiên Bản Mới Sẵn Sàng!")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title_lbl.setFont(title_font)
        title_lbl.setStyleSheet("color: #0d6efd;")
        card_layout.addWidget(title_lbl)

        version_info = QLabel(
            f"<b>Phiên bản hiện tại:</b> <span style='color: #6c757d;'>v{self.current_version}</span> "
            f"➔ <b>Phiên bản mới:</b> <span style='color: #198754;'>v{self.candidate.version}</span>"
        )
        version_info.setTextFormat(Qt.TextFormat.RichText)
        card_layout.addWidget(version_info)

        size_mb = self.candidate.package_size / (1024 * 1024)
        meta_info = QLabel(f"Dung lượng gói: {size_mb:.2f} MB | Chính sách: HASH_ONLY_LAN")
        meta_info.setStyleSheet("color: #6c757d; font-size: 11px;")
        card_layout.addWidget(meta_info)

        layout.addWidget(header_card)

        # Notes section
        notes_lbl = QLabel("Nội dung thay đổi & ghi chú phát hành:")
        notes_lbl.setStyleSheet("font-weight: bold;")
        layout.addWidget(notes_lbl)

        self.notes_edit = QTextEdit()
        self.notes_edit.setReadOnly(True)
        self.notes_edit.setPlainText(self.candidate.release_notes or "Không có ghi chú bổ sung.")
        self.notes_edit.setStyleSheet(
            "background-color: #ffffff; border: 1px solid #ced4da; border-radius: 4px; padding: 6px;"
        )
        layout.addWidget(self.notes_edit)

        # Progress bar (hidden initially)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Status text
        self.status_lbl = QLabel("Sẵn sàng tải xuống từ mạng nội bộ công ty.")
        self.status_lbl.setStyleSheet("color: #495057; font-style: italic;")
        layout.addWidget(self.status_lbl)

        # Action buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Để sau")
        self.btn_cancel.setIcon(theme_mgr.get_styled_icon("x-circle"))
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_action = QPushButton("Tải & Cập Nhật Ngay")
        self.btn_action.setIcon(theme_mgr.get_styled_icon("download", color="#FFFFFF"))
        self.btn_action.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 6px 16px; border-radius: 4px;"
        )
        self.btn_action.clicked.connect(self._start_update)
        btn_layout.addWidget(self.btn_action)

        layout.addLayout(btn_layout)

    def _start_update(self) -> None:
        """Start downloading and updating."""
        self.btn_action.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.worker = UpdateWorker(
            self.delivery_service,
            self.update_manager,
            self.candidate,
        )
        self.worker.progress_signal.connect(self._on_progress)
        self.worker.status_signal.connect(self._on_status)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()

    @pyqtSlot(int, int)
    def _on_progress(self, copied: int, total: int) -> None:
        if total > 0:
            percent = int((copied / total) * 100)
            self.progress_bar.setValue(min(percent, 100))

    @pyqtSlot(str)
    def _on_status(self, msg: str) -> None:
        self.status_lbl.setText(msg)

    @pyqtSlot(bool, str)
    def _on_finished(self, success: bool, message: str) -> None:
        self.btn_cancel.setEnabled(True)
        theme_mgr = get_theme_manager()
        if success:
            self.progress_bar.setValue(100)
            self.status_lbl.setText("✓ " + message)
            self.status_lbl.setStyleSheet("color: #198754; font-weight: bold;")
            self.btn_action.setText("Khởi Động Lại Ngay")
            self.btn_action.setIcon(theme_mgr.get_styled_icon("refresh", color="#FFFFFF"))
            self.btn_action.setEnabled(True)
            self.btn_action.setStyleSheet(
                "background-color: #198754; color: white; font-weight: bold; padding: 6px 16px; border-radius: 4px;"
            )
            self.btn_action.clicked.disconnect()
            self.btn_action.clicked.connect(self._restart_app)
        else:
            self.status_lbl.setText("✗ " + message)
            self.status_lbl.setStyleSheet("color: #dc3545; font-weight: bold;")
            self.btn_action.setText("Thử lại")
            self.btn_action.setIcon(theme_mgr.get_styled_icon("rotate-ccw", color="#FFFFFF"))
            self.btn_action.setEnabled(True)
            self.btn_action.clicked.disconnect()
            self.btn_action.clicked.connect(self._start_update)

    def _restart_app(self) -> None:
        """Trigger application restart via SSBOM_Launcher."""
        logger.info("Đang khởi động lại ứng dụng SSBOM sau cập nhật...")
        install_root = self.update_manager.install_root

        # Check for SSBOM_Launcher.exe or SSBOM_Launcher.bat
        launcher_exe = install_root / "SSBOM_Launcher.exe"
        launcher_bat = install_root / "SSBOM_Launcher.bat"
        launcher_py = install_root / "SSBOM_Launcher.py"

        if launcher_exe.is_file():
            subprocess.Popen([str(launcher_exe)], cwd=str(install_root))
        elif launcher_bat.is_file():
            subprocess.Popen(["cmd.exe", "/c", str(launcher_bat)], cwd=str(install_root))
        elif launcher_py.is_file():
            subprocess.Popen([sys.executable, str(launcher_py)], cwd=str(install_root))

        sys.exit(0)
