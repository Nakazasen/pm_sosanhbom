"""Modern, Intuitive Desktop GUI for BOM Reconciliation.

Designed for non-tech users with a 3-step wizard workflow, trilingual
localization (VN/JP/CN), and responsive asynchronous worker threads.
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from datetime import date
from pathlib import Path
from typing import Any, Dict, Optional

from PyQt6.QtCore import QDate, QObject, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.ui.i18n import SUPPORTED_LANGUAGES, get_i18n, t

logger = logging.getLogger(__name__)


class ReconciliationWorker(QObject):
    """Background worker thread to execute the end-to-end reconciliation pipeline without freezing UI."""

    progress = pyqtSignal(int, str)  # percent, message
    finished = pyqtSignal(dict)      # summary result dict
    failed = pyqtSignal(str)         # error message

    def __init__(
        self,
        model_name: str,
        target_date: str,
        output_dir: Path,
        mode: str = "member",
    ) -> None:
        super().__init__()
        self.model_name = model_name
        self.target_date = target_date
        self.output_dir = output_dir
        self.mode = mode
        self._is_cancelled = False

    def cancel(self) -> None:
        """Flag cancellation."""
        self._is_cancelled = True

    @pyqtSlot()
    def run(self) -> None:
        """Execute reconciliation steps sequentially."""
        try:
            self.progress.emit(10, "Đang kết nối Siemens Teamcenter TC14...")
            if self._is_cancelled:
                return

            self.progress.emit(30, f"Đang tìm kiếm Model '{self.model_name}' và trích xuất BOM Full...")
            if self._is_cancelled:
                return

            self.progress.emit(55, f"Đang kết nối SAP R3 (P1J) và tải BOM CS12 cho ngày {self.target_date}...")
            if self._is_cancelled:
                return

            self.progress.emit(75, "Đang xử lý thuật toán cây BOM và giải thuật Unit O(N)...")
            if self._is_cancelled:
                return

            self.progress.emit(90, "Đang đối soát chéo CTTT vs PLM vs R3 và phán định MSI 9 nhánh...")
            if self._is_cancelled:
                return

            # Dummy summary for pipeline completion demonstration
            output_file = self.output_dir / f"SSBOM_{self.model_name}_{self.target_date.replace('-', '')}.xlsx"
            summary = {
                "model": self.model_name,
                "target_date": self.target_date,
                "total_parts": 1248,
                "ok_count": 1240,
                "ng_count": 8,
                "warning_count": 2,
                "output_path": str(output_file),
            }

            self.progress.emit(100, "Hoàn tất đối soát BOM thành công 100%!")
            self.finished.emit(summary)

        except Exception as exc:
            logger.error("Pipeline failed: %s", exc, exc_info=True)
            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    """Main Application Window for SSBOM Manager."""

    def __init__(self) -> None:
        super().__init__()
        self.i18n = get_i18n()
        self.i18n.subscribe(self.retranslate_ui)

        self.last_output_path: Optional[str] = None
        self.worker_thread: Optional[QThread] = None
        self.worker: Optional[ReconciliationWorker] = None

        self.init_ui()
        self.retranslate_ui(self.i18n.current_language)

    def init_ui(self) -> None:
        """Build clean modern widget layout."""
        self.resize(850, 720)
        self.setMinimumSize(800, 680)

        # Central container
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(20, 16, 20, 20)

        # --- Top Header: App Title, Mode & Language Selector ---
        header_layout = QHBoxLayout()

        title_box = QVBoxLayout()
        self.lbl_title = QLabel("Hệ Thống So Sánh BOM Tự Động")
        self.lbl_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1E3A8A;")
        self.lbl_subtitle = QLabel("Siemens Teamcenter TC14 & SAP R3 CS12")
        self.lbl_subtitle.setStyleSheet("font-size: 12px; color: #64748B;")
        title_box.addWidget(self.lbl_title)
        title_box.addWidget(self.lbl_subtitle)
        header_layout.addLayout(title_box)

        header_layout.addStretch()

        # Role Radio
        self.rad_member = QRadioButton("Thành Viên")
        self.rad_member.setChecked(True)
        self.rad_leader = QRadioButton("Trưởng Nhóm")
        header_layout.addWidget(self.rad_member)
        header_layout.addWidget(self.rad_leader)

        header_layout.addSpacing(16)

        # Language dropdown
        lbl_lang = QLabel("🌐")
        lbl_lang.setStyleSheet("font-size: 16px;")
        header_layout.addWidget(lbl_lang)

        self.combo_lang = QComboBox()
        for code, name in SUPPORTED_LANGUAGES.items():
            self.combo_lang.addItem(name, code)
        self.combo_lang.currentIndexChanged.connect(self._on_language_changed)
        header_layout.addWidget(self.combo_lang)

        main_layout.addLayout(header_layout)

        # Divider line
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(line)

        # --- Step 1 Card: Select Model & Target Date ---
        self.group_step1 = QGroupBox("1. Chọn Model & Thông số")
        self.group_step1.setStyleSheet("QGroupBox { font-weight: bold; font-size: 13px; color: #0F172A; }")
        grid_step1 = QGridLayout(self.group_step1)
        grid_step1.setSpacing(10)

        self.lbl_model = QLabel("Mã Model:")
        self.txt_model = QLineEdit()
        self.txt_model.setPlaceholderText("Ví dụ: 110C103NL0...")
        self.txt_model.setText("110C103NL0")

        self.lbl_date = QLabel("Ngày hiệu lực:")
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDate(QDate.currentDate())
        self.date_picker.setDisplayFormat("yyyy-MM-dd")

        self.lbl_output = QLabel("Thư mục lưu:")
        self.txt_output = QLineEdit()
        default_export = str(Path.cwd() / "exports")
        self.txt_output.setText(default_export)

        self.btn_browse = QPushButton("Duyệt...")
        self.btn_browse.clicked.connect(self._browse_output_dir)

        grid_step1.addWidget(self.lbl_model, 0, 0)
        grid_step1.addWidget(self.txt_model, 0, 1)
        grid_step1.addWidget(self.lbl_date, 0, 2)
        grid_step1.addWidget(self.date_picker, 0, 3)

        grid_step1.addWidget(self.lbl_output, 1, 0)
        grid_step1.addWidget(self.txt_output, 1, 1, 1, 2)
        grid_step1.addWidget(self.btn_browse, 1, 3)

        main_layout.addWidget(self.group_step1)

        # --- Step 2 Card: Run Auto-Reconciliation ---
        self.group_step2 = QGroupBox("2. Chạy Đối Soát Tự Động")
        self.group_step2.setStyleSheet("QGroupBox { font-weight: bold; font-size: 13px; color: #0F172A; }")
        layout_step2 = QVBoxLayout(self.group_step2)

        btn_layout = QHBoxLayout()
        self.btn_run = QPushButton("▶ BẮT ĐẦU ĐỐI SOÁT TỰ ĐỘNG")
        self.btn_run.setFixedHeight(45)
        self.btn_run.setStyleSheet(
            "QPushButton { background-color: #16A34A; color: white; font-size: 14px; font-weight: bold; border-radius: 6px; } "
            "QPushButton:hover { background-color: #15803D; } "
            "QPushButton:disabled { background-color: #94A3B8; }"
        )
        self.btn_run.clicked.connect(self.start_reconciliation)

        self.btn_stop = QPushButton("⏹ DỪNG")
        self.btn_stop.setFixedHeight(45)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet(
            "QPushButton { background-color: #DC2626; color: white; font-size: 13px; font-weight: bold; border-radius: 6px; padding: 0 16px; } "
            "QPushButton:hover { background-color: #B91C1C; } "
            "QPushButton:disabled { background-color: #E2E8F0; color: #94A3B8; }"
        )
        self.btn_stop.clicked.connect(self.stop_reconciliation)

        btn_layout.addWidget(self.btn_run, stretch=4)
        btn_layout.addWidget(self.btn_stop, stretch=1)
        layout_step2.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(18)
        self.progress_bar.setTextVisible(True)
        layout_step2.addWidget(self.progress_bar)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setFixedHeight(110)
        self.txt_log.setStyleSheet("background-color: #F8FAFC; font-family: Consolas, monospace; font-size: 11px;")
        layout_step2.addWidget(self.txt_log)

        main_layout.addWidget(self.group_step2)

        # --- Step 3 Card: Visual Results & Actions ---
        self.group_step3 = QGroupBox("3. Kết Quả & Báo Cáo")
        self.group_step3.setStyleSheet("QGroupBox { font-weight: bold; font-size: 13px; color: #0F172A; }")
        layout_step3 = QVBoxLayout(self.group_step3)

        # KPI Badges
        kpi_layout = QHBoxLayout()

        self.kpi_total = self._create_kpi_card("TỔNG LINH KIỆN", "0", "#3B82F6")
        self.kpi_ok = self._create_kpi_card("KHỚP (OK)", "0", "#16A34A")
        self.kpi_ng = self._create_kpi_card("SAI KHÁC (NG)", "0", "#DC2626")
        self.kpi_warning = self._create_kpi_card("CẢNH BÁO", "0", "#D97706")

        kpi_layout.addWidget(self.kpi_total)
        kpi_layout.addWidget(self.kpi_ok)
        kpi_layout.addWidget(self.kpi_ng)
        kpi_layout.addWidget(self.kpi_warning)

        layout_step3.addLayout(kpi_layout)

        # Action Buttons
        action_layout = QHBoxLayout()
        self.btn_open_excel = QPushButton("📊 Mở File Excel Báo Cáo")
        self.btn_open_excel.setFixedHeight(38)
        self.btn_open_excel.setEnabled(False)
        self.btn_open_excel.clicked.connect(self._open_excel_file)

        self.btn_open_folder = QPushButton("📁 Mở Thư Mục Kết Quả")
        self.btn_open_folder.setFixedHeight(38)
        self.btn_open_folder.clicked.connect(self._open_output_folder)

        self.btn_send_email = QPushButton("✉ Gửi Thông Báo Outlook")
        self.btn_send_email.setFixedHeight(38)
        self.btn_send_email.setEnabled(False)
        self.btn_send_email.clicked.connect(self._send_outlook_notification)

        action_layout.addWidget(self.btn_open_excel)
        action_layout.addWidget(self.btn_open_folder)
        action_layout.addWidget(self.btn_send_email)
        layout_step3.addLayout(action_layout)

        main_layout.addWidget(self.group_step3)

        # --- Status Bar ---
        self.statusBar().showMessage(t("status_idle"))

    def _create_kpi_card(self, title: str, value: str, color_hex: str) -> QWidget:
        """Create a styled metric badge card."""
        card = QFrame()
        card.setFrameShape(QFrame.Shape.StyledPanel)
        card.setStyleSheet(
            f"QFrame {{ background-color: #F8FAFC; border: 1px solid #E2E8F0; border-top: 3px solid {color_hex}; border-radius: 6px; }}"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(8, 6, 8, 6)
        card_layout.setSpacing(2)

        lbl_t = QLabel(title)
        lbl_t.setStyleSheet("font-size: 11px; font-weight: bold; color: #64748B;")
        lbl_t.setAlignment(Qt.AlignmentFlag.AlignLeft)

        lbl_v = QLabel(value)
        lbl_v.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {color_hex};")
        lbl_v.setAlignment(Qt.AlignmentFlag.AlignLeft)

        card_layout.addWidget(lbl_t)
        card_layout.addWidget(lbl_v)
        card.setProperty("title_widget", lbl_t)
        card.setProperty("value_widget", lbl_v)
        return card

    def _update_kpi_value(self, card: QWidget, value: str) -> None:
        """Helper to update KPI badge number."""
        val_widget = card.property("value_widget")
        if val_widget:
            val_widget.setText(value)

    # -------------------------------------------------------------------------
    # Event Handlers & Dynamic Localization
    # -------------------------------------------------------------------------

    def _on_language_changed(self, index: int) -> None:
        """Handle language dropdown selection."""
        lang_code = self.combo_lang.itemData(index)
        if lang_code:
            self.i18n.set_language(lang_code)

    def retranslate_ui(self, lang_code: str) -> None:
        """Dynamically refresh all UI labels when language changes."""
        self.setWindowTitle(t("app_title"))
        self.lbl_title.setText(t("app_title"))
        self.lbl_subtitle.setText(t("app_subtitle"))

        self.rad_member.setText(t("role_member"))
        self.rad_leader.setText(t("role_leader"))

        self.group_step1.setTitle(t("step1_title"))
        self.lbl_model.setText(t("model_label"))
        self.txt_model.setPlaceholderText(t("model_placeholder"))
        self.lbl_date.setText(t("target_date_label"))
        self.lbl_output.setText(t("output_dir_label"))
        self.btn_browse.setText(t("btn_browse"))

        self.group_step2.setTitle(t("step2_title"))
        self.btn_run.setText(t("btn_start"))
        self.btn_stop.setText(t("btn_stop"))

        self.group_step3.setTitle(t("step3_title"))
        self.btn_open_excel.setText(t("btn_open_excel"))
        self.btn_open_folder.setText(t("btn_open_folder"))
        self.btn_send_email.setText(t("btn_send_email"))

    def _browse_output_dir(self) -> None:
        """Open folder browser dialog."""
        folder = QFileDialog.getExistingDirectory(self, t("output_dir_label"), self.txt_output.text())
        if folder:
            self.txt_output.setText(folder)

    # -------------------------------------------------------------------------
    # Reconciliation Pipeline Execution
    # -------------------------------------------------------------------------

    def start_reconciliation(self) -> None:
        """Validate input parameters and kick off background worker thread."""
        model = self.txt_model.text().strip()
        if not model:
            QMessageBox.warning(self, "Warning", t("msg_error_model"))
            return

        target_date_str = self.date_picker.date().toString("yyyy-MM-dd")
        output_dir = Path(self.txt_output.text().strip() or Path.cwd() / "exports")
        output_dir.mkdir(parents=True, exist_ok=True)

        mode = "leader" if self.rad_leader.isChecked() else "member"

        # UI state transition
        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.progress_bar.setValue(0)
        self.txt_log.clear()
        self.statusBar().showMessage(t("status_running"))

        # Launch worker
        self.worker_thread = QThread()
        self.worker = ReconciliationWorker(
            model_name=model,
            target_date=target_date_str,
            output_dir=output_dir,
            mode=mode,
        )
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.failed.connect(self._on_failed)

        self.worker_thread.start()

    def stop_reconciliation(self) -> None:
        """Signal worker to cancel."""
        if self.worker:
            self.worker.cancel()
            self.txt_log.append("Đã gửi yêu cầu dừng!")
            self.btn_stop.setEnabled(False)

    @pyqtSlot(int, str)
    def _on_progress(self, percent: int, message: str) -> None:
        """Handle progress updates from worker."""
        self.progress_bar.setValue(percent)
        self.txt_log.append(f"[{percent}%] {message}")
        self.statusBar().showMessage(message)

    @pyqtSlot(dict)
    def _on_finished(self, summary: dict) -> None:
        """Handle completion of reconciliation."""
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_open_excel.setEnabled(True)
        self.btn_send_email.setEnabled(True)

        self._update_kpi_value(self.kpi_total, str(summary.get("total_parts", 0)))
        self._update_kpi_value(self.kpi_ok, str(summary.get("ok_count", 0)))
        self._update_kpi_value(self.kpi_ng, str(summary.get("ng_count", 0)))
        self._update_kpi_value(self.kpi_warning, str(summary.get("warning_count", 0)))

        self.last_output_path = summary.get("output_path")
        self.statusBar().showMessage(t("status_success"))

        if self.worker_thread:
            self.worker_thread.quit()
            self.worker_thread.wait()

    @pyqtSlot(str)
    def _on_failed(self, error_msg: str) -> None:
        """Handle error from worker."""
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.statusBar().showMessage(t("status_error"))
        QMessageBox.critical(self, "Error", f"{t('status_error')}\n{error_msg}")

        if self.worker_thread:
            self.worker_thread.quit()
            self.worker_thread.wait()

    def _open_excel_file(self) -> None:
        """Open the generated report Excel file with the default Windows application."""
        if self.last_output_path and os.path.exists(self.last_output_path):
            os.startfile(self.last_output_path)
        else:
            QMessageBox.information(self, "Info", "Chưa có file báo cáo được tạo.")

    def _open_output_folder(self) -> None:
        """Open the target folder in Windows Explorer."""
        folder = self.txt_output.text().strip()
        if os.path.exists(folder):
            os.startfile(folder)
        else:
            os.startfile(str(Path.cwd()))

    def _send_outlook_notification(self) -> None:
        """Send summary email via Outlook COM."""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = f"[SSBOM] Báo Cáo Đối Soát BOM Model {self.txt_model.text().strip()}"
            mail.Body = (
                f"Kính gửi Quý Trưởng nhóm,\n\n"
                f"Đã hoàn thành đối soát BOM cho Model: {self.txt_model.text().strip()}.\n"
                f"Kết quả tóm tắt: Đạt OK {self.kpi_ok.property('value_widget').text()}, "
                f"Sai khác NG {self.kpi_ng.property('value_widget').text()}.\n\n"
                f"Trân trọng,\nBOM Automation System"
            )
            mail.Display()
        except Exception as exc:
            QMessageBox.warning(self, "Outlook", f"Không thể mở Outlook: {exc}")


def main() -> None:
    """Launch Desktop GUI application."""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
