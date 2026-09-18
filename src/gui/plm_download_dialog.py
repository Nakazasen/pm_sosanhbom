"""PLM BOM Auto-Download Dialog for Siemens Teamcenter 2412 (TC2412).

Provides an intuitive UI for entering BOM Part Numbers (e.g. T10C423NL0)
and downloading/standardizing them automatically in a background QThread.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, List, Optional

from PyQt6.QtCore import QObject, QThread, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.automation.tc2412.standardizer import (
    CANONICAL_14_COLUMNS,
    standardize_plm_file,
)

logger = logging.getLogger(__name__)

VIRGO_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)


class PLMDownloadWorker(QObject):
    """Executes Teamcenter 2412 download and standardization in a background thread."""

    progress = pyqtSignal(int, str)       # percent, message
    log_message = pyqtSignal(str)         # detailed log line
    finished = pyqtSignal(bool, str)      # success, summary_message

    def __init__(
        self,
        part_numbers: List[str],
        account: str,
        auto_standardize: bool,
        dest_dir: Path,
    ) -> None:
        super().__init__()
        self.part_numbers = part_numbers
        self.account = account
        self.auto_standardize = auto_standardize
        self.dest_dir = dest_dir

    @pyqtSlot()
    def run(self) -> None:
        """Run download and standardization pipeline."""
        total = len(self.part_numbers)
        success_count = 0

        self.progress.emit(5, f"Bắt đầu tiến trình tải {total} BOM từ Teamcenter 2412...")
        self.log_message.emit(f"[*] Khởi chạy phiên tải TC2412 với tài khoản: {self.account}")

        for idx, part in enumerate(self.part_numbers, 1):
            pct_start = int((idx - 1) / total * 90) + 5
            self.progress.emit(pct_start, f"[{idx}/{total}] Đang xử lý mã: {part}...")
            self.log_message.emit(f"\n--- Xử lý mã BOM: {part} ---")

            try:
                # Check if already present in backup or needs fresh download
                bak_dir = self.dest_dir / "backup_before_virgo_filter"
                raw_file = bak_dir / f"PLM_{part}.xlsx" if bak_dir.exists() else None

                target_dest = self.dest_dir / f"PLM_{part}.xlsx"

                if raw_file and raw_file.exists():
                    self.log_message.emit(f"[+] Tìm thấy bản gốc lưu trữ: {raw_file.name}")
                    orig_c, final_c = standardize_plm_file(
                        raw_file,
                        target_dest,
                        prune_electrical=self.auto_standardize,
                    )
                    self.log_message.emit(
                        f"[+] Chuẩn hóa 14 cột thành công: {orig_c:,} dòng raw -> {final_c:,} dòng chuẩn."
                    )
                else:
                    self.log_message.emit(f"[*] Tải trực tiếp từ Teamcenter 2412 cho {part}...")
                    # Lazy-load TC2412 Automation Client
                    from src.automation.tc2412.client import TC2412AutomationClient
                    from src.automation.tc2412.session import BrowserConfig, TC2412SessionManager

                    download_scratch = Path("scratch/downloads").resolve()
                    download_scratch.mkdir(parents=True, exist_ok=True)

                    config = BrowserConfig(download_dir=download_scratch, headless=True)
                    sm = TC2412SessionManager(config=config)
                    client = TC2412AutomationClient(session_manager=sm)

                    # Search, expand, select all, export
                    client.search_item(part)
                    time.sleep(3)
                    client.select_all_bom_tree()
                    time.sleep(2)
                    dl_path = client.trigger_export_download(f"PLM_{part}")

                    if dl_path and Path(dl_path).exists():
                        orig_c, final_c = standardize_plm_file(
                            Path(dl_path),
                            target_dest,
                            prune_electrical=self.auto_standardize,
                        )
                        self.log_message.emit(f"[+] Tải và chuẩn hóa thành công: {final_c:,} dòng.")
                    else:
                        raise RuntimeError("Không tải được tệp từ Teamcenter.")

                # Also sync into subfolder if exists
                for d in self.dest_dir.iterdir():
                    if d.is_dir() and d.name.upper().startswith(part.upper()):
                        import shutil
                        sub_file = d / f"PLM_{part}.xlsx"
                        shutil.copy2(target_dest, sub_file)
                        self.log_message.emit(f"[+] Đã đồng bộ sang thư mục kỹ sư: {d.name}")
                        break

                success_count += 1
            except Exception as exc:
                self.log_message.emit(f"[-] LỖI khi xử lý {part}: {exc}")

        self.progress.emit(100, "Hoàn thành toàn bộ tiến trình tải!")
        msg = f"Đã hoàn thành tải và chuẩn hóa {success_count}/{total} bản BOM PLM 14 cột tiêu chuẩn."
        self.finished.emit(success_count > 0, msg)


class PLMDownloadDialog(QDialog):
    """User-friendly dialog to input BOM Part Numbers and trigger TC2412 automated download."""

    def __init__(self, parent: QWidget | None = None, default_dir: Path | None = None) -> None:
        super().__init__(parent)
        self.dest_dir = Path(default_dir or VIRGO_DIR)
        self.thread: QThread | None = None
        self.worker: PLMDownloadWorker | None = None

        self.setWindowTitle("📥 Tải Tự Động BOM PLM Từ Siemens Teamcenter (TC2412)")
        self.resize(680, 520)
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Title / Description
        header_lbl = QLabel(
            "Nhập mã BOM (Part Number) để chương trình tự động đăng nhập Teamcenter 2412,\n"
            "bung cây cấu trúc, trích xuất dữ liệu và chuẩn hóa 14 cột tiêu chuẩn cho Tool so sánh."
        )
        header_lbl.setStyleSheet("color: #343a40; font-size: 13px; font-weight: bold;")
        layout.addWidget(header_lbl)

        # Form Group
        form_group = QGroupBox("Thông tin BOM cần tải")
        form_layout = QFormLayout(form_group)

        self.edit_parts = QLineEdit()
        self.edit_parts.setPlaceholderText("Ví dụ: T10C423NL0 hoặc T10C423NL0, T10C433NL0...")
        self.edit_parts.setFont(QFont("Consolas", 10))
        form_layout.addRow("Mã BOM / Part Number:", self.edit_parts)

        self.combo_account = QComboBox()
        self.combo_account.addItems([
            "vn_pe02 (Trọng - Chế tạo 2 / Preset KTCT_Trong)",
            "vn_pe03 (Phát triển hệ thống / Chế tạo 3)",
            "vn_pe01 (Chế tạo 1)",
            "vn_pe04 (Kỹ thuật Điện)",
        ])
        form_layout.addRow("Tài khoản tải:", self.combo_account)

        self.chk_standardize = QCheckBox("Tự động chuẩn hóa 14 cột tiêu chuẩn & gọt linh kiện điện Chế tạo 4")
        self.chk_standardize.setChecked(True)
        form_layout.addRow("Quy chuẩn dữ liệu:", self.chk_standardize)

        self.lbl_dest = QLabel(f"{self.dest_dir}")
        self.lbl_dest.setStyleSheet("color: #495057; font-size: 11px;")
        form_layout.addRow("Thư mục lưu:", self.lbl_dest)

        layout.addWidget(form_group)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("Sẵn sàng")
        self.lbl_status.setStyleSheet("font-style: italic; color: #495057;")
        layout.addWidget(self.lbl_status)

        # Log box
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 9))
        layout.addWidget(self.log_box)

        # Button row
        btn_layout = QHBoxLayout()
        self.btn_start = QPushButton("🚀 Bắt đầu tải BOM")
        self.btn_start.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_start.setStyleSheet(
            "background-color: #0d6efd; color: white; padding: 6px 14px; border-radius: 4px;"
        )
        self.btn_start.clicked.connect(self._on_start_download)
        btn_layout.addWidget(self.btn_start)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.clicked.connect(self.close)
        btn_layout.addWidget(self.btn_close)

        layout.addLayout(btn_layout)

    def _on_start_download(self) -> None:
        raw_text = self.edit_parts.text().strip()
        if not raw_text:
            QMessageBox.warning(self, "Chưa nhập mã", "Vui lòng nhập ít nhất một mã BOM cần tải!")
            return

        parts = [p.strip().upper() for p in raw_text.replace(";", ",").split(",") if p.strip()]
        if not parts:
            QMessageBox.warning(self, "Chưa nhập mã", "Vui lòng nhập mã BOM hợp lệ!")
            return

        account = self.combo_account.currentText().split()[0]
        auto_std = self.chk_standardize.isChecked()

        # Disable button during execution
        self.btn_start.setEnabled(False)
        self.log_box.clear()
        self.progress_bar.setValue(0)

        # Start QThread
        self.thread = QThread()
        self.worker = PLMDownloadWorker(
            part_numbers=parts,
            account=account,
            auto_standardize=auto_std,
            dest_dir=self.dest_dir,
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.log_message.connect(self._on_log)
        self.worker.finished.connect(self._on_finished)

        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def _on_progress(self, pct: int, msg: str) -> None:
        self.progress_bar.setValue(pct)
        self.lbl_status.setText(msg)

    def _on_log(self, text: str) -> None:
        self.log_box.append(text)

    def _on_finished(self, success: bool, msg: str) -> None:
        self.btn_start.setEnabled(True)
        if success:
            QMessageBox.information(self, "Tải BOM Hoàn tất", msg)
        else:
            QMessageBox.warning(self, "Thất bại", f"Không thể hoàn tất tải BOM:\n{msg}")
