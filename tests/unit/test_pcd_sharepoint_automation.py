"""Unit tests for PCD SharePoint Online browser automation downloader."""

from __future__ import annotations

from pathlib import Path
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.services.pcd_sharepoint_automation import (
    build_pcd_sharepoint_folder_url,
    PCDSharePointDownloadError,
)
from src.gui.pcd_plan_dialog import PCDPlanScanDialog, PCDSharePointDownloadWorker
from src.services.pcd_plan_service import PCDPlanService


class TestPCDSharePointAutomation:
    """Test suite for SharePoint headless download service and dialog integration."""

    @pytest.fixture(scope="session")
    def qapp(self) -> QApplication:
        app = QApplication.instance()
        if app is None:
            app = QApplication([])
        return app

    def test_build_pcd_sharepoint_folder_url(self) -> None:
        """Verify SharePoint folder URL contains correct encoded year/month path."""
        url = build_pcd_sharepoint_folder_url(year=2026, month=9)
        assert "kdcf.sharepoint.com" in url
        assert "sites/kdtvn_PCD" in url
        assert "ProductionPlan" in url
        assert "Theo%20th%C3%A1ng%20%28%E6%9C%88%E5%88%A5%29" in url
        assert "2026%2F202609" in url

    def test_dialog_has_auto_download_button_and_progress(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify dialog has the auto download button and progress bar."""
        plan_svc = PCDPlanService(base_dir=tmp_path)
        dlg = PCDPlanScanDialog(plan_service=plan_svc)

        assert hasattr(dlg, "btn_auto_download")
        assert "SharePoint" in dlg.btn_auto_download.text()
        assert hasattr(dlg, "pbar_download")
        assert hasattr(dlg, "lbl_download_status")

    def test_dialog_on_download_finished_updates_ui_and_scans(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Finished download sets path, updates label, and automatically triggers scan."""
        plan_svc = PCDPlanService(base_dir=tmp_path)
        dlg = PCDPlanScanDialog(plan_service=plan_svc)

        downloaded_file = tmp_path / "9月度定期計画後明細計画.xlsx"
        downloaded_file.write_text("dummy", encoding="utf-8")

        scan_called = []
        monkeypatch.setattr(dlg, "_on_scan_clicked", lambda: scan_called.append(True))
        monkeypatch.setattr(dlg, "_save_pcd_base_dir_to_settings", lambda p: None)

        dlg._on_download_finished(downloaded_file)

        assert dlg.edit_plan_file.text() == str(downloaded_file)
        assert "✓ Đã tải xong" in dlg.lbl_download_status.text()
        assert len(scan_called) == 1

    def test_dialog_on_download_error_handles_gracefully(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Download error shows warning dialog and resets button state."""
        plan_svc = PCDPlanService(base_dir=tmp_path)
        dlg = PCDPlanScanDialog(plan_service=plan_svc)

        warning_shown = []
        monkeypatch.setattr(
            QMessageBox,
            "warning",
            lambda *args, **kwargs: warning_shown.append(args),
        )

        dlg._on_download_error("Timeout waiting for SharePoint elements")

        assert dlg.btn_auto_download.isEnabled() is True
        assert dlg.pbar_download.isVisible() is False
        assert "thất bại" in dlg.lbl_download_status.text()
        assert len(warning_shown) == 1
        assert "Timeout" in warning_shown[0][2]

    def test_worker_signals_and_thread_lifecycle(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Worker thread properly emits progress and finished signals."""
        import src.services.pcd_sharepoint_automation as spa_mod

        fake_plan = tmp_path / "plan.xlsx"
        fake_plan.write_text("excel", encoding="utf-8")

        def fake_download(year, month, progress_callback=None, **kwargs):
            if progress_callback:
                progress_callback("[*] Testing progress")
            return fake_plan

        monkeypatch.setattr(spa_mod, "download_pcd_plan_from_sharepoint", fake_download)

        worker = PCDSharePointDownloadWorker(year=2026, month=9)

        progress_msgs = []
        results = []
        worker.progress.connect(lambda msg: progress_msgs.append(msg))
        worker.finished.connect(lambda res: results.append(res))

        worker.run()  # Run synchronously for unit test

        assert len(progress_msgs) == 1
        assert "Testing progress" in progress_msgs[0]
        assert len(results) == 1
        assert results[0] == fake_plan
