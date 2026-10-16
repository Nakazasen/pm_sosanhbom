"""Unit tests verifying Bug 1 (consecutive BOM downloads) and Bug 2 (automatic update detection).

Covers:
1. TEST-BUG1-01: PLMDownloadDialog consecutive downloads (Download 1 -> complete -> change BOM -> Download 2 -> complete) without crash.
2. TEST-BUG1-02: PLMDownloadDialog thread and worker cleanup on dialog close.
3. TEST-BUG1-03: PLMDownloadDialog safe extraction of stats in _on_finished even if worker deleted.
4. TEST-BUG2-01: SilentUpdateCheckWorker background QThread detects candidate and emits candidate_found signal.
5. TEST-BUG2-02: SSBOMMainWindow._check_updates_silent correctly routes candidate_found to _show_update_candidate.
6. TEST-BUG2-03: SSBOMMainWindow._show_update_candidate guards against opening duplicate dialogs.
"""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtCore import QEventLoop, QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.gui.app import SSBOMMainWindow, SilentUpdateCheckWorker
from src.gui.plm_download_dialog import (
    MODE_PLM,
    PLMDownloadDialog,
    UnifiedBOMDownloadWorker,
)
from src.services.update_delivery import UpdateCandidate


class TestBug1ConsecutiveBOMDownload:
    """Verifies that downloading BOM, changing BOM code, and downloading again does not crash."""

    def test_consecutive_downloads_no_crash(self, qapp: QApplication, tmp_path: Path) -> None:
        """Download batch 1, complete, enter new BOM, download batch 2, complete without QThread crash."""
        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo")

        # Fake execution that emits progress and finished
        def fake_run(worker_self, total_parts):
            worker_self.successful_ops = total_parts
            worker_self.total_ops = total_parts
            worker_self.error_messages = []
            worker_self.progress.emit(100, f"Done {total_parts}")
            worker_self.finished.emit(True, f"Completed {total_parts} parts")

        with patch.object(QMessageBox, "information"), \
             patch.object(QMessageBox, "warning"), \
             patch.object(QMessageBox, "critical"), \
             patch.object(UnifiedBOMDownloadWorker, "_execute_run", fake_run):

            # 1. First download
            dlg.txt_parts.setPlainText("110C113NL0")
            dlg._trigger_download(MODE_PLM)

            # Wait for first thread to finish
            loop = QEventLoop()
            dlg.thread.finished.connect(loop.quit)
            loop.exec()
            qapp.processEvents()

            assert dlg.progress_bar.value() == 100
            assert "100%" in dlg.progress_bar.text()
            assert dlg.thread is None
            assert dlg.worker is None

            # 2. Enter new BOM code
            dlg.txt_parts.clear()
            qapp.processEvents()
            dlg.txt_parts.setPlainText("110C122US0")
            qapp.processEvents()
            assert len(dlg.current_items) == 1
            assert dlg.current_items[0].part_number == "110C122US0"

            # 3. Second download (previously crashed due to QThread: Destroyed while thread is still running)
            dlg._trigger_download(MODE_PLM)
            assert dlg.thread is not None
            assert dlg.thread.isRunning()

            loop2 = QEventLoop()
            dlg.thread.finished.connect(loop2.quit)
            loop2.exec()
            qapp.processEvents()

            assert dlg.progress_bar.value() == 100
            assert "100%" in dlg.progress_bar.text()
            assert dlg.thread is None
            assert dlg.worker is None

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_close_event_cleans_up_running_thread(self, qapp: QApplication, tmp_path: Path) -> None:
        """Closing dialog while a thread is running safely quits and waits on the thread."""
        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo")

        mock_thread = MagicMock()
        mock_thread.isRunning.return_value = True
        dlg.thread = mock_thread
        mock_worker = MagicMock()
        dlg.worker = mock_worker

        dlg.close()
        assert mock_thread.quit.called
        assert mock_thread.wait.called
        assert dlg.thread is None
        assert dlg.worker is None

        dlg.deleteLater()
        qapp.processEvents()

    def test_on_finished_safe_when_worker_deleted(self, qapp: QApplication, tmp_path: Path) -> None:
        """_on_finished does not raise RuntimeError if worker object is already None or deleted."""
        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo")
        dlg.worker = None

        with patch.object(QMessageBox, "information"):
            # Should not raise exception
            dlg._on_finished(True, "All operations completed")

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()


    def test_clean_up_thread_stuck_thread_safety(self, qapp: QApplication, tmp_path: Path) -> None:
        """Stuck thread (wait returns False) is retained in _stale_threads to prevent Python GC crash."""
        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo")

        mock_thread = MagicMock()
        mock_thread.isRunning.return_value = True
        mock_thread.wait.return_value = False  # Timed out!
        dlg.thread = mock_thread

        dlg._clean_up_thread()
        # Thread was retained in _stale_threads so GC doesn't delete it while running
        assert hasattr(dlg, "_stale_threads")
        assert mock_thread in dlg._stale_threads
        assert dlg.thread is None

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_consecutive_three_downloads_with_mode_switch(self, qapp: QApplication, tmp_path: Path) -> None:
        """Run 3 downloads in sequence switching between modes without thread corruption."""
        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo")

        def fake_run(worker_self, total_parts):
            worker_self.successful_ops = total_parts
            worker_self.total_ops = total_parts
            worker_self.error_messages = []
            worker_self.progress.emit(100, f"Done {total_parts}")
            worker_self.finished.emit(True, f"Completed {total_parts}")

        from src.gui.plm_download_dialog import MODE_BOTH, MODE_SAP

        with patch.object(QMessageBox, "information"), \
             patch.object(QMessageBox, "warning"), \
             patch.object(QMessageBox, "critical"), \
             patch.object(UnifiedBOMDownloadWorker, "_execute_run", fake_run):

            for part_code, mode in [("110C113NL0", MODE_PLM), ("110C122US0", MODE_SAP), ("110C130EU0", MODE_BOTH)]:
                dlg.txt_parts.setPlainText(part_code)
                dlg._trigger_download(mode)

                loop = QEventLoop()
                dlg.thread.finished.connect(loop.quit)
                loop.exec()
                qapp.processEvents()

                assert dlg.progress_bar.value() == 100
                assert dlg.thread is None

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()


class TestBug2AutoUpdateDetection:
    """Verifies that automatic update detection discovers candidate and triggers UI suggestion."""

    def test_load_sources_configuration_falls_back_to_settings_json(self, tmp_path: Path) -> None:
        """When update_sources.default.json is missing, sources are populated from config/settings.json."""
        from src.services.update_delivery import UpdateDeliveryService

        cfg_dir = tmp_path / "config"
        cfg_dir.mkdir(parents=True)
        custom_lan = r"\\custom-server\updates"
        (cfg_dir / "settings.json").write_text(
            json.dumps({"paths": {"update_dir": custom_lan}}),
            encoding="utf-8",
        )

        delivery = UpdateDeliveryService(base_dir=tmp_path)
        cfg = delivery.load_sources_configuration()

        matching = [s for s in cfg.get("sources", []) if s.get("location") == custom_lan]
        assert len(matching) == 1
        assert matching[0]["enabled"] is True

    def test_silent_update_worker_emits_candidate_found(self, qapp: QApplication, tmp_path: Path) -> None:
        """SilentUpdateCheckWorker discovers candidate and emits signal across threads."""
        worker = SilentUpdateCheckWorker(base_dir=tmp_path)

        mock_candidate = UpdateCandidate(
            version="2.0.0",
            package_name="SSBOM_Manager-2.0.0.mpupdate",
            package_sha256="a" * 64,
            package_size=1024,
            release_notes="New features in v2.0.0",
            source_location=str(tmp_path),
        )

        emitted = []
        def on_candidate(candidate, cur_ver, delivery, mgr):
            emitted.append((candidate, cur_ver))

        worker.candidate_found.connect(on_candidate)

        with patch("src.services.update_delivery.UpdateDeliveryService.load_sources_configuration", return_value={"startup_check": True}), \
             patch("src.services.update_delivery.UpdateDeliveryService.check_for_updates", return_value=mock_candidate), \
             patch("src.services.app_updates.AppUpdateManager.get_active_version", return_value="1.0.3"), \
             patch("src.gui.app.get_build_version", return_value="1.0.3"):
            worker.start()
            worker.wait(5000)
            qapp.processEvents()

        assert len(emitted) == 1
        cand, cur = emitted[0]
        assert cand.version == "2.0.0"
        assert cur == "1.0.3"

    def test_check_updates_silent_launches_worker_and_shows_dialog(self, qapp: QApplication, tmp_path: Path) -> None:
        """_check_updates_silent spawns QThread and routes candidate_found to _show_update_candidate."""
        with patch.object(SSBOMMainWindow, "_check_updates_silent"):
            win = SSBOMMainWindow(base_dir=tmp_path)

        mock_candidate = UpdateCandidate(
            version="2.1.0",
            package_name="SSBOM_Manager-2.1.0.mpupdate",
            package_sha256="b" * 64,
            package_size=2048,
            release_notes="Bug fixes",
            source_location=str(tmp_path),
        )

        shown = []
        def mock_show(candidate, cur_ver, delivery, mgr):
            shown.append(candidate.version)

        win._show_update_candidate = mock_show

        with patch("src.services.update_delivery.UpdateDeliveryService.load_sources_configuration", return_value={"startup_check": True}), \
             patch("src.services.update_delivery.UpdateDeliveryService.check_for_updates", return_value=mock_candidate), \
             patch("src.services.app_updates.AppUpdateManager.get_active_version", return_value="1.0.3"):
            # Now run the actual fixed _check_updates_silent
            SSBOMMainWindow._check_updates_silent(win)

            assert win._update_check_worker is not None
            win._update_check_worker.wait(5000)
            qapp.processEvents()

        assert len(shown) == 1
        assert shown[0] == "2.1.0"

        win.close()
        win.deleteLater()
        qapp.processEvents()

    def test_show_update_candidate_guards_against_duplicates(self, qapp: QApplication, tmp_path: Path) -> None:
        """_show_update_candidate does not open a second dialog if one is already open."""
        with patch.object(SSBOMMainWindow, "_check_updates_silent"):
            win = SSBOMMainWindow(base_dir=tmp_path)

        mock_dialog = MagicMock()
        mock_dialog.isVisible.return_value = True
        win._active_update_dialog = mock_dialog

        mock_candidate = MagicMock()
        # Call _show_update_candidate while active dialog is visible
        win._show_update_candidate(mock_candidate, "1.0.0", MagicMock(), MagicMock())

        # Should not create new UpdateDialog or call exec
        assert mock_dialog.exec.call_count == 0

        win.close()
        win.deleteLater()
        qapp.processEvents()

    def test_window_close_while_worker_running_does_not_crash(self, qapp: QApplication, tmp_path: Path) -> None:
        """Closing SSBOMMainWindow while silent worker is still running does not cause abort."""
        with patch.object(SSBOMMainWindow, "_check_updates_silent"):
            win = SSBOMMainWindow(base_dir=tmp_path)

        mock_worker = MagicMock()
        mock_worker.isRunning.return_value = True
        win._update_check_worker = mock_worker

        win.close()
        assert mock_worker.quit.called
        assert mock_worker.wait.called
        assert win._update_check_worker is None

        win.deleteLater()
        qapp.processEvents()

    def test_window_close_when_worker_already_deleted_does_not_crash(self, qapp: QApplication, tmp_path: Path) -> None:
        """Closing SSBOMMainWindow when worker C++ object was deleted never raises RuntimeError."""
        with patch.object(SSBOMMainWindow, "_check_updates_silent"):
            win = SSBOMMainWindow(base_dir=tmp_path)

        deleted_worker = MagicMock()
        deleted_worker.isRunning.side_effect = RuntimeError(
            "wrapped C/C++ object of type SilentUpdateCheckWorker has been deleted"
        )
        win._update_check_worker = deleted_worker

        # Must not raise RuntimeError
        win.close()
        assert win._update_check_worker is None

        win.deleteLater()
        qapp.processEvents()

    def test_silent_worker_clears_reference_on_finished(self, qapp: QApplication, tmp_path: Path) -> None:
        """When SilentUpdateCheckWorker finishes, _update_check_worker reference is cleared."""
        with patch.object(SSBOMMainWindow, "_check_updates_silent"):
            win = SSBOMMainWindow(base_dir=tmp_path)

        mock_worker = SilentUpdateCheckWorker(base_dir=tmp_path)
        win._update_check_worker = mock_worker
        mock_worker.finished.connect(lambda w=mock_worker: win._on_silent_worker_finished(w))

        # Emit finished
        mock_worker.finished.emit()
        qapp.processEvents()

        assert win._update_check_worker is None

        win.close()
        win.deleteLater()
        qapp.processEvents()

