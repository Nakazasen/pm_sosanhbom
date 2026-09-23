"""Unit tests for Milestone 3: Dialogs & App Shell Theme Integration.

Covers:
- TEST-M3-01: PLMDownloadDialog Data-Dense layout, 14px progress bar, 32px date table, SVG icons, 0 emojis.
- TEST-M3-02: SettingsDialog UI Theme tab, hot-reload instant switching, persistence in config.
- TEST-M3-03: SSBOMMainWindow App Shell theme application, tab icons, menu & toolbar icons, Ctrl+T toggle.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from unittest.mock import patch

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QPoint, QPointF, Qt
from PyQt6.QtGui import QWheelEvent
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.gui.app import SSBOMMainWindow
from src.gui.plm_download_dialog import NoScrollComboBox, PLMDownloadDialog
from src.gui.settings_dialog import SettingsDialog
from src.gui.styles import get_theme_manager, reset_theme_manager


class TestPLMDownloadDialogM3:
    """Test Data-Dense modernization of PLMDownloadDialog."""

    def test_plm_dialog_data_dense_elements(self, qapp: QApplication, tmp_path: Path) -> None:
        dlg = PLMDownloadDialog(default_dir=tmp_path)

        # 1. Zero raw emojis in title
        assert "📥" not in dlg.windowTitle()
        assert "Tải Tự Động BOM Đa Nguồn" in dlg.windowTitle()
        assert not dlg.windowIcon().isNull()

        # 2. Date table standard: 32px row height and visible grid
        assert dlg.date_table.showGrid() is True
        assert dlg.date_table.verticalHeader().defaultSectionSize() == 32

        # 3. Progress bar: compact height (14-16px including 1px border)
        assert dlg.progress_bar.height() in (14, 16) or dlg.progress_bar.maximumHeight() in (14, 16)

        # 4. Log box: Consolas font, read only
        assert dlg.log_box.isReadOnly() is True
        assert "consolas" in dlg.log_box.styleSheet().lower() or dlg.log_box.font().family().lower() in ("consolas", "monospace")

        # 5. Buttons have SVG icons
        for btn in (dlg.btn_paste, dlg.btn_clear, dlg.btn_browse, dlg.btn_both, dlg.btn_plm_only, dlg.btn_sap_only):
            assert not btn.icon().isNull(), f"Button {btn.text()} has null icon"
            # Ensure no raw emojis in button text
            assert "📋" not in btn.text()
            assert "🧹" not in btn.text()
            assert "📂" not in btn.text()
            assert "⚡" not in btn.text()
            assert "📥" not in btn.text()

        # 6. Auto-scroll check
        dlg._on_log("Test log entry 1")
        dlg._on_log("Test log entry 2")
        assert "Test log entry 2" in dlg.log_box.toPlainText()
        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_plm_dialog_accurate_progress_states(self, qapp: QApplication, tmp_path: Path, monkeypatch) -> None:
        """Verify that UI accurately displays Red on Failure, Amber on Partial, Green on Full Success."""
        monkeypatch.setattr(QMessageBox, "critical", lambda *args, **kwargs: QMessageBox.StandardButton.Ok)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: QMessageBox.StandardButton.Ok)
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: QMessageBox.StandardButton.Ok)

        dlg = PLMDownloadDialog(default_dir=tmp_path)

        class MockWorker:
            def __init__(self, succ: int, total: int, errs: list):
                self.successful_ops = succ
                self.total_ops = total
                self.error_messages = errs

        # Case 1: Total Failure (0/2 success)
        dlg.worker = MockWorker(0, 2, ["TC24 failed", "SAP failed"])
        dlg._on_finished(False, "Summary: 0/2")
        assert "ef4444" in dlg.progress_bar.styleSheet().lower() or "dc2626" in dlg.progress_bar.styleSheet().lower()
        assert "thất bại" in dlg.progress_bar.text().lower()
        assert "thất bại" in dlg.lbl_status.text().lower()
        assert "dc2626" in dlg.lbl_status.styleSheet().lower()

        # Case 2: Partial Success (1/2 success)
        dlg.worker = MockWorker(1, 2, ["SAP failed"])
        dlg._on_finished(True, "Summary: 1/2")
        assert "f59e0b" in dlg.progress_bar.styleSheet().lower() or "d97706" in dlg.progress_bar.styleSheet().lower()
        assert "một phần" in dlg.progress_bar.text().lower()
        assert "hoàn thành một phần" in dlg.lbl_status.text().lower()

        # Case 3: Full Success (2/2 success)
        dlg.worker = MockWorker(2, 2, [])
        dlg._on_finished(True, "Summary: 2/2")
        assert "10b981" in dlg.progress_bar.styleSheet().lower() or "059669" in dlg.progress_bar.styleSheet().lower()
        assert "thành công" in dlg.progress_bar.text().lower()
        assert "hoàn tất thành công" in dlg.lbl_status.text().lower()

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_plm_dialog_model_autodetection_and_dynamic_backup(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify that PLMDownloadDialog defaults to Auto model and dynamically adjusts backup dir."""
        # 1. Test with initial_model: default combo remains index 0 (-- Tự động nhận diện theo mã BOM --)
        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Polaris Next")
        assert dlg.combo_model.currentIndex() == 0
        assert dlg.combo_model.currentText() == "-- Tự động nhận diện theo mã BOM --"
        assert dlg.get_current_model_name() == "Polaris Next"
        assert "backup_before_polaris_next_filter" in dlg.lbl_backup_preview.text()

        # 2. Test auto-detection by pasting Polaris Next BOM part 110C103NL0
        dlg.txt_parts.setPlainText("110C103NL0")
        assert dlg.combo_model.currentIndex() == 0
        assert dlg.combo_model.currentText() == "-- Tự động nhận diện theo mã BOM --"
        assert "".join(dlg.get_current_model_name().split()) == "PolarisNext"
        assert "Polaris" in dlg.lbl_model_badge.text()
        assert "backup_before_polaris" in dlg.lbl_backup_preview.text()

        # 3. Test changing model manually updates backup preview
        dlg.combo_model.setCurrentText("Virgo")
        assert dlg.get_current_model_name() == "Virgo"
        assert "backup_before_virgo_filter" in dlg.lbl_backup_preview.text()

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_plm_dialog_comboboxes_ignore_mouse_wheel_scrolling(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify that scrolling the mouse wheel does NOT accidentally change TC24 account or Model.
        
        The user must explicitly click to open the dropdown and choose an item.
        """
        dlg = PLMDownloadDialog(default_dir=tmp_path)

        assert isinstance(dlg.combo_account, NoScrollComboBox)
        assert isinstance(dlg.combo_model, NoScrollComboBox)

        # TC24 Account Combobox test
        dlg.combo_account.setCurrentIndex(0)
        assert "vn_pe02" in dlg.combo_account.currentText()

        # Send wheel event (scroll down: angleDelta -120)
        wheel_down = QWheelEvent(
            QPointF(10, 10), QPointF(10, 10),
            QPoint(0, 0), QPoint(0, -120),
            Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
            Qt.ScrollPhase.NoScrollPhase, False
        )
        qapp.sendEvent(dlg.combo_account, wheel_down)
        assert dlg.combo_account.currentIndex() == 0, "combo_account changed on mouse wheel scroll down!"

        # Send wheel event (scroll up: angleDelta 120)
        wheel_up = QWheelEvent(
            QPointF(10, 10), QPointF(10, 10),
            QPoint(0, 0), QPoint(0, 120),
            Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
            Qt.ScrollPhase.NoScrollPhase, False
        )
        qapp.sendEvent(dlg.combo_account, wheel_up)
        assert dlg.combo_account.currentIndex() == 0, "combo_account changed on mouse wheel scroll up!"

        # Model Combobox test
        dlg.combo_model.setCurrentIndex(0)
        qapp.sendEvent(dlg.combo_model, wheel_down)
        assert dlg.combo_model.currentIndex() == 0, "combo_model changed on mouse wheel scroll down!"
        qapp.sendEvent(dlg.combo_model, wheel_up)
        assert dlg.combo_model.currentIndex() == 0, "combo_model changed on mouse wheel scroll up!"

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_plm_dialog_multi_model_autodetection_and_segregated_backup(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify multi-model BOM input detects all distinct models and shows individual breakdown."""
        dlg = PLMDownloadDialog(default_dir=tmp_path)
        # Paste 3 distinct models: Libra235ppm, Virgo, Sirius2
        dlg.txt_parts.setPlainText("110C3F3NL0\n1102Z42US0\n110C242US0")

        # 1. Check auto-detection map
        assert len(dlg._detected_models_map) == 3
        assert dlg._detected_models_map["110C3F3NL0"] == "Libra235ppm"
        assert dlg._detected_models_map["1102Z42US0"] == "Virgo"
        assert "Sirius2" in dlg._detected_models_map["110C242US0"]

        # 2. Check UI badge
        badge_text = dlg.lbl_model_badge.text()
        assert "Nhận diện đa dòng máy" in badge_text
        assert "Libra235ppm" in badge_text
        assert "Virgo" in badge_text
        assert "Sirius2" in badge_text
        assert not dlg.lbl_model_badge.isHidden()

        # 3. Check backup preview shows separate directories
        preview_text = dlg.lbl_backup_preview.text()
        assert "backup_before_libra235ppm_filter" in preview_text
        assert "backup_before_virgo_filter" in preview_text
        assert "backup_before_sirius2" in preview_text

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

    def test_plm_dialog_bolocbom_checkbox_and_worker_propagation(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify that chk_apply_bolocbom is enabled/checked by default, synchronizes with chk_standardize, and propagates to worker."""
        from src.gui.plm_download_dialog import UnifiedBOMDownloadWorker

        dlg = PLMDownloadDialog(default_dir=tmp_path)

        # 1. Check default state: checked and enabled
        assert hasattr(dlg, "chk_apply_bolocbom")
        assert dlg.chk_apply_bolocbom.isChecked() is True
        assert dlg.chk_apply_bolocbom.isEnabled() is True
        assert "BolocBom" in dlg.chk_apply_bolocbom.text()

        # 2. Toggle chk_standardize OFF -> chk_apply_bolocbom should be disabled and unchecked
        dlg.chk_standardize.setChecked(False)
        assert dlg.chk_apply_bolocbom.isEnabled() is False
        assert dlg.chk_apply_bolocbom.isChecked() is False

        # 3. Toggle chk_standardize back ON -> chk_apply_bolocbom should be enabled and checked
        dlg.chk_standardize.setChecked(True)
        assert dlg.chk_apply_bolocbom.isEnabled() is True
        assert dlg.chk_apply_bolocbom.isChecked() is True

        # 4. Test worker instantiation and propagation
        dlg.txt_parts.setPlainText("110C3F3NL0")
        assert len(dlg.current_items) == 1

        with patch("src.gui.plm_download_dialog.UnifiedBOMDownloadWorker") as MockWorkerClass, \
             patch("src.gui.plm_download_dialog.QThread") as MockThreadClass:
            # Case A: chk_apply_bolocbom is checked
            dlg._trigger_download("TC24")
            assert MockWorkerClass.called
            _, kwargs = MockWorkerClass.call_args
            assert kwargs.get("apply_bolocbom") is True

            # Case B: chk_apply_bolocbom is unchecked
            dlg.chk_apply_bolocbom.setChecked(False)
            dlg._trigger_download("TC24")
            _, kwargs2 = MockWorkerClass.call_args
            assert kwargs2.get("apply_bolocbom") is False

        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()



class TestSettingsDialogM3:
    """Test Theme integration in SettingsDialog."""

    def test_settings_dialog_theme_tab_and_hot_reload(self, qapp: QApplication, tmp_path: Path) -> None:
        config_file = tmp_path / "test_config.json"
        dlg = SettingsDialog(config_path=config_file)

        # 1. Tab count is at least 4, Tab 4 is UI Theme
        assert dlg.tab_widget.count() >= 4
        assert dlg.tab_widget.tabText(3) == "Giao diện (UI Theme)"

        # 2. Theme combo contains light, dark, system
        theme_values = [dlg.theme_combo.itemData(i) for i in range(dlg.theme_combo.count())]
        assert "light" in theme_values
        assert "dark" in theme_values
        assert "system" in theme_values

        # 3. Test hot-reload when combo selection changes
        theme_mgr = get_theme_manager()
        dlg.theme_combo.setCurrentIndex(dlg.theme_combo.findData("dark"))
        assert theme_mgr.current_theme == "dark"

        dlg.theme_combo.setCurrentIndex(dlg.theme_combo.findData("light"))
        assert theme_mgr.current_theme == "light"

        # 4. Save and persistence
        with patch.object(QMessageBox, "information"):
            dlg.theme_combo.setCurrentIndex(dlg.theme_combo.findData("dark"))
            dlg.save_and_close()

        assert config_file.exists()
        with open(config_file, "r", encoding="utf-8") as fp:
            saved_data = json.load(fp)
        assert saved_data.get("ui", {}).get("theme") == "dark"

        # 5. Re-load from file
        dlg2 = SettingsDialog(config_path=config_file)
        assert dlg2.theme_combo.currentData() == "dark"
        dlg.close()
        dlg.deleteLater()
        dlg2.close()
        dlg2.deleteLater()
        qapp.processEvents()


class TestAppShellThemeM3:
    """Test App Shell theme integration, tabs, menus, toolbars, and Ctrl+T."""

    def test_app_shell_theme_integration(self, qapp: QApplication, tmp_path: Path) -> None:
        with patch.object(SSBOMMainWindow, "_check_updates_silent"):
            main_win = SSBOMMainWindow(base_dir=tmp_path)

            # 1. ThemeManager is initialized
            assert hasattr(main_win, "theme_mgr")
            assert main_win.theme_mgr is not None
            assert not main_win.windowIcon().isNull()

            # 2. Tabs have clean text without emojis and valid SVG icons
            for i in range(main_win.tabs.count()):
                tab_text = main_win.tabs.tabText(i)
                assert "👔" not in tab_text
                assert "👷" not in tab_text
                assert not main_win.tabs.tabIcon(i).isNull()

            # Test toggle theme
            initial_theme = main_win.theme_mgr.current_theme
            main_win._toggle_theme()
            toggled_theme = main_win.theme_mgr.current_theme
            assert toggled_theme != initial_theme

            main_win._toggle_theme()
            assert main_win.theme_mgr.current_theme == initial_theme

            main_win.close()
            main_win.deleteLater()
            qapp.processEvents()
