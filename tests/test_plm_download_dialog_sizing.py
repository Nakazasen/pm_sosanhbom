"""Tests for PLMDownloadDialog responsive sizing and taskbar avoidance."""

import os
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.plm_download_dialog import PLMDownloadDialog
from src.gui.styles.screen_utils import ScreenMetrics


@pytest.fixture(scope="session")
def qapp():
    """Ensure QApplication instance exists for GUI tests."""
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestPLMDownloadDialogSizing:
    """Test suite ensuring PLMDownloadDialog never gets cut off by the taskbar."""

    def test_dialog_fits_laptop_768p_screen(self, qapp, tmp_path):
        """Simulate a 1366x768 laptop screen (work area 1366x728 due to Windows taskbar)."""
        fake_metrics = ScreenMetrics(
            screen_w=1366,
            screen_h=768,
            avail_x=0,
            avail_y=0,
            avail_w=1366,
            avail_h=728,
            device_pixel_ratio=1.0,
            dpi=96.0,
            is_compact=True,
            is_ultra_compact=False,
        )

        with patch("src.gui.styles.screen_utils.get_screen_metrics", return_value=fake_metrics):
            dlg = PLMDownloadDialog(
                default_dir=tmp_path,
                initial_model="Virgo",
                project_parts=["110C113NL0"],
            )
            try:
                geom = dlg.geometry()
                # 1. Height must be comfortably less than work area height 728
                assert geom.height() <= 640
                # 2. Window bottom must be strictly ABOVE the taskbar (at least 16px buffer)
                assert geom.bottom() <= (728 - 16)
                # 3. Window top must be on screen
                assert geom.top() >= 0
                # 4. Action buttons in Step 3 must be visible and properly initialized
                assert dlg.btn_both.isVisibleTo(dlg) or not dlg.isHidden()
                assert dlg.btn_both.text().strip().startswith("Tải đồng thời")
                assert dlg.btn_close.text() == "Đóng"
            finally:
                dlg.close()

    def test_dialog_fits_high_dpi_scaled_screen(self, qapp, tmp_path):
        """Simulate a 1920x1080 screen with 150% Windows display scaling (effective 1280x720, avail 672)."""
        fake_metrics = ScreenMetrics(
            screen_w=1280,
            screen_h=720,
            avail_x=0,
            avail_y=0,
            avail_w=1280,
            avail_h=672,
            device_pixel_ratio=1.5,
            dpi=144.0,
            is_compact=True,
            is_ultra_compact=False,
        )

        with patch("src.gui.styles.screen_utils.get_screen_metrics", return_value=fake_metrics):
            dlg = PLMDownloadDialog(
                default_dir=tmp_path,
                initial_model="Virgo",
            )
            try:
                geom = dlg.geometry()
                # Dialog height must scale down to fit inside 672 work area with margin
                assert geom.height() <= 590
                assert geom.bottom() <= (672 - 16)
                assert geom.top() >= 0
            finally:
                dlg.close()
