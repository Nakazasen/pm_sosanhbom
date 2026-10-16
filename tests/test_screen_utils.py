"""Unit tests for ScreenUtils and Display Scaling adaptation."""

import os
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtCore import QRect
from PyQt6.QtWidgets import QApplication, QDialog

from src.gui.styles.screen_utils import (
    ScreenMetrics,
    apply_smart_dialog_geometry,
    compute_smart_geometry,
    ensure_window_fits_screen,
    get_screen_metrics,
)


@pytest.fixture(scope="session")
def qapp():
    """Ensure a single QApplication instance exists for GUI tests."""
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestScreenUtils:
    """Test suite for screen geometry calculations and taskbar overflow prevention."""

    def test_get_screen_metrics_headless_or_default(self, qapp):
        """Verify screen metrics returns valid dataclass structure."""
        metrics = get_screen_metrics()
        assert isinstance(metrics, ScreenMetrics)
        assert metrics.screen_w > 0
        assert metrics.screen_h > 0
        assert metrics.avail_w > 0
        assert metrics.avail_h > 0
        assert metrics.avail_h <= metrics.screen_h

    def test_compute_smart_geometry_standard_fhd(self):
        """On 1080p screen (avail 1920x1032), ideal dimensions should be preserved."""
        fake_metrics = ScreenMetrics(
            screen_w=1920,
            screen_h=1080,
            avail_x=0,
            avail_y=0,
            avail_w=1920,
            avail_h=1032,
            device_pixel_ratio=1.0,
            dpi=96.0,
            is_compact=False,
            is_ultra_compact=False,
        )
        with patch("src.gui.styles.screen_utils.get_screen_metrics", return_value=fake_metrics):
            x, y, w, h, is_compact = compute_smart_geometry(
                ideal_w=880,
                ideal_h=680,
                min_w=640,
                min_h=380,
                margin_bottom=16,
            )
            assert w == 880
            assert h == 680
            assert is_compact is False
            # Ensure window bottom is well above taskbar
            assert (y + h) <= (1032 - 16)
            # Ensure centered horizontally
            assert x == (1920 - 880) // 2

    def test_compute_smart_geometry_laptop_768p_taskbar_cut_off_prevention(self):
        """On 1366x768 screen (avail 1366x728), height MUST be clamped so bottom never touches taskbar."""
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
            x, y, w, h, is_compact = compute_smart_geometry(
                ideal_w=880,
                ideal_h=720,  # User requested 720 which would normally collide with 728
                min_w=640,
                min_h=380,
                max_h_ratio=0.86,
                margin_bottom=16,
            )
            assert is_compact is True
            # Height MUST be clamped below 728 - 16 = 712, and specifically <= int(728 * 0.86) = 626
            assert h <= 626
            # Crucial: Bottom of window MUST be at least 16px above the taskbar!
            assert (y + h) <= (728 - 16)
            assert y >= 0

    def test_compute_smart_geometry_high_dpi_scaled_laptop(self):
        """On 1080p laptop with 150% scaling, effective logical resolution is 1280x720 (avail 1280x672)."""
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
            x, y, w, h, is_compact = compute_smart_geometry(
                ideal_w=880,
                ideal_h=700,
                min_w=600,
                min_h=380,
                max_h_ratio=0.86,
                margin_bottom=20,
            )
            assert is_compact is True
            # Max allowed height is min(int(672 * 0.86), 672 - 20) = 577
            assert h <= 577
            assert (y + h) <= (672 - 20)
            assert x >= 0

    def test_apply_smart_dialog_geometry(self, qapp):
        """Verify apply_smart_dialog_geometry configures the dialog correctly."""
        dlg = QDialog()
        try:
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
                is_compact = apply_smart_dialog_geometry(dlg, ideal_w=880, ideal_h=700)
                assert is_compact is True
                geom = dlg.geometry()
                assert geom.width() <= 880
                assert geom.height() <= 626
                assert geom.bottom() <= (728 - 16)
        finally:
            dlg.close()

    def test_ensure_window_fits_screen_repositions_overflowing_window(self, qapp):
        """Verify ensure_window_fits_screen pulls back a window that hung below the taskbar."""
        dlg = QDialog()
        try:
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
            # Position dialog so bottom is at 780 (52px below taskbar)
            dlg.setGeometry(100, 200, 800, 580)
            with patch("src.gui.styles.screen_utils.get_screen_metrics", return_value=fake_metrics):
                ensure_window_fits_screen(dlg, margin_bottom=16)
                new_geom = dlg.geometry()
                # Must be pushed up so bottom <= 728 - 16 = 712
                assert new_geom.bottom() <= (728 - 16)
                assert new_geom.top() >= 0
        finally:
            dlg.close()
