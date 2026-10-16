"""Intelligent Screen & Display Scaling Utilities for Desktop GUI (PyQt6).

Provides:
- Resolution & High-DPI Display Scaling detection (Windows 10/11 taskbar & DPI awareness)
- Dynamic geometry calculations ensuring dialogs and windows never get cut off by the taskbar
- Compact-mode detection for low-resolution / high-scaling laptop displays
- Auto-centering and bounds-clamping within QScreen.availableGeometry()
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Optional, Tuple

from PyQt6.QtCore import QPoint, QRect
from PyQt6.QtGui import QCursor, QGuiApplication, QScreen
from PyQt6.QtWidgets import QApplication, QDialog, QWidget

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ScreenMetrics:
    """Represents display resolution and available work area metrics."""

    screen_w: int
    screen_h: int
    avail_x: int
    avail_y: int
    avail_w: int
    avail_h: int
    device_pixel_ratio: float
    dpi: float
    is_compact: bool
    is_ultra_compact: bool


def get_target_screen(widget_or_parent: Optional[QWidget] = None) -> Optional[QScreen]:
    """Retrieve the most relevant QScreen for the given widget or user context.

    Search priority:
    1. Widget's own screen (if already assigned/shown).
    2. Window/Parent's screen.
    3. Screen containing the current mouse cursor (crucial for multi-monitor setups).
    4. Primary screen.
    """
    if widget_or_parent is not None:
        try:
            scr = widget_or_parent.screen()
            if scr:
                return scr
        except Exception:
            pass

        try:
            win = widget_or_parent.window()
            if win and win is not widget_or_parent:
                scr = win.screen()
                if scr:
                    return scr
        except Exception:
            pass

    app = QApplication.instance()
    if app is not None and isinstance(app, QGuiApplication):
        try:
            pos = QCursor.pos()
            scr = QGuiApplication.screenAt(pos)
            if scr:
                return scr
        except Exception:
            pass

        primary = app.primaryScreen()
        if primary:
            return primary

    return None


def get_screen_metrics(widget_or_parent: Optional[QWidget] = None) -> ScreenMetrics:
    """Safely extract screen dimensions, available geometry, and scaling metrics.

    Falls back to safe desktop standard values (1366x768, 1366x720 avail) if headless.
    """
    screen = get_target_screen(widget_or_parent)
    if screen is not None:
        geom: QRect = screen.geometry()
        avail: QRect = screen.availableGeometry()
        dpr: float = float(screen.devicePixelRatio())
        dpi: float = float(screen.logicalDotsPerInch())
        screen_w = geom.width()
        screen_h = geom.height()
        avail_x = avail.x()
        avail_y = avail.y()
        avail_w = avail.width()
        avail_h = avail.height()
    else:
        # Standard fallback for headless/test environments
        screen_w, screen_h = 1366, 768
        avail_x, avail_y = 0, 0
        avail_w, avail_h = 1366, 720
        dpr = 1.0
        dpi = 96.0

    # Compact display detection:
    # Typical: 1366x768 display, or 1920x1080 at 125%/150% scaling where avail_h <= 740
    is_compact = (avail_h <= 740) or (avail_w <= 1280) or (dpr >= 1.25 and avail_h <= 800)
    is_ultra_compact = (avail_h <= 620) or (avail_w <= 1024)

    return ScreenMetrics(
        screen_w=screen_w,
        screen_h=screen_h,
        avail_x=avail_x,
        avail_y=avail_y,
        avail_w=avail_w,
        avail_h=avail_h,
        device_pixel_ratio=dpr,
        dpi=dpi,
        is_compact=is_compact,
        is_ultra_compact=is_ultra_compact,
    )


def compute_smart_geometry(
    widget_or_parent: Optional[QWidget] = None,
    ideal_w: int = 880,
    ideal_h: int = 680,
    min_w: int = 640,
    min_h: int = 380,
    max_w_ratio: float = 0.94,
    max_h_ratio: float = 0.86,
    margin_bottom: int = 16,
) -> Tuple[int, int, int, int, bool]:
    """Calculate (x, y, target_w, target_h, is_compact) fitted strictly within available geometry.

    Ensures the window never extends beneath the Windows taskbar and leaves a safe cushion.
    """
    metrics = get_screen_metrics(widget_or_parent)

    # Maximum usable bounds inside available geometry (excluding taskbar)
    max_allowed_w = max(min_w, int(metrics.avail_w * max_w_ratio))
    max_allowed_h = max(min_h, min(int(metrics.avail_h * max_h_ratio), metrics.avail_h - margin_bottom))

    # Target dimensions: take ideal, clamped to max allowed and min size
    target_w = min(ideal_w, max_allowed_w)
    target_w = max(min(min_w, metrics.avail_w - 20), target_w)

    target_h = min(ideal_h, max_allowed_h)
    target_h = max(min(min_h, metrics.avail_h - margin_bottom), target_h)

    # Centering position within available work area
    x = metrics.avail_x + max(0, (metrics.avail_w - target_w) // 2)
    y = metrics.avail_y + max(0, (metrics.avail_h - target_h) // 2)

    # Boundary guards: Ensure strict containment within availableGeometry
    if y + target_h > (metrics.avail_y + metrics.avail_h - margin_bottom):
        y = max(metrics.avail_y, metrics.avail_y + metrics.avail_h - target_h - margin_bottom)
    if x + target_w > (metrics.avail_x + metrics.avail_w):
        x = max(metrics.avail_x, metrics.avail_x + metrics.avail_w - target_w)

    return x, y, target_w, target_h, metrics.is_compact


def apply_smart_dialog_geometry(
    dialog: QDialog,
    ideal_w: int = 880,
    ideal_h: int = 680,
    min_w: int = 640,
    min_h: int = 380,
    max_w_ratio: float = 0.94,
    max_h_ratio: float = 0.86,
    margin_bottom: int = 16,
) -> bool:
    """Apply responsive dimensions, set safe min/max bounds, and center dialog inside available work area.

    Returns:
        bool: True if the current screen is in compact mode (laptop / scaled DPI).
    """
    metrics = get_screen_metrics(dialog)
    x, y, target_w, target_h, is_compact = compute_smart_geometry(
        widget_or_parent=dialog,
        ideal_w=ideal_w,
        ideal_h=ideal_h,
        min_w=min_w,
        min_h=min_h,
        max_w_ratio=max_w_ratio,
        max_h_ratio=max_h_ratio,
        margin_bottom=margin_bottom,
    )

    safe_min_w = min(min_w, target_w, max(300, metrics.avail_w - 20))
    safe_min_h = min(min_h, target_h, max(260, metrics.avail_h - margin_bottom))

    try:
        dialog.setMinimumSize(safe_min_w, safe_min_h)
        dialog.setMaximumSize(metrics.avail_w, metrics.avail_h)
        dialog.setGeometry(x, y, target_w, target_h)
    except Exception as exc:
        logger.warning("Could not set smart dialog geometry: %s", exc)

    return is_compact


def ensure_window_fits_screen(window: QWidget, margin_bottom: int = 16) -> None:
    """Inspect current window geometry and clamp it if any portion extends beyond the work area."""
    metrics = get_screen_metrics(window)
    current_geom = window.geometry()

    new_x = current_geom.x()
    new_y = current_geom.y()
    new_w = current_geom.width()
    new_h = current_geom.height()

    changed = False

    # Height overflow check
    max_allowed_h = metrics.avail_h - margin_bottom
    if new_h > max_allowed_h:
        new_h = max_allowed_h
        changed = True

    # Width overflow check
    if new_w > metrics.avail_w:
        new_w = metrics.avail_w
        changed = True

    # Bottom boundary check (over taskbar)
    max_bottom = metrics.avail_y + metrics.avail_h - margin_bottom
    if new_y + new_h > max_bottom:
        new_y = max(metrics.avail_y, max_bottom - new_h)
        changed = True

    # Top boundary check
    if new_y < metrics.avail_y:
        new_y = metrics.avail_y
        changed = True

    # Right boundary check
    max_right = metrics.avail_x + metrics.avail_w
    if new_x + new_w > max_right:
        new_x = max(metrics.avail_x, max_right - new_w)
        changed = True

    # Left boundary check
    if new_x < metrics.avail_x:
        new_x = metrics.avail_x
        changed = True

    if changed:
        try:
            window.setGeometry(new_x, new_y, new_w, new_h)
        except Exception as exc:
            logger.warning("Could not adjust window bounds: %s", exc)
