"""ThemeManager for SSBOM Application.

Coordinates centralized theming, stylesheet hot-reloading, SVG iconography tinting,
and persistent theme configuration between Light (Slate Industrial) and Dark (Industrial Dark Mode).
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QByteArray, QObject, QSize, Qt, pyqtSignal
from PyQt6.QtGui import QIcon, QPainter, QPixmap
from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtWidgets import QApplication

from src.gui.styles.tokens import (
    COLOR_DARK_TEXT_PRIMARY,
    COLOR_LIGHT_TEXT_PRIMARY,
)

logger = logging.getLogger(__name__)

DEFAULT_CONFIG_PATH = Path("config/settings.json")
FALLBACK_CONFIG_PATH = Path.home() / ".ssbom" / "config.json"


class ThemeManager(QObject):
    """Centralized Theme and Icon Manager with Qt6 Hot-Reload support."""

    theme_changed = pyqtSignal(str)

    def __init__(self, config_path: Path | str | None = None, apply_stylesheet: bool = False) -> None:
        super().__init__()
        if config_path is not None:
            self._config_path = Path(config_path)
        elif FALLBACK_CONFIG_PATH.exists():
            self._config_path = FALLBACK_CONFIG_PATH
        else:
            self._config_path = DEFAULT_CONFIG_PATH

        self._styles_dir = Path(__file__).resolve().parent
        self._icons_dir = self._styles_dir.parent / "assets" / "icons"

        self._current_theme: str = "light"
        self._effective_theme: str = "light"

        # Load saved theme preference
        saved = self.load_saved_theme()
        self.set_theme(saved, save_preference=False, apply_stylesheet=apply_stylesheet)

    @property
    def icons_dir(self) -> Path:
        """Directory where SVG icons reside."""
        return self._icons_dir

    @property
    def current_theme(self) -> str:
        """Get current user-selected theme setting ('light', 'dark', or 'system')."""
        return self._current_theme

    @property
    def effective_theme(self) -> str:
        """Get effective resolved theme applied ('light' or 'dark')."""
        return self._effective_theme

    def get_current_theme(self) -> str:
        """Get current user-selected theme setting ('light', 'dark', or 'system')."""
        return self._current_theme

    def get_effective_theme(self) -> str:
        """Get effective resolved theme applied ('light' or 'dark')."""
        return self._effective_theme

    def is_dark(self) -> bool:
        """Return True if currently active effective theme is dark."""
        return self._effective_theme == "dark"

    @staticmethod
    def detect_system_theme() -> str:
        """Detect whether host OS is in Dark or Light mode.

        On Windows, queries the 'AppsUseLightTheme' registry key.
        Defaults to 'light' on error or non-Windows platforms.
        """
        if sys.platform == "win32":
            try:
                import winreg

                key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    return "light" if value != 0 else "dark"
            except Exception as e:
                logger.debug(f"Could not read Windows theme registry: {e}")
        return "light"

    def load_saved_theme(self) -> str:
        """Load saved theme preference from config file, defaulting to 'light'."""
        if self._config_path.exists():
            try:
                data = json.loads(self._config_path.read_text(encoding="utf-8"))
                if isinstance(data, dict) and "theme" in data:
                    theme = data["theme"].strip().lower()
                    if theme in ("light", "dark", "system"):
                        return theme
            except Exception as e:
                logger.warning(f"Failed to read theme config from {self._config_path}: {e}")
        return "light"

    def save_theme(self, theme_name: str) -> None:
        """Persist theme preference to config file while preserving other config keys."""
        try:
            self._config_path.parent.mkdir(parents=True, exist_ok=True)
            existing: dict[str, Any] = {}
            if self._config_path.exists():
                try:
                    existing = json.loads(self._config_path.read_text(encoding="utf-8"))
                    if not isinstance(existing, dict):
                        existing = {}
                except Exception:
                    existing = {}

            existing["theme"] = theme_name
            self._config_path.write_text(json.dumps(existing, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            logger.warning(f"Failed to save theme setting to {self._config_path}: {e}")

    def get_stylesheet(self, theme_name: str | None = None) -> str:
        """Get stylesheet content for given theme ('light' or 'dark').

        If not specified, uses the current effective theme.
        """
        target = theme_name or self._effective_theme
        if target == "system":
            target = self.detect_system_theme()

        file_name = "dark_theme.qss" if target == "dark" else "light_theme.qss"
        qss_path = self._styles_dir / file_name
        if qss_path.exists():
            return qss_path.read_text(encoding="utf-8")
        logger.error(f"QSS stylesheet file not found: {qss_path}")
        return ""

    def set_theme(
        self,
        theme_name: str,
        save_preference: bool = True,
        apply_stylesheet: bool = True,
    ) -> None:
        """Set theme, hot-reload active QApplication stylesheet, and emit signal.

        Args:
            theme_name: 'light', 'dark', or 'system'
            save_preference: If True, writes setting to config file
            apply_stylesheet: If True, applies stylesheet to active QApplication
        """
        theme_clean = theme_name.strip().lower()
        if theme_clean not in ("light", "dark", "system"):
            raise ValueError(f"Unknown theme '{theme_name}'. Expected 'light', 'dark', or 'system'.")

        self._current_theme = theme_clean
        if theme_clean == "system":
            self._effective_theme = self.detect_system_theme()
        else:
            self._effective_theme = theme_clean

        if save_preference:
            self.save_theme(self._current_theme)

        # Apply to running QApplication only if requested and changed
        if apply_stylesheet:
            app = QApplication.instance()
            if app is not None and isinstance(app, QApplication):
                qss = self.get_stylesheet(self._effective_theme)
                if app.styleSheet() != qss:
                    app.setStyleSheet(qss)

        self.theme_changed.emit(self._effective_theme)
        logger.info(f"Theme switched to: {self._current_theme} (effective: {self._effective_theme})")

    def apply_theme_to_app(self) -> None:
        """Apply the current effective stylesheet to active QApplication."""
        app = QApplication.instance()
        if app is not None and isinstance(app, QApplication):
            app.setStyleSheet(self.get_stylesheet(self._effective_theme))

    def get_icon_path(self, icon_name: str) -> Path:
        """Return Path to an SVG icon by base name."""
        name = icon_name if icon_name.endswith(".svg") else f"{icon_name}.svg"
        return self._icons_dir / name

    def get_styled_icon(
        self,
        icon_name: str,
        color: str | None = None,
        size: QSize | None = None,
    ) -> QIcon:
        """Load and tint an SVG icon with the specified or theme-appropriate color.

        Args:
            icon_name: Base name of SVG icon (e.g. 'download' or 'download.svg')
            color: Hex color string to tint (e.g. '#2563EB'). Defaults to theme's primary text.
            size: Target QSize for rendering (defaults to 24x24).

        Returns:
            QIcon rendered with the requested tint, or empty QIcon on missing file.
        """
        svg_path = self.get_icon_path(icon_name)
        if not svg_path.exists():
            logger.warning(f"SVG icon not found: {svg_path}")
            return QIcon()

        target_color = color or (COLOR_DARK_TEXT_PRIMARY if self.is_dark() else COLOR_LIGHT_TEXT_PRIMARY)

        try:
            svg_content = svg_path.read_text(encoding="utf-8")
            # Replace currentColor with requested tint color
            tinted = svg_content.replace("currentColor", target_color)

            renderer = QSvgRenderer(QByteArray(tinted.encode("utf-8")))
            if not renderer.isValid():
                return QIcon(str(svg_path))

            target_size = size or QSize(24, 24)
            pixmap = QPixmap(target_size)
            pixmap.fill(Qt.GlobalColor.transparent)

            painter = QPainter(pixmap)
            renderer.render(painter)
            painter.end()

            return QIcon(pixmap)
        except Exception as e:
            logger.warning(f"Failed to render styled icon '{icon_name}': {e}")
            return QIcon(str(svg_path))


_global_theme_manager: ThemeManager | None = None


def get_theme_manager(
    config_path: Path | str | None = None,
    apply_stylesheet: bool = False,
) -> ThemeManager:
    """Retrieve or create singleton ThemeManager instance."""
    global _global_theme_manager
    if _global_theme_manager is None:
        _global_theme_manager = ThemeManager(config_path=config_path, apply_stylesheet=apply_stylesheet)
    elif config_path is not None and _global_theme_manager._config_path != Path(config_path):
        _global_theme_manager = ThemeManager(config_path=config_path, apply_stylesheet=apply_stylesheet)
    return _global_theme_manager


def reset_theme_manager() -> None:
    """Reset global ThemeManager instance (useful for test isolation)."""
    global _global_theme_manager
    _global_theme_manager = None
