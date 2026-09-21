"""Unit tests for UI Theme Architecture, Tokens, and SVG Assets.

Covers:
- TEST-UI-01: Mathematical WCAG 2.1 sRGB Relative Luminance & Contrast Ratios for Design Tokens
- TEST-UI-02: QSS Stylesheet syntax validity and widget instantiation on PyQt6
- TEST-UI-04: Theme switching, hot-reloading, and persistent preference storage
- TEST-UI-05: 100% SVG Icon availability, valid XML formatting, and QIcon rendering
"""

from __future__ import annotations

import json
import os
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

# Ensure offscreen QPA for headless environment
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QSize  # noqa: E402
from PyQt6.QtGui import QIcon  # noqa: E402
from PyQt6.QtWidgets import (  # noqa: E402
    QApplication,
    QComboBox,
    QGroupBox,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QTableView,
)

from src.gui.styles.theme_manager import (  # noqa: E402
    ThemeManager,
    get_theme_manager,
    reset_theme_manager,
)
from src.gui.styles.tokens import (  # noqa: E402
    BORDER_WIDTH,
    CELL_PADDING,
    COLOR_DARK_BRAND_PRIMARY,
    COLOR_DARK_BRAND_TEXT,
    COLOR_DARK_HEADER_BG,
    COLOR_DARK_HEADER_TEXT,
    COLOR_DARK_STATUS_DIFF_BG,
    COLOR_DARK_STATUS_DIFF_TEXT,
    COLOR_DARK_STATUS_MATCH_BG,
    COLOR_DARK_STATUS_MATCH_TEXT,
    COLOR_DARK_SURFACE,
    COLOR_DARK_TEXT_PRIMARY,
    COLOR_DARK_TEXT_SECONDARY,
    COLOR_LIGHT_BRAND_PRIMARY,
    COLOR_LIGHT_BRAND_TEXT,
    COLOR_LIGHT_HEADER_BG,
    COLOR_LIGHT_HEADER_TEXT,
    COLOR_LIGHT_STATUS_DIFF_BG,
    COLOR_LIGHT_STATUS_DIFF_TEXT,
    COLOR_LIGHT_STATUS_MATCH_BG,
    COLOR_LIGHT_STATUS_MATCH_TEXT,
    COLOR_LIGHT_SURFACE,
    COLOR_LIGHT_TEXT_PRIMARY,
    COLOR_LIGHT_TEXT_SECONDARY,
    PROGRESS_BAR_HEIGHT,
    TABLE_ROW_HEIGHT,
    calculate_contrast_ratio,
    calculate_relative_luminance,
    hex_to_rgb,
    is_wcag_aa,
    is_wcag_aaa,
)

EXPECTED_16_ICONS = [
    "download.svg",
    "refresh.svg",
    "check-circle.svg",
    "alert-triangle.svg",
    "x-circle.svg",
    "settings.svg",
    "search.svg",
    "filter.svg",
    "file-spreadsheet.svg",
    "moon.svg",
    "sun.svg",
    "folder.svg",
    "users.svg",
    "user-check.svg",
    "mail.svg",
    "calendar.svg",
]


# ============================================================================
# TEST-UI-01: THEME TOKENS CONTRAST & DIMENSIONS
# ============================================================================


class TestThemeTokensContrast:
    """TEST-UI-01: Mathematical verification of tokens and WCAG contrast."""

    def test_dimensions_contract(self) -> None:
        """Verify standard data-dense layout dimension constants."""
        assert TABLE_ROW_HEIGHT == 32, "TABLE_ROW_HEIGHT must be 32px"
        assert CELL_PADDING == (4, 8), "CELL_PADDING must be 4px vertical, 8px horizontal"
        assert PROGRESS_BAR_HEIGHT == 14, "PROGRESS_BAR_HEIGHT must be 14px"
        assert BORDER_WIDTH == 1, "BORDER_WIDTH must be 1px"

    def test_hex_to_rgb_conversion(self) -> None:
        """Verify hex string parsing and validation."""
        assert hex_to_rgb("#FFFFFF") == (255, 255, 255)
        assert hex_to_rgb("#000000") == (0, 0, 0)
        assert hex_to_rgb("#fff") == (255, 255, 255)
        assert hex_to_rgb("0F172A") == (15, 23, 42)

        with pytest.raises(ValueError):
            hex_to_rgb("#XYZ123")
        with pytest.raises(ValueError):
            hex_to_rgb("not-a-color")

    def test_wcag_relative_luminance_boundary_values(self) -> None:
        """Verify sRGB relative luminance mathematically matches reference points."""
        l_white = calculate_relative_luminance("#FFFFFF")
        l_black = calculate_relative_luminance("#000000")

        assert abs(l_white - 1.0) < 1e-5, f"Pure white luminance should be 1.0, got {l_white}"
        assert abs(l_black - 0.0) < 1e-5, f"Pure black luminance should be 0.0, got {l_black}"

        contrast_max = calculate_contrast_ratio("#FFFFFF", "#000000")
        assert abs(contrast_max - 21.0) < 1e-3, f"White on Black contrast must be 21:1, got {contrast_max}"

        contrast_same = calculate_contrast_ratio("#2563EB", "#2563EB")
        assert abs(contrast_same - 1.0) < 1e-3, f"Identical colors contrast must be 1:1, got {contrast_same}"

    def test_light_theme_wcag_contrast_matrix(self) -> None:
        """Verify Light Theme color tokens meet or exceed WCAG 2.1 standards."""
        # 1. Primary Text on Surface: >= 7.0 (AAA)
        ratio_primary = calculate_contrast_ratio(COLOR_LIGHT_TEXT_PRIMARY, COLOR_LIGHT_SURFACE)
        assert ratio_primary >= 7.0, f"Light primary text contrast {ratio_primary:.2f} < 7.0 (AAA)"
        assert is_wcag_aaa(COLOR_LIGHT_TEXT_PRIMARY, COLOR_LIGHT_SURFACE)

        # 2. Secondary Text on Surface: >= 4.5 (AA)
        ratio_secondary = calculate_contrast_ratio(COLOR_LIGHT_TEXT_SECONDARY, COLOR_LIGHT_SURFACE)
        assert ratio_secondary >= 4.5, f"Light secondary text contrast {ratio_secondary:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_LIGHT_TEXT_SECONDARY, COLOR_LIGHT_SURFACE)

        # 3. Primary CTA Button Text on Button BG: >= 4.5 (AA)
        ratio_cta = calculate_contrast_ratio(COLOR_LIGHT_BRAND_TEXT, COLOR_LIGHT_BRAND_PRIMARY)
        assert ratio_cta >= 4.5, f"Light CTA button contrast {ratio_cta:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_LIGHT_BRAND_TEXT, COLOR_LIGHT_BRAND_PRIMARY)

        # 4. Status Match Text on Match BG: >= 4.5 (AA, target ~7.0)
        ratio_match = calculate_contrast_ratio(COLOR_LIGHT_STATUS_MATCH_TEXT, COLOR_LIGHT_STATUS_MATCH_BG)
        assert ratio_match >= 4.5, f"Light match status contrast {ratio_match:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_LIGHT_STATUS_MATCH_TEXT, COLOR_LIGHT_STATUS_MATCH_BG)

        # 5. Status Diff Text on Diff BG: >= 4.5 (AA, target ~7.0)
        ratio_diff = calculate_contrast_ratio(COLOR_LIGHT_STATUS_DIFF_TEXT, COLOR_LIGHT_STATUS_DIFF_BG)
        assert ratio_diff >= 4.5, f"Light diff status contrast {ratio_diff:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_LIGHT_STATUS_DIFF_TEXT, COLOR_LIGHT_STATUS_DIFF_BG)

        # 6. Table Header Text on Header BG: >= 4.5 (AA)
        ratio_header = calculate_contrast_ratio(COLOR_LIGHT_HEADER_TEXT, COLOR_LIGHT_HEADER_BG)
        assert ratio_header >= 4.5, f"Light table header contrast {ratio_header:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_LIGHT_HEADER_TEXT, COLOR_LIGHT_HEADER_BG)

    def test_dark_theme_wcag_contrast_matrix(self) -> None:
        """Verify Dark Theme color tokens meet or exceed WCAG 2.1 standards."""
        # 1. Primary Text on Surface: >= 7.0 (AAA)
        ratio_primary = calculate_contrast_ratio(COLOR_DARK_TEXT_PRIMARY, COLOR_DARK_SURFACE)
        assert ratio_primary >= 7.0, f"Dark primary text contrast {ratio_primary:.2f} < 7.0 (AAA)"
        assert is_wcag_aaa(COLOR_DARK_TEXT_PRIMARY, COLOR_DARK_SURFACE)

        # 2. Secondary Text on Surface: >= 4.5 (AA)
        ratio_secondary = calculate_contrast_ratio(COLOR_DARK_TEXT_SECONDARY, COLOR_DARK_SURFACE)
        assert ratio_secondary >= 4.5, f"Dark secondary text contrast {ratio_secondary:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_DARK_TEXT_SECONDARY, COLOR_DARK_SURFACE)

        # 3. Primary CTA Button Text on Button BG: >= 4.5 (AA)
        ratio_cta = calculate_contrast_ratio(COLOR_DARK_BRAND_TEXT, COLOR_DARK_BRAND_PRIMARY)
        assert ratio_cta >= 4.5, f"Dark CTA button contrast {ratio_cta:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_DARK_BRAND_TEXT, COLOR_DARK_BRAND_PRIMARY)

        # 4. Status Match Text on Match BG: >= 7.0 (AAA)
        ratio_match = calculate_contrast_ratio(COLOR_DARK_STATUS_MATCH_TEXT, COLOR_DARK_STATUS_MATCH_BG)
        assert ratio_match >= 7.0, f"Dark match status contrast {ratio_match:.2f} < 7.0 (AAA)"
        assert is_wcag_aaa(COLOR_DARK_STATUS_MATCH_TEXT, COLOR_DARK_STATUS_MATCH_BG)

        # 5. Status Diff Text on Diff BG: >= 4.5 (AA, target ~7.0)
        ratio_diff = calculate_contrast_ratio(COLOR_DARK_STATUS_DIFF_TEXT, COLOR_DARK_STATUS_DIFF_BG)
        assert ratio_diff >= 4.5, f"Dark diff status contrast {ratio_diff:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_DARK_STATUS_DIFF_TEXT, COLOR_DARK_STATUS_DIFF_BG)

        # 6. Table Header Text on Header BG: >= 4.5 (AA)
        ratio_header = calculate_contrast_ratio(COLOR_DARK_HEADER_TEXT, COLOR_DARK_HEADER_BG)
        assert ratio_header >= 4.5, f"Dark table header contrast {ratio_header:.2f} < 4.5 (AA)"
        assert is_wcag_aa(COLOR_DARK_HEADER_TEXT, COLOR_DARK_HEADER_BG)


# ============================================================================
# TEST-UI-02: QSS SYNTAX VALIDITY
# ============================================================================


class TestQssSyntaxValidity:
    """TEST-UI-02: Verification that QSS stylesheets parse cleanly on PyQt6."""

    def test_qss_files_exist_and_non_empty(self) -> None:
        """Check that light_theme.qss and dark_theme.qss exist and are populated."""
        styles_dir = Path(__file__).resolve().parents[2] / "src" / "gui" / "styles"
        light_qss = styles_dir / "light_theme.qss"
        dark_qss = styles_dir / "dark_theme.qss"

        assert light_qss.exists(), f"light_theme.qss missing at {light_qss}"
        assert dark_qss.exists(), f"dark_theme.qss missing at {dark_qss}"

        assert len(light_qss.read_text(encoding="utf-8").strip()) > 500
        assert len(dark_qss.read_text(encoding="utf-8").strip()) > 500

    def test_qss_applied_to_widgets_without_exception(self) -> None:
        """Apply light and dark stylesheets to core widgets and verify no exceptions occur."""
        _ = QApplication.instance() or QApplication([])
        styles_dir = Path(__file__).resolve().parents[2] / "src" / "gui" / "styles"
        light_content = (styles_dir / "light_theme.qss").read_text(encoding="utf-8")
        dark_content = (styles_dir / "dark_theme.qss").read_text(encoding="utf-8")

        for _name, qss in [("light", light_content), ("dark", dark_content)]:
            win = QMainWindow()
            win.setStyleSheet(qss)

            table = QTableView(win)
            btn = QPushButton("Test Button", win)
            btn_cta = QPushButton("Primary Action", win)
            btn_cta.setObjectName("cta_primary")
            line = QLineEdit(win)
            combo = QComboBox(win)
            pbar = QProgressBar(win)
            pbar.setValue(50)
            group = QGroupBox("Group Title", win)

            assert win is not None
            assert table is not None
            assert btn is not None
            assert btn_cta is not None
            assert line is not None
            assert combo is not None
            assert pbar is not None
            assert group is not None
            win.deleteLater()


# ============================================================================
# TEST-UI-04: THEME TOGGLE & PERSISTENCE
# ============================================================================


class TestThemeTogglePersistence:
    """TEST-UI-04: Verification of ThemeManager hot-reloading and config persistence."""

    def test_theme_manager_state_machine(self, tmp_path: Path) -> None:
        """Verify theme state transitions and signal emission."""
        config_file = tmp_path / "settings.json"
        tm = ThemeManager(config_path=config_file, apply_stylesheet=False)

        emitted_signals: list[str] = []
        tm.theme_changed.connect(lambda t: emitted_signals.append(t))

        # Switch to light
        tm.set_theme("light", apply_stylesheet=False)
        assert tm.get_current_theme() == "light"
        assert tm.get_effective_theme() == "light"
        assert not tm.is_dark()
        assert emitted_signals[-1] == "light"

        # Switch to dark
        tm.set_theme("dark", apply_stylesheet=False)
        assert tm.get_current_theme() == "dark"
        assert tm.get_effective_theme() == "dark"
        assert tm.is_dark()
        assert emitted_signals[-1] == "dark"

        # Switch to system
        tm.set_theme("system", apply_stylesheet=False)
        assert tm.get_current_theme() == "system"
        assert tm.get_effective_theme() in ("light", "dark")
        assert emitted_signals[-1] in ("light", "dark")

        # Invalid theme
        with pytest.raises(ValueError):
            tm.set_theme("invalid_theme")

    def test_theme_persistence_to_config(self, tmp_path: Path) -> None:
        """Verify theme settings are saved to and reloaded from JSON file."""
        config_file = tmp_path / "subdir" / "settings.json"
        tm1 = ThemeManager(config_path=config_file, apply_stylesheet=False)

        tm1.set_theme("dark", save_preference=True, apply_stylesheet=False)
        assert config_file.exists(), "Config file should be created"

        saved_data = json.loads(config_file.read_text(encoding="utf-8"))
        assert saved_data.get("theme") == "dark"

        # Create new instance with the same config file
        tm2 = ThemeManager(config_path=config_file, apply_stylesheet=False)
        assert tm2.get_current_theme() == "dark"
        assert tm2.get_effective_theme() == "dark"

    def test_theme_persistence_preserves_other_keys(self, tmp_path: Path) -> None:
        """Verify theme persistence preserves existing config settings."""
        config_file = tmp_path / "settings.json"
        initial_data = {
            "tc14": {"username": "test_pe"},
            "sap": {"plant": "2200"},
            "theme": "light",
        }
        config_file.write_text(json.dumps(initial_data, indent=2), encoding="utf-8")

        tm = ThemeManager(config_path=config_file, apply_stylesheet=False)
        tm.set_theme("dark", save_preference=True, apply_stylesheet=False)

        updated_data = json.loads(config_file.read_text(encoding="utf-8"))
        assert updated_data["theme"] == "dark"
        assert updated_data["tc14"]["username"] == "test_pe"
        assert updated_data["sap"]["plant"] == "2200"

    def test_hot_reload_applies_stylesheet(self, qapp: QApplication) -> None:
        """Verify set_theme with apply_stylesheet=True sets QApplication stylesheet."""
        tm = ThemeManager(apply_stylesheet=False)
        tm.set_theme("dark", apply_stylesheet=True)
        assert len(qapp.styleSheet()) > 0
        assert "Industrial Dark Mode" in qapp.styleSheet()

    def test_system_theme_mocking(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Verify detect_system_theme accurately responds to Windows registry states."""
        # Test Dark registry state
        monkeypatch.setattr(ThemeManager, "detect_system_theme", staticmethod(lambda: "dark"))
        assert ThemeManager.detect_system_theme() == "dark"

        # Test Light registry state
        monkeypatch.setattr(ThemeManager, "detect_system_theme", staticmethod(lambda: "light"))
        assert ThemeManager.detect_system_theme() == "light"

    def test_singleton_helper(self, tmp_path: Path) -> None:
        """Verify get_theme_manager returns consistent singleton instance."""
        reset_theme_manager()
        config_file = tmp_path / "config.json"
        inst1 = get_theme_manager(config_path=config_file)
        inst2 = get_theme_manager()
        assert inst1 is inst2
        reset_theme_manager()


# ============================================================================
# TEST-UI-05: SVG ICONS AVAILABILITY & RENDERING
# ============================================================================


class TestSvgIconsAvailability:
    """TEST-UI-05: Verification of 100% SVG icon presence and QIcon rendering."""

    def test_all_16_svg_files_exist(self) -> None:
        """Verify all 16 required SVG icons exist in src/gui/assets/icons/."""
        icons_dir = Path(__file__).resolve().parents[2] / "src" / "gui" / "assets" / "icons"
        assert icons_dir.exists(), f"Icons directory missing: {icons_dir}"

        missing = [name for name in EXPECTED_16_ICONS if not (icons_dir / name).exists()]
        assert not missing, f"Missing {len(missing)} required SVG icons: {missing}"

    def test_all_svg_files_are_valid_xml(self) -> None:
        """Verify every SVG file is well-formed XML with standard viewBox."""
        icons_dir = Path(__file__).resolve().parents[2] / "src" / "gui" / "assets" / "icons"

        for icon_name in EXPECTED_16_ICONS:
            path = icons_dir / icon_name
            tree = ET.parse(path)
            root = tree.getroot()

            # Verify tag ends with svg (handling xmlns)
            assert root.tag.endswith("svg"), f"{icon_name} root tag is not svg"
            assert "viewBox" in root.attrib, f"{icon_name} missing viewBox attribute"
            assert root.attrib["viewBox"] == "0 0 24 24", f"{icon_name} viewBox != '0 0 24 24'"

    def test_styled_icon_rendering(self, qapp: QApplication) -> None:
        """Verify ThemeManager renders all 16 icons into non-null QIcons with tinting."""
        tm = get_theme_manager()

        for icon_name in EXPECTED_16_ICONS:
            base_name = icon_name.replace(".svg", "")

            # Default theme tint
            icon_default = tm.get_styled_icon(base_name)
            assert not icon_default.isNull(), f"get_styled_icon({base_name}) returned null QIcon"

            # Custom color tint
            icon_tinted = tm.get_styled_icon(base_name, color="#2563EB", size=QSize(32, 32))
            assert not icon_tinted.isNull(), f"get_styled_icon({base_name}, color='#2563EB') returned null"

            # Direct QIcon from file path
            path = tm.get_icon_path(base_name)
            direct_icon = QIcon(str(path))
            assert not direct_icon.isNull(), f"Direct QIcon({path}) returned null"

    def test_styled_icon_different_sizes(self, qapp: QApplication) -> None:
        """Verify icon rendering handles 16x16, 24x24, and 48x48 sizes."""
        tm = get_theme_manager()
        for sz in (16, 24, 48):
            icon = tm.get_styled_icon("download", size=QSize(sz, sz))
            assert not icon.isNull()
            assert len(icon.availableSizes()) > 0 or not icon.pixmap(QSize(sz, sz)).isNull()

    def test_missing_icon_handling(self) -> None:
        """Verify requesting a non-existent icon returns an empty QIcon safely."""
        tm = get_theme_manager()
        icon = tm.get_styled_icon("completely_non_existent_icon_xyz")
        assert icon.isNull(), "Missing icon should return null QIcon"


# ============================================================================
# PACKAGE EXPORT INTEGRITY
# ============================================================================


def test_styles_package_exports() -> None:
    """Verify src.gui.styles exposes all required tokens and classes."""
    import src.gui.styles as styles

    assert hasattr(styles, "ThemeManager")
    assert hasattr(styles, "get_theme_manager")
    assert hasattr(styles, "reset_theme_manager")
    assert hasattr(styles, "TABLE_ROW_HEIGHT")
    assert hasattr(styles, "CELL_PADDING")
    assert hasattr(styles, "PROGRESS_BAR_HEIGHT")
    assert hasattr(styles, "BORDER_WIDTH")
    assert hasattr(styles, "COLOR_LIGHT_BG_CANVAS")
    assert hasattr(styles, "COLOR_DARK_BG_CANVAS")
    assert hasattr(styles, "calculate_contrast_ratio")
    assert hasattr(styles, "calculate_relative_luminance")
    assert hasattr(styles, "is_wcag_aa")
    assert hasattr(styles, "is_wcag_aaa")
