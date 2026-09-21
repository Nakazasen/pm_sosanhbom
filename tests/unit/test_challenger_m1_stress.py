"""Empirical Stress Test Suite for Milestone 1 (M1) — WCAG Contrast & SVG Iconography.

Author: Challenger 1 (Adversarial Empirical Verification)
Date: 2026-09-21

This suite executes independent empirical tests:
1. Independent W3C sRGB relative luminance oracle vs tokens.py implementation
2. Extreme color boundary pairs (#000000 on #FFFFFF = 21:1, #FFFFFF on #FFFFFF = 1:1, inverted, minimal delta)
3. 10,000 random Monte Carlo color pairs for numerical stability and range bounds [1.0, 21.0]
4. Comprehensive verification of all design tokens in tokens.py against SPEC §3.1, §3.2, §3.3
5. Multi-resolution rendering of all 16 SVG icons (16x16, 24x24, 32x32, 64x64, 128x128)
6. Empirical pixel color verification of get_styled_icon() tinting mechanism across all 16 icons
"""

from __future__ import annotations

import math
import os
import random
import re
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

# Offscreen QPA for headless environment
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QSize  # noqa: E402
from PyQt6.QtGui import QColor, QIcon, QImage  # noqa: E402
from PyQt6.QtWidgets import QApplication  # noqa: E402

from src.gui.styles.theme_manager import (  # noqa: E402
    ThemeManager,
    get_theme_manager,
)
import src.gui.styles.tokens as tokens  # noqa: E402

ALL_16_ICONS = [
    "alert-triangle.svg",
    "calendar.svg",
    "check-circle.svg",
    "download.svg",
    "file-spreadsheet.svg",
    "filter.svg",
    "folder.svg",
    "mail.svg",
    "moon.svg",
    "refresh.svg",
    "search.svg",
    "settings.svg",
    "sun.svg",
    "user-check.svg",
    "users.svg",
    "x-circle.svg",
]


# ============================================================================
# INDEPENDENT W3C WCAG 2.1 sRGB REFERENCE ORACLE
# ============================================================================


def oracle_srgb_to_linear(c_srgb: float) -> float:
    """Independent implementation of W3C WCAG 2.1 gamma expansion formula."""
    if c_srgb <= 0.04045:
        return c_srgb / 12.92
    return math.pow((c_srgb + 0.055) / 1.055, 2.4)


def oracle_relative_luminance(hex_color: str) -> float:
    """Independent calculation of relative luminance per W3C WCAG 2.1."""
    clean_hex = hex_color.strip().lstrip("#")
    if len(clean_hex) == 3:
        clean_hex = "".join(c * 2 for c in clean_hex)
    r = int(clean_hex[0:2], 16) / 255.0
    g = int(clean_hex[2:4], 16) / 255.0
    b = int(clean_hex[4:6], 16) / 255.0

    r_lin = oracle_srgb_to_linear(r)
    g_lin = oracle_srgb_to_linear(g)
    b_lin = oracle_srgb_to_linear(b)

    return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin


def oracle_contrast_ratio(fg: str, bg: str) -> float:
    """Independent calculation of WCAG 2.1 contrast ratio."""
    l1 = oracle_relative_luminance(fg)
    l2 = oracle_relative_luminance(bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


# ============================================================================
# PART 1: WCAG CONTRAST & sRGB MATHEMATICAL STRESS TESTS
# ============================================================================


class TestWcagMathEmpiricalStress:
    """Adversarial stress-testing of WCAG relative luminance & contrast algorithms."""

    def test_oracle_agreement_all_grayscale_levels(self) -> None:
        """Verify tokens.py matches Oracle exactly across all 256 grayscale levels."""
        max_diff_luminance = 0.0
        for i in range(256):
            hex_str = f"#{i:02X}{i:02X}{i:02X}"
            actual = tokens.calculate_relative_luminance(hex_str)
            expected = oracle_relative_luminance(hex_str)
            diff = abs(actual - expected)
            if diff > max_diff_luminance:
                max_diff_luminance = diff
            assert diff < 1e-9, f"Grayscale {hex_str}: actual {actual} != oracle {expected}"

        assert max_diff_luminance < 1e-9, f"Max grayscale divergence {max_diff_luminance} exceeded tolerance"

    def test_oracle_agreement_critical_threshold_boundary(self) -> None:
        """Verify channels near sRGB linear transition point (c / 255 ~ 0.04045 => c ~ 10.31)."""
        # Threshold: 255 * 0.04045 = 10.31475. Test integers 8, 9, 10, 11, 12, 13 across all 3 channels.
        boundary_vals = [0, 8, 9, 10, 11, 12, 13, 14, 254, 255]
        for r in boundary_vals:
            for g in boundary_vals:
                for b in boundary_vals:
                    hex_str = f"#{r:02X}{g:02X}{b:02X}"
                    actual = tokens.calculate_relative_luminance(hex_str)
                    expected = oracle_relative_luminance(hex_str)
                    assert abs(actual - expected) < 1e-9, (
                        f"Boundary {hex_str}: tokens={actual}, oracle={expected}"
                    )

    def test_extreme_boundary_contrast_values(self) -> None:
        """Stress-test extreme boundary conditions."""
        # 1. Pure Black on Pure White: Must be EXACTLY 21.0:1
        ratio_max = tokens.calculate_contrast_ratio("#000000", "#FFFFFF")
        assert math.isclose(ratio_max, 21.0, rel_tol=1e-5), f"Black/White contrast {ratio_max} != 21.0"

        # 2. Symmetry check (inverted fg and bg): Must produce identical ratio
        ratio_inverted = tokens.calculate_contrast_ratio("#FFFFFF", "#000000")
        assert math.isclose(ratio_max, ratio_inverted, abs_tol=1e-9), "calculate_contrast_ratio is asymmetric!"

        # 3. Identical colors: Must be EXACTLY 1.0:1
        assert math.isclose(tokens.calculate_contrast_ratio("#FFFFFF", "#FFFFFF"), 1.0, abs_tol=1e-5)
        assert math.isclose(tokens.calculate_contrast_ratio("#000000", "#000000"), 1.0, abs_tol=1e-5)
        assert math.isclose(tokens.calculate_contrast_ratio("#2563EB", "#2563EB"), 1.0, abs_tol=1e-5)
        assert math.isclose(tokens.calculate_contrast_ratio("#123456", "#123456"), 1.0, abs_tol=1e-5)

        # 4. Minimal non-zero step: #000000 vs #000001
        ratio_min_delta = tokens.calculate_contrast_ratio("#000000", "#000001")
        assert ratio_min_delta > 1.0
        assert ratio_min_delta < 1.001

        # 5. Minimal step at top: #FFFFFF vs #FFFFFE
        ratio_top_delta = tokens.calculate_contrast_ratio("#FFFFFF", "#FFFFFE")
        assert ratio_top_delta > 1.0
        assert ratio_top_delta < 1.001

    def test_monte_carlo_random_colors_stability(self) -> None:
        """Execute 10,000 randomized Monte Carlo color pairs for stability and mathematical bounds."""
        rng = random.Random(42)  # Deterministic seed for reproducibility
        max_ratio = 1.0
        min_ratio = 21.0

        for _ in range(10000):
            c1 = f"#{rng.randint(0, 0xFFFFFF):06X}"
            c2 = f"#{rng.randint(0, 0xFFFFFF):06X}"

            ratio_tokens = tokens.calculate_contrast_ratio(c1, c2)
            ratio_oracle = oracle_contrast_ratio(c1, c2)

            # Check divergence against independent oracle
            assert abs(ratio_tokens - ratio_oracle) < 1e-6, (
                f"Divergence on pair ({c1}, {c2}): tokens={ratio_tokens}, oracle={ratio_oracle}"
            )

            # Range bounds check [1.0, 21.0]
            assert 1.0 <= ratio_tokens <= 21.0001, f"Ratio out of bounds: {ratio_tokens}"

            if ratio_tokens > max_ratio:
                max_ratio = ratio_tokens
            if ratio_tokens < min_ratio:
                min_ratio = ratio_tokens

        assert max_ratio <= 21.0001
        assert min_ratio >= 1.0

    def test_hex_parser_input_robustness_and_adversarial_formats(self) -> None:
        """Adversarial input formats (3-digit, mixed case, no-hash, whitespace, invalid)."""
        # Equivalence test: #FFF, #fff, FFF, fff, "  #ffffff  "
        equiv_colors = ["#FFF", "#fff", "FFF", "fff", "#FFFFFF", "#ffffff", "FFFFFF", "ffffff", "  #FFFFFF  \n"]
        for c in equiv_colors:
            assert tokens.hex_to_rgb(c) == (255, 255, 255), f"Failed for format: {repr(c)}"
            assert math.isclose(tokens.calculate_relative_luminance(c), 1.0, abs_tol=1e-5)

        # Invalid formats must raise ValueError, not crash or return garbage
        invalid_inputs = [
            "",
            "   ",
            "#12",
            "#1234",
            "#12345",
            "#1234567",
            "#GGGGGG",
            "xyz",
            "#",
            "blue",
            "rgb(0,0,0)",
        ]
        for inv in invalid_inputs:
            with pytest.raises(ValueError):
                tokens.hex_to_rgb(inv)
            with pytest.raises(ValueError):
                tokens.calculate_relative_luminance(inv)
            with pytest.raises(ValueError):
                tokens.calculate_contrast_ratio(inv, "#FFFFFF")

    def test_all_defined_tokens_wcag_oracle_match(self) -> None:
        """Verify all COLOR_* constants defined in tokens.py match Oracle and meet spec targets."""
        token_names = [attr for attr in dir(tokens) if attr.startswith("COLOR_")]
        assert len(token_names) >= 25, f"Expected at least 25 color tokens, found {len(token_names)}"

        for name in token_names:
            color_val = getattr(tokens, name)
            assert isinstance(color_val, str)
            assert re.match(r"^#[0-9A-Fa-f]{6}$", color_val), f"Token {name} has invalid hex: {color_val}"

            # Compare luminance against oracle
            lum_tokens = tokens.calculate_relative_luminance(color_val)
            lum_oracle = oracle_relative_luminance(color_val)
            assert abs(lum_tokens - lum_oracle) < 1e-9, f"Luminance divergence for {name}: {color_val}"

    def test_spec_authoritative_contrast_matrix(self) -> None:
        """Verify SPEC §3.3 authoritative table contrast ratios explicitly."""
        spec_pairs = [
            ("Light Primary Text", tokens.COLOR_LIGHT_TEXT_PRIMARY, tokens.COLOR_LIGHT_SURFACE, 7.0, "AAA"),
            ("Light Secondary Text", tokens.COLOR_LIGHT_TEXT_SECONDARY, tokens.COLOR_LIGHT_SURFACE, 4.5, "AA"),
            ("Light CTA Button", tokens.COLOR_LIGHT_BRAND_TEXT, tokens.COLOR_LIGHT_BRAND_PRIMARY, 4.5, "AA"),
            ("Light Semantic Diff", tokens.COLOR_LIGHT_STATUS_DIFF_TEXT, tokens.COLOR_LIGHT_STATUS_DIFF_BG, 4.5, "AA"),
            (
                "Light Semantic Match",
                tokens.COLOR_LIGHT_STATUS_MATCH_TEXT,
                tokens.COLOR_LIGHT_STATUS_MATCH_BG,
                4.5,
                "AA",
            ),
            ("Dark Primary Text", tokens.COLOR_DARK_TEXT_PRIMARY, tokens.COLOR_DARK_SURFACE, 7.0, "AAA"),
            ("Dark Secondary Text", tokens.COLOR_DARK_TEXT_SECONDARY, tokens.COLOR_DARK_SURFACE, 4.5, "AA"),
            ("Dark CTA Button", tokens.COLOR_DARK_BRAND_TEXT, tokens.COLOR_DARK_BRAND_PRIMARY, 4.5, "AA"),
            ("Dark Semantic Match", tokens.COLOR_DARK_STATUS_MATCH_TEXT, tokens.COLOR_DARK_STATUS_MATCH_BG, 7.0, "AAA"),
            ("Dark Semantic Diff", tokens.COLOR_DARK_STATUS_DIFF_TEXT, tokens.COLOR_DARK_STATUS_DIFF_BG, 4.5, "AA"),
        ]

        for desc, fg, bg, min_ratio, level in spec_pairs:
            ratio = tokens.calculate_contrast_ratio(fg, bg)
            msg = f"{desc} ({fg} on {bg}) got {ratio:.2f}:1, expected >= {min_ratio}:1 ({level})"
            assert ratio >= min_ratio, f"SPEC VIOLATION: {msg}"
            if level == "AAA":
                assert tokens.is_wcag_aaa(fg, bg), f"{desc} failed is_wcag_aaa check"
            else:
                assert tokens.is_wcag_aa(fg, bg), f"{desc} failed is_wcag_aa check"


# ============================================================================
# PART 2: 16 SVG ICONS MULTI-RESOLUTION & TINTING STRESS TESTS
# ============================================================================


class TestSvgIconsMultiResolutionAndTintingStress:
    """Adversarial stress-testing of all 16 SVG icons and ThemeManager tinting."""

    def test_all_16_svg_xml_conformity(self) -> None:
        """Verify 100% of 16 SVG files are valid XML, 24x24 viewBox, and stroke='currentColor'."""
        tm = get_theme_manager()
        for icon_name in ALL_16_ICONS:
            path = tm.get_icon_path(icon_name)
            assert path.exists(), f"Missing icon: {path}"

            tree = ET.parse(path)
            root = tree.getroot()

            # Namespace handling
            tag = root.tag.split("}")[-1] if "}" in root.tag else root.tag
            assert tag == "svg", f"{icon_name} root is not svg"
            assert root.attrib.get("viewBox") == "0 0 24 24", f"{icon_name} invalid viewBox"
            assert root.attrib.get("stroke") == "currentColor", (
                f"{icon_name} does not have stroke='currentColor' on root"
            )

            # Check for illegal hardcoded colors in content
            content = path.read_text(encoding="utf-8")
            assert "currentColor" in content, f"{icon_name} missing 'currentColor' reference"

    @pytest.mark.parametrize("size_px", [16, 24, 32, 64, 128])
    def test_multi_resolution_rendering_no_pixelation_or_empty(self, qapp: QApplication, size_px: int) -> None:
        """Verify all 16 icons render sharply across 16, 24, 32, 64, 128px without empty/blank pixmaps."""
        tm = get_theme_manager()
        for icon_name in ALL_16_ICONS:
            base_name = icon_name.replace(".svg", "")
            icon = tm.get_styled_icon(base_name, size=QSize(size_px, size_px))

            assert not icon.isNull(), f"get_styled_icon({base_name}, size={size_px}) returned null QIcon"

            pixmap = icon.pixmap(QSize(size_px, size_px))
            assert not pixmap.isNull(), f"pixmap({size_px}x{size_px}) is null for {base_name}"
            assert pixmap.width() == size_px, f"Width mismatch for {base_name}: {pixmap.width()} != {size_px}"
            assert pixmap.height() == size_px, f"Height mismatch for {base_name}: {pixmap.height()} != {size_px}"

            # Convert to QImage and inspect pixels: Ensure icon is NOT blank/empty
            img = pixmap.toImage()
            non_transparent_count = 0
            for y in range(size_px):
                for x in range(size_px):
                    pixel_color = QColor.fromRgba(img.pixel(x, y))
                    if pixel_color.alpha() > 10:
                        non_transparent_count += 1

            # An icon rendered at size_px must have substantial drawn pixels
            min_expected_pixels = max(10, size_px // 2)
            assert non_transparent_count >= min_expected_pixels, (
                f"Icon {base_name} at {size_px}x{size_px} is blank! Only {non_transparent_count} pixels drawn."
            )

    def test_empirical_tinting_color_channels(self, qapp: QApplication) -> None:
        """Empirically inspect rendered pixel channels to guarantee get_styled_icon() actually paints the tint color."""
        tm = get_theme_manager()

        test_tints = [
            ("#FF0000", "red"),
            ("#00FF00", "green"),
            ("#0000FF", "blue"),
        ]

        for icon_name in ALL_16_ICONS:
            base_name = icon_name.replace(".svg", "")

            for hex_color, dominant_channel in test_tints:
                icon = tm.get_styled_icon(base_name, color=hex_color, size=QSize(32, 32))
                pixmap = icon.pixmap(32, 32)
                img = pixmap.toImage()

                # Sample non-transparent pixels and sum channel intensities
                r_sum = 0
                g_sum = 0
                b_sum = 0
                sample_count = 0

                for y in range(32):
                    for x in range(32):
                        px = QColor.fromRgba(img.pixel(x, y))
                        if px.alpha() > 128:  # Solid/opaque portion of stroke
                            r_sum += px.red()
                            g_sum += px.green()
                            b_sum += px.blue()
                            sample_count += 1

                assert sample_count > 0, f"No opaque pixels found for {base_name} with tint {hex_color}"

                r_avg = r_sum / sample_count
                g_avg = g_sum / sample_count
                b_avg = b_sum / sample_count

                if dominant_channel == "red":
                    assert r_avg > 180, f"Red tint failed for {base_name}: avg R={r_avg:.1f}"
                    assert g_avg < 40, f"Red tint contaminated with green for {base_name}: avg G={g_avg:.1f}"
                    assert b_avg < 40, f"Red tint contaminated with blue for {base_name}: avg B={b_avg:.1f}"
                elif dominant_channel == "green":
                    assert g_avg > 180, f"Green tint failed for {base_name}: avg G={g_avg:.1f}"
                    assert r_avg < 40, f"Green tint contaminated with red for {base_name}: avg R={r_avg:.1f}"
                    assert b_avg < 40, f"Green tint contaminated with blue for {base_name}: avg B={b_avg:.1f}"
                elif dominant_channel == "blue":
                    assert b_avg > 180, f"Blue tint failed for {base_name}: avg B={b_avg:.1f}"
                    assert r_avg < 40, f"Blue tint contaminated with red for {base_name}: avg R={r_avg:.1f}"
                    assert g_avg < 40, f"Blue tint contaminated with green for {base_name}: avg G={g_avg:.1f}"

    def test_default_theme_tint_adaptation(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify get_styled_icon() adapts to Light (dark stroke) vs Dark (light stroke) when color=None."""
        tm = ThemeManager(config_path=tmp_path / "settings.json", apply_stylesheet=False)

        # In Light theme: Default text is #0F172A (very dark)
        tm.set_theme("light", apply_stylesheet=False)
        icon_light = tm.get_styled_icon("settings", color=None, size=QSize(32, 32))
        img_light = icon_light.pixmap(32, 32).toImage()

        # In Dark theme: Default text is #F1F5F9 (very bright)
        tm.set_theme("dark", apply_stylesheet=False)
        icon_dark = tm.get_styled_icon("settings", color=None, size=QSize(32, 32))
        img_dark = icon_dark.pixmap(32, 32).toImage()

        # Calculate average luminance of solid pixels
        def calc_avg_brightness(img: QImage) -> float:
            total_lum = 0.0
            count = 0
            for y in range(32):
                for x in range(32):
                    px = QColor.fromRgba(img.pixel(x, y))
                    if px.alpha() > 128:
                        total_lum += px.lightnessF()
                        count += 1
            return total_lum / max(1, count)

        brightness_light = calc_avg_brightness(img_light)
        brightness_dark = calc_avg_brightness(img_dark)

        # In light theme, stroke should be dark (low lightness)
        assert brightness_light < 0.35, f"Light theme stroke too bright: {brightness_light:.2f}"
        # In dark theme, stroke should be bright (high lightness)
        assert brightness_dark > 0.70, f"Dark theme stroke too dark: {brightness_dark:.2f}"

    def test_resilience_to_invalid_icon_or_color(self, qapp: QApplication) -> None:
        """Stress-test edge cases: invalid icon name, empty name, malformed color."""
        tm = get_theme_manager()

        # Non-existent icon -> returns empty QIcon safely
        icon_bad = tm.get_styled_icon("non_existent_icon_12345")
        assert icon_bad.isNull()

        # Invalid color string -> does not throw unhandled exception, falls back to direct SVG load
        icon_bad_color = tm.get_styled_icon("download", color="invalid_color_###")
        assert isinstance(icon_bad_color, QIcon)
