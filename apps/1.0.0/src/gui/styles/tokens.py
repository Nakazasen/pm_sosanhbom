"""Design Tokens and WCAG Math Utilities for SSBOM Application.

Defines:
- Slate Industrial (Light Theme) and Industrial Dark Mode color constants
- Data-Dense table dimensions (row height 32, padding 4x8, border 1px)
- Typography scale and font families
- W3C WCAG 2.1 relative luminance and contrast ratio calculation engine
"""

from __future__ import annotations

import re

# ============================================================================
# LIGHT THEME TOKENS (Slate Industrial)
# ============================================================================
COLOR_LIGHT_BG_CANVAS = "#F8FAFC"
COLOR_LIGHT_SURFACE = "#FFFFFF"
COLOR_LIGHT_TABLE_BG = "#FFFFFF"
COLOR_LIGHT_ZEBRA_EVEN = "#FFFFFF"
COLOR_LIGHT_ZEBRA_ODD = "#F8FAFC"
COLOR_LIGHT_ROW_HOVER = "#F1F5F9"
COLOR_LIGHT_ROW_SELECTED = "#E0E7FF"
COLOR_LIGHT_ROW_SELECTED_TEXT = "#0F172A"
COLOR_LIGHT_CELL_BORDER_BOTTOM = "#F1F5F9"
COLOR_LIGHT_BORDER_GRID = "#CBD5E1"
COLOR_LIGHT_HEADER_BG = "#F1F5F9"
COLOR_LIGHT_HEADER_TEXT = "#475569"
COLOR_LIGHT_TEXT_PRIMARY = "#0F172A"
COLOR_LIGHT_TEXT_SECONDARY = "#475569"
COLOR_LIGHT_BRAND_PRIMARY = "#2563EB"
COLOR_LIGHT_BRAND_HOVER = "#1D4ED8"
COLOR_LIGHT_BRAND_TEXT = "#FFFFFF"

# Semantic Status (Light)
COLOR_LIGHT_STATUS_MATCH_BG = "#DCFCE7"
COLOR_LIGHT_STATUS_MATCH_TEXT = "#166534"
COLOR_LIGHT_STATUS_MATCH_BORDER = "#86EFAC"

COLOR_LIGHT_STATUS_DIFF_BG = "#FEE2E2"
COLOR_LIGHT_STATUS_DIFF_TEXT = "#991B1B"
COLOR_LIGHT_STATUS_DIFF_BORDER = "#FCA5A5"

COLOR_LIGHT_STATUS_WARN_BG = "#FEF3C7"
COLOR_LIGHT_STATUS_WARN_TEXT = "#92400E"
COLOR_LIGHT_STATUS_WARN_BORDER = "#FCD34D"

COLOR_LIGHT_STATUS_INFO_BG = "#E0F2FE"
COLOR_LIGHT_STATUS_INFO_TEXT = "#075985"
COLOR_LIGHT_STATUS_INFO_BORDER = "#7DD3FC"

# ============================================================================
# DARK THEME TOKENS (Industrial Dark Mode)
# ============================================================================
COLOR_DARK_BG_CANVAS = "#0B0F17"
COLOR_DARK_SURFACE = "#151D2A"
COLOR_DARK_TABLE_BG = "#151D2A"
COLOR_DARK_ZEBRA_EVEN = "#151D2A"
COLOR_DARK_ZEBRA_ODD = "#111722"
COLOR_DARK_ROW_HOVER = "#1A2436"
COLOR_DARK_ROW_SELECTED = "#1E3A5F"
COLOR_DARK_ROW_SELECTED_TEXT = "#F1F5F9"
COLOR_DARK_CELL_BORDER_BOTTOM = "#1F293D"
COLOR_DARK_BORDER_GRID = "#2A374A"
COLOR_DARK_HEADER_BG = "#1A2436"
COLOR_DARK_HEADER_TEXT = "#94A3B8"
COLOR_DARK_TEXT_PRIMARY = "#F1F5F9"
COLOR_DARK_TEXT_SECONDARY = "#94A3B8"
COLOR_DARK_BRAND_PRIMARY = "#3B82F6"
COLOR_DARK_BRAND_HOVER = "#60A5FA"
COLOR_DARK_BRAND_TEXT = "#0B0F17"

# Semantic Status (Dark)
COLOR_DARK_STATUS_MATCH_BG = "#064E3B"
COLOR_DARK_STATUS_MATCH_TEXT = "#A7F3D0"
COLOR_DARK_STATUS_MATCH_BORDER = "#047857"

COLOR_DARK_STATUS_DIFF_BG = "#7F1D1D"
COLOR_DARK_STATUS_DIFF_TEXT = "#FECACA"
COLOR_DARK_STATUS_DIFF_BORDER = "#B91C1C"

COLOR_DARK_STATUS_WARN_BG = "#78350F"
COLOR_DARK_STATUS_WARN_TEXT = "#FDE68A"
COLOR_DARK_STATUS_WARN_BORDER = "#D97706"

COLOR_DARK_STATUS_INFO_BG = "#0C4A6E"
COLOR_DARK_STATUS_INFO_TEXT = "#BAE6FD"
COLOR_DARK_STATUS_INFO_BORDER = "#0284C7"

# ============================================================================
# DIMENSIONS & SIZING CONSTANTS (Data-Dense Standard)
# ============================================================================
TABLE_ROW_HEIGHT: int = 32
CELL_PADDING: tuple[int, int] = (4, 8)
HEADER_PADDING: tuple[int, int] = (6, 8)
PROGRESS_BAR_HEIGHT: int = 14
BORDER_WIDTH: int = 1
BORDER_RADIUS: int = 4
TRANSITION_DURATION_MS: int = 150
MIN_FONT_SIZE: int = 11

# ============================================================================
# TYPOGRAPHY SCALE
# ============================================================================
FONT_FAMILY_PRIMARY: str = (
    '"Segoe UI", "Aptos", -apple-system, BlinkMacSystemFont, "Meiryo", "Microsoft YaHei", sans-serif'
)
FONT_FAMILY_MONO: str = '"Consolas", "Courier New", monospace'

FONT_SIZE_APP_TITLE: int = 18
FONT_SIZE_HEADING: int = 15
FONT_SIZE_TABLE_HEADER: int = 12
FONT_SIZE_TABLE_CELL: int = 12
FONT_SIZE_CODE: int = 12
FONT_SIZE_CAPTION: int = 11
FONT_SIZE_LOG: int = 11

FONT_WEIGHT_REGULAR: int = 400
FONT_WEIGHT_MEDIUM: int = 500
FONT_WEIGHT_SEMIBOLD: int = 600

# ============================================================================
# WCAG 2.1 RELATIVE LUMINANCE & CONTRAST RATIO ENGINE
# ============================================================================

_HEX_REGEX = re.compile(r"^#?([A-Fa-f0-9]{3}|[A-Fa-f0-9]{6})$")


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert hex string (e.g. '#0F172A' or '#FFF') to RGB tuple (0..255)."""
    match = _HEX_REGEX.match(hex_color.strip())
    if not match:
        raise ValueError(f"Invalid hex color format: '{hex_color}'")

    raw = match.group(1)
    if len(raw) == 3:
        raw = "".join(ch * 2 for ch in raw)

    r = int(raw[0:2], 16)
    g = int(raw[2:4], 16)
    b = int(raw[4:6], 16)
    return r, g, b


def calculate_relative_luminance(hex_color: str) -> float:
    """Calculate relative luminance L of an sRGB color per W3C WCAG 2.1 specs.

    L = 0.2126 * R_lin + 0.7152 * G_lin + 0.0722 * B_lin
    where C_lin = C/12.92 if C <= 0.04045 else ((C + 0.055)/1.055) ** 2.4
    """
    r, g, b = hex_to_rgb(hex_color)

    def _to_linear(channel: int) -> float:
        s = channel / 255.0
        if s <= 0.04045:
            return s / 12.92
        return ((s + 0.055) / 1.055) ** 2.4

    r_lin = _to_linear(r)
    g_lin = _to_linear(g)
    b_lin = _to_linear(b)

    return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin


def calculate_contrast_ratio(fg: str, bg: str) -> float:
    """Calculate WCAG 2.1 contrast ratio between foreground and background.

    Ratio = (L1 + 0.05) / (L2 + 0.05) where L1 is lighter and L2 is darker.
    Returns float in range [1.0, 21.0].
    """
    l1 = calculate_relative_luminance(fg)
    l2 = calculate_relative_luminance(bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def is_wcag_aa(fg: str, bg: str, is_large_text: bool = False) -> bool:
    """Check if color combination passes WCAG 2.1 AA (>= 4.5:1, or >= 3.0:1 for large text)."""
    threshold = 3.0 if is_large_text else 4.5
    return calculate_contrast_ratio(fg, bg) >= threshold


def is_wcag_aaa(fg: str, bg: str, is_large_text: bool = False) -> bool:
    """Check if color combination passes WCAG 2.1 AAA (>= 7.0:1, or >= 4.5:1 for large text)."""
    threshold = 4.5 if is_large_text else 7.0
    return calculate_contrast_ratio(fg, bg) >= threshold
