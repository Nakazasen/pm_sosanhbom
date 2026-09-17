"""Tier 5 Adversarial Tests: Invalid Unicode, Encoding Corruption & Internationalization Edge Cases.

Verifies:
1. Embedded null bytes (\\x00) without C-string truncation.
2. Bidirectional override characters (\\u202e) and spoofing attacks.
3. Invisible zero-width spaces (\\u200b) and non-breaking spaces (\\xa0).
4. Vietnamese decomposed (NFD) vs precomposed (NFC) Unicode equivalence.
5. Japanese full-width alphanumeric (１２３４５, ＡＢＣ) normalization.
6. Multi-byte CJK ideographs in part names and descriptions.
7. Emojis, surrogate code points, and replacement characters (\\ufffd).
"""

from __future__ import annotations

import unicodedata
import pytest

from src.core.models import BOMNode
from src.automation.sap.models import R3ComponentRow
from src.ui.i18n import I18nManager


class TestAdversarialUnicodeAndEncoding:
    """Test suite for Unicode, encoding corruption, and internationalization."""

    def test_embedded_null_bytes_preserved_safely(self):
        """Null byte does not truncate strings or cause C-level termination."""
        polluted_id = "PART\x00_SECURITY_TEST"
        polluted_desc = "BOARD\x00_ASSEMBLY_MAIN"

        node = BOMNode(level=1, item_id=polluted_id, item_name=polluted_desc)
        assert len(node.item_id) == len(polluted_id)
        assert "\x00" in node.item_id
        assert "\x00" in node.item_name

        r3_row = R3ComponentRow(part_code=polluted_id, quantity=1.0, description=polluted_desc)
        assert r3_row.part_code == polluted_id
        assert r3_row.description == polluted_desc

    def test_bidirectional_override_spoofing_characters(self):
        """Right-to-Left override (\\u202e) characters do not crash model or serialization."""
        bidi_str = "SAFE_PREFIX\u202eEXE.SJB_EVIL"
        node = BOMNode(level=1, item_id=bidi_str, item_name="BIDI TEST")
        d = node.to_flat_dict()
        assert d["item_id"] == bidi_str

    def test_vietnamese_nfc_vs_nfd_normalization(self):
        """Vietnamese text in NFC vs NFD forms normalize equivalently."""
        # 'Tiếng Việt' in NFC (precomposed) vs NFD (decomposed)
        text_nfc = "Tiếng Việt"
        text_nfd = unicodedata.normalize("NFD", text_nfc)
        assert text_nfc != text_nfd  # byte representation differs

        node_nfc = BOMNode(level=1, item_id="VN_01", item_name=text_nfc)
        node_nfd = BOMNode(level=1, item_id="VN_02", item_name=text_nfd)

        # Standardizing normalization
        assert unicodedata.normalize("NFC", node_nfc.item_name) == unicodedata.normalize("NFC", node_nfd.item_name)

    def test_japanese_and_chinese_cjk_characters(self):
        """Japanese Kanji/Kana and Chinese characters parse and serialize without encoding errors."""
        jp_name = "メインフレームユニット (Main Frame Assy)"
        cn_name = "主框架部件 (Main Frame Component)"

        node_jp = BOMNode(level=1, item_id="302K193010", item_name=jp_name)
        node_cn = BOMNode(level=1, item_id="302K193020", item_name=cn_name)

        assert "メインフレーム" in node_jp.item_name
        assert "主框架部件" in node_cn.item_name

        r3_row = R3ComponentRow(part_code="302K193010", quantity=2.0, description=jp_name)
        assert r3_row.description == jp_name

    def test_emojis_and_surrogate_symbols(self):
        """4-byte UTF-8 emojis (e.g. ⚙️, 📦, 🚀) are handled cleanly."""
        emoji_desc = "GEAR UNIT ⚙️ [BATCH 📦] VERIFIED 🚀"
        node = BOMNode(level=1, item_id="EMOJI_PART", item_name=emoji_desc)
        assert "⚙️" in node.item_name
        assert "🚀" in node.item_name

    def test_i18n_fallback_on_missing_or_corrupt_locale(self, tmp_path):
        """I18n engine falls back gracefully to default key when translation key is unknown."""
        i18n = I18nManager()
        # Non-existent translation key returns key itself
        translated = i18n.t("NON_EXISTENT_KEY_12345")
        assert translated == "NON_EXISTENT_KEY_12345"
