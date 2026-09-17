"""Unit tests for i18n dynamic trilingual translation engine."""

import pytest
from src.ui.i18n import I18nManager, SUPPORTED_LANGUAGES, get_i18n, t


class TestI18nEngine:
    """Test suite for trilingual localization."""

    def test_all_supported_languages_configured(self):
        """Verify Vietnamese, Japanese, and Chinese are supported."""
        assert "vi" in SUPPORTED_LANGUAGES
        assert "ja" in SUPPORTED_LANGUAGES
        assert "zh" in SUPPORTED_LANGUAGES

    def test_loading_catalogs(self):
        """Verify JSON catalogs are successfully loaded."""
        mgr = I18nManager()
        assert len(mgr._catalogs["vi"]) > 0
        assert len(mgr._catalogs["ja"]) > 0
        assert len(mgr._catalogs["zh"]) > 0

    def test_translation_switch(self):
        """Verify dynamic switching between languages."""
        mgr = I18nManager()

        mgr.set_language("vi")
        assert "So Sánh BOM" in mgr.t("app_title")

        mgr.set_language("ja")
        assert "自動BOM照合" in mgr.t("app_title")

        mgr.set_language("zh")
        assert "自动BOM比对" in mgr.t("app_title")

    def test_translation_formatting(self):
        """Verify string format parameters."""
        mgr = I18nManager()
        mgr.set_language("vi")
        res = mgr.t("msg_confirm_run", model="TEST1234")
        assert "TEST1234" in res

    def test_subscriber_notification(self):
        """Verify callback notification upon language change."""
        mgr = I18nManager()
        notifications = []

        def callback(code):
            notifications.append(code)

        mgr.subscribe(callback)
        mgr.set_language("ja")
        assert "ja" in notifications

