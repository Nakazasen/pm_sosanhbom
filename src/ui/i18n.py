"""Internationalization (i18n) Engine for SSBOM.

Provides dynamic runtime switching between Vietnamese (vi), Japanese (ja),
and Chinese (zh) without restarting the application.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Callable, Dict, List, Optional

logger = logging.getLogger(__name__)

SUPPORTED_LANGUAGES = {
    "vi": "Tiếng Việt",
    "ja": "日本語",
    "zh": "中文",
}

DEFAULT_LANGUAGE = "vi"


class I18nManager:
    """Manages application translation strings and notify subscribers on change."""

    def __init__(self, locales_dir: Optional[Path] = None) -> None:
        if locales_dir is None:
            # Look up default locales/ directory in project root
            self.locales_dir = Path(__file__).resolve().parent.parent.parent / "locales"
        else:
            self.locales_dir = Path(locales_dir).resolve()

        self.current_language: str = DEFAULT_LANGUAGE
        self._catalogs: Dict[str, Dict[str, str]] = {}
        self._listeners: List[Callable[[str], None]] = []

        self._load_catalogs()

    def _load_catalogs(self) -> None:
        """Load JSON translations from the locales folder."""
        for lang_code in SUPPORTED_LANGUAGES:
            json_file = self.locales_dir / f"{lang_code}.json"
            if json_file.exists():
                try:
                    with open(json_file, "r", encoding="utf-8") as fp:
                        self._catalogs[lang_code] = json.load(fp)
                    logger.debug("Loaded catalog for %s (%d entries)", lang_code, len(self._catalogs[lang_code]))
                except Exception as exc:
                    logger.error("Failed to load catalog %s: %s", json_file, exc)
                    self._catalogs[lang_code] = {}
            else:
                logger.warning("Locale file not found: %s", json_file)
                self._catalogs[lang_code] = {}

    def set_language(self, lang_code: str) -> bool:
        """Switch active language and inform UI widgets."""
        if lang_code not in SUPPORTED_LANGUAGES:
            logger.warning("Unsupported language requested: %s", lang_code)
            return False

        if self.current_language == lang_code:
            return True

        self.current_language = lang_code
        logger.info("Language changed to %s (%s)", lang_code, SUPPORTED_LANGUAGES[lang_code])

        # Notify all registered UI refresh callbacks
        for listener in self._listeners:
            try:
                listener(lang_code)
            except Exception as exc:
                logger.error("Error invoking language listener: %s", exc)

        return True

    def subscribe(self, callback: Callable[[str], None]) -> None:
        """Register a callback when language changes."""
        if callback not in self._listeners:
            self._listeners.append(callback)

    def unsubscribe(self, callback: Callable[[str], None]) -> None:
        """Remove a callback."""
        if callback in self._listeners:
            self._listeners.remove(callback)

    def t(self, key: str, **kwargs) -> str:
        """Translate a key into the active language with string formatting."""
        catalog = self._catalogs.get(self.current_language, {})
        text = catalog.get(key)

        if text is None:
            # Fallback to default language
            fallback_catalog = self._catalogs.get(DEFAULT_LANGUAGE, {})
            text = fallback_catalog.get(key, key)

        if kwargs:
            try:
                return text.format(**kwargs)
            except Exception:
                return text

        return text


# Global singleton
_i18n_instance: Optional[I18nManager] = None


def get_i18n() -> I18nManager:
    """Get the global I18nManager singleton."""
    global _i18n_instance
    if _i18n_instance is None:
        _i18n_instance = I18nManager()
    return _i18n_instance


def t(key: str, **kwargs) -> str:
    """Convenience translation function."""
    return get_i18n().t(key, **kwargs)

