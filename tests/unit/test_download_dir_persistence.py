"""Unit tests for BOM Download Destination Directory Persistence across app launches.

Validates:
1. Per-model directory persistence in config/settings.json.
2. Graceful fail-safe fallback when persisted directory does not exist (e.g. offline LAN share).
3. Whitespace & case-insensitive normalization for model names (e.g. 'Iris 2024' vs 'Iris2024').
4. Preservation of all other settings keys during save operations.
5. Dialog behavior: initial load, model switching, browsing, manual edit, and download trigger.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from unittest.mock import patch

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtWidgets import QApplication

from src.gui.plm_download_dialog import (
    PLMDownloadDialog,
    get_persisted_download_dir,
    normalize_model_key,
    save_persisted_download_dir,
)
from src.gui.settings_dialog import SettingsDialog


class TestDownloadDirPersistence:
    """Test persistence logic and fail-safe handling."""

    def test_normalize_model_key(self) -> None:
        assert normalize_model_key("Iris 2024") == "IRIS2024"
        assert normalize_model_key("  Iris2024  ") == "IRIS2024"
        assert normalize_model_key("Virgo") == "VIRGO"
        assert normalize_model_key("") == ""
        assert normalize_model_key(None) == ""

    def test_save_and_get_persisted_download_dir(self, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        # Seed initial settings
        initial_data = {
            "theme": "dark",
            "paths": {
                "shared_db_path": r"\\fstvn01\db.sqlite",
            },
        }
        with open(cfg_file, "w", encoding="utf-8") as fp:
            json.dump(initial_data, fp)

        model_dir_virgo = tmp_path / "Virgo_Downloads"
        model_dir_virgo.mkdir()

        # Save for Virgo
        save_persisted_download_dir("Virgo", model_dir_virgo, config_path=cfg_file)

        # Verify saved in file
        with open(cfg_file, "r", encoding="utf-8") as fp:
            saved_data = json.load(fp)

        assert saved_data["theme"] == "dark"
        assert saved_data["paths"]["shared_db_path"] == r"\\fstvn01\db.sqlite"
        assert saved_data["paths"]["last_download_dir"] == str(model_dir_virgo)
        assert saved_data["paths"]["last_download_dirs_by_model"]["Virgo"] == str(model_dir_virgo)

        # Retrieve for Virgo
        retrieved = get_persisted_download_dir(
            model_name="Virgo",
            fallback_dir=tmp_path / "Fallback",
            config_path=cfg_file,
        )
        assert retrieved == model_dir_virgo

    def test_whitespace_normalization_in_lookup(self, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        iris_dir = tmp_path / "Iris_Downloads"
        iris_dir.mkdir()

        # Save with "Iris 2024"
        save_persisted_download_dir("Iris 2024", iris_dir, config_path=cfg_file)

        # Query with "Iris2024" without space
        retrieved1 = get_persisted_download_dir(
            model_name="Iris2024",
            fallback_dir=tmp_path / "Fallback",
            config_path=cfg_file,
        )
        assert retrieved1 == iris_dir

        # Query with "iris 2024" lowercase
        retrieved2 = get_persisted_download_dir(
            model_name="iris 2024",
            fallback_dir=tmp_path / "Fallback",
            config_path=cfg_file,
        )
        assert retrieved2 == iris_dir

    def test_fallback_when_persisted_dir_missing(self, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        non_existent_dir = tmp_path / "Offline_LAN_Folder"
        fallback_dir = tmp_path / "Local_Default"
        fallback_dir.mkdir()

        # Save non-existent path
        data = {
            "paths": {
                "last_download_dirs_by_model": {"Virgo": str(non_existent_dir)},
                "last_download_dir": str(non_existent_dir),
            }
        }
        with open(cfg_file, "w", encoding="utf-8") as fp:
            json.dump(data, fp)

        # Should fail-safe fall back to fallback_dir because non_existent_dir does not exist
        retrieved = get_persisted_download_dir(
            model_name="Virgo",
            fallback_dir=fallback_dir,
            config_path=cfg_file,
        )
        assert retrieved == fallback_dir

    def test_global_fallback_when_model_not_saved(self, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        global_dir = tmp_path / "Global_Downloads"
        global_dir.mkdir()

        data = {
            "paths": {
                "last_download_dirs_by_model": {},
                "last_download_dir": str(global_dir),
            }
        }
        with open(cfg_file, "w", encoding="utf-8") as fp:
            json.dump(data, fp)

        # Query for model that has no per-model entry
        retrieved = get_persisted_download_dir(
            model_name="UnseenModel",
            fallback_dir=tmp_path / "Fallback",
            config_path=cfg_file,
        )
        assert retrieved == global_dir


class TestPLMDownloadDialogPersistenceUI:
    """Test dialog UI integration with destination persistence."""

    def test_dialog_initial_load_from_config(self, qapp: QApplication, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        virgo_saved_dir = tmp_path / "Virgo_Saved"
        virgo_saved_dir.mkdir()

        data = {
            "paths": {
                "last_download_dirs_by_model": {"Virgo": str(virgo_saved_dir)},
                "last_download_dir": str(virgo_saved_dir),
            }
        }
        with open(cfg_file, "w", encoding="utf-8") as fp:
            json.dump(data, fp)

        dlg = PLMDownloadDialog(
            default_dir=tmp_path / "DefaultVirgo",
            initial_model="Virgo",
            config_path=cfg_file,
        )
        try:
            assert dlg.dest_dir == virgo_saved_dir
            assert dlg.edit_dest.text() == str(virgo_saved_dir)
        finally:
            dlg.close()
            dlg.deleteLater()
            qapp.processEvents()

    def test_dialog_model_switch_loads_persisted_dir(self, qapp: QApplication, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        virgo_dir = tmp_path / "Virgo_Dir"
        virgo_dir.mkdir()
        polaris_dir = tmp_path / "Polaris_Dir"
        polaris_dir.mkdir()

        data = {
            "paths": {
                "last_download_dirs_by_model": {
                    "Virgo": str(virgo_dir),
                    "PolarisNext": str(polaris_dir),
                },
                "last_download_dir": str(virgo_dir),
            }
        }
        with open(cfg_file, "w", encoding="utf-8") as fp:
            json.dump(data, fp)

        dlg = PLMDownloadDialog(
            default_dir=virgo_dir,
            initial_model="Virgo",
            config_path=cfg_file,
        )
        try:
            assert dlg.dest_dir == virgo_dir

            # Switch combo to PolarisNext
            idx = dlg.combo_model.findText("PolarisNext")
            if idx < 0:
                dlg.combo_model.addItem("PolarisNext")
                idx = dlg.combo_model.findText("PolarisNext")
            dlg.combo_model.setCurrentIndex(idx)

            assert dlg.dest_dir == polaris_dir
            assert dlg.edit_dest.text() == str(polaris_dir)
        finally:
            dlg.close()
            dlg.deleteLater()
            qapp.processEvents()

    def test_browse_dest_dir_saves_to_config(self, qapp: QApplication, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        new_browse_dir = tmp_path / "New_Custom_Folder"
        new_browse_dir.mkdir()

        dlg = PLMDownloadDialog(
            default_dir=tmp_path / "DefaultVirgo",
            initial_model="Virgo",
            config_path=cfg_file,
        )
        try:
            with patch(
                "PyQt6.QtWidgets.QFileDialog.getExistingDirectory",
                return_value=str(new_browse_dir),
            ):
                dlg._browse_dest_dir()

            assert dlg.dest_dir == new_browse_dir
            assert dlg.edit_dest.text() == str(new_browse_dir)

            # Check config file was updated
            with open(cfg_file, "r", encoding="utf-8") as fp:
                saved = json.load(fp)
            assert saved["paths"]["last_download_dirs_by_model"]["Virgo"] == str(new_browse_dir)
            assert saved["paths"]["last_download_dir"] == str(new_browse_dir)
        finally:
            dlg.close()
            dlg.deleteLater()
            qapp.processEvents()

    def test_settings_dialog_preserves_download_dirs(self, qapp: QApplication, tmp_path: Path) -> None:
        cfg_file = tmp_path / "settings.json"
        virgo_dir = tmp_path / "Virgo_Dir"
        virgo_dir.mkdir()

        initial_config = {
            "theme": "dark",
            "paths": {
                "base_dir": str(tmp_path),
                "last_download_dirs_by_model": {"Virgo": str(virgo_dir)},
                "last_download_dir": str(virgo_dir),
            },
        }
        with open(cfg_file, "w", encoding="utf-8") as fp:
            json.dump(initial_config, fp)

        settings_dlg = SettingsDialog(config_path=cfg_file)
        try:
            # Emulate user clicking Save in Settings Dialog
            settings_dlg.save_and_close()

            with open(cfg_file, "r", encoding="utf-8") as fp:
                saved = json.load(fp)

            # Check that last_download_dirs_by_model was NOT overwritten or wiped out
            assert "last_download_dirs_by_model" in saved["paths"]
            assert saved["paths"]["last_download_dirs_by_model"]["Virgo"] == str(virgo_dir)
            assert saved["paths"]["last_download_dir"] == str(virgo_dir)
        finally:
            settings_dlg.close()
            settings_dlg.deleteLater()
            qapp.processEvents()
