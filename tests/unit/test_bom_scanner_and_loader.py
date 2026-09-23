"""Unit tests for BOM Status Scanner and Quick Loader.

Validates:
1. Regex extraction of Kyocera 10-character machine codes from filenames/folders.
2. Directory scanning and detection of PLM & SAP R3 availability (needs_r3, needs_plm, is_complete).
3. BOMScanDialog modal UI: source switching, table rendering, 1-click status filtering, and selection.
4. Smart Suggestion Banner in PLMDownloadDialog: automatic detection, 1-click fill, dismissal, and settings toggle.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from unittest.mock import patch

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication

from src.gui.plm_download_dialog import (
    BOMPartStatus,
    BOMScanDialog,
    PLMDownloadDialog,
    extract_part_from_name,
    is_bom_scan_suggestion_enabled,
    scan_bom_directory,
    set_bom_scan_suggestion_enabled,
)


class TestBOMScannerLogic:
    """Test part extraction and filesystem scanning logic."""

    def test_extract_part_from_name(self) -> None:
        assert extract_part_from_name("PLM_110C2MTTW0.xlsx") == "110C2MTTW0"
        assert extract_part_from_name("R3_110C2L9JP0.xlsx") == "110C2L9JP0"
        assert extract_part_from_name("PLM_110c103nl0_v2.pure.xlsx") == "110C103NL0"
        assert extract_part_from_name("110C0Z3LV1") == "110C0Z3LV1"
        assert extract_part_from_name("backup_before_virgo_filter") is None
        assert extract_part_from_name("") is None
        assert extract_part_from_name(None) is None

    def test_scan_bom_directory_plm_and_r3_status(self, tmp_path: Path) -> None:
        # Create a test directory with 3 scenarios:
        # 1. 110C103NL0: PLM only (needs R3)
        # 2. 110C103NL1: R3 only (needs PLM)
        # 3. 110C2MTTW0: Both PLM and R3 (is complete)
        # 4. 110C0Z3LV1: In a subfolder with PLM only

        (tmp_path / "PLM_110C103NL0.xlsx").write_bytes(b"plm dummy data 1")
        (tmp_path / "R3_110C103NL1.xls").write_bytes(b"r3 dummy data 2")
        (tmp_path / "PLM_110C2MTTW0.xlsx").write_bytes(b"plm dummy data 3")
        (tmp_path / "R3_110C2MTTW0.xlsx").write_bytes(b"r3 dummy data 3")

        sub_folder = tmp_path / "110C0Z3LV1"
        sub_folder.mkdir()
        (sub_folder / "PLM_110C0Z3LV1.xlsx").write_bytes(b"subfolder plm data")

        results = scan_bom_directory(tmp_path)
        res_map = {p.part_number: p for p in results}

        assert len(res_map) == 4

        # 110C103NL0: has PLM, missing R3
        s1 = res_map["110C103NL0"]
        assert s1.has_plm is True
        assert s1.has_r3 is False
        assert s1.needs_r3 is True
        assert s1.needs_plm is False
        assert s1.is_complete is False
        assert "Cần tải bù SAP R3" in s1.status_summary

        # 110C103NL1: has R3, missing PLM
        s2 = res_map["110C103NL1"]
        assert s2.has_plm is False
        assert s2.has_r3 is True
        assert s2.needs_r3 is False
        assert s2.needs_plm is True
        assert s2.is_complete is False
        assert "Cần tải bù PLM" in s2.status_summary

        # 110C2MTTW0: has both
        s3 = res_map["110C2MTTW0"]
        assert s3.has_plm is True
        assert s3.has_r3 is True
        assert s3.needs_r3 is False
        assert s3.needs_plm is False
        assert s3.is_complete is True
        assert "Đã đầy đủ" in s3.status_summary

        # 110C0Z3LV1: from subfolder
        s4 = res_map["110C0Z3LV1"]
        assert s4.has_plm is True
        assert s4.has_r3 is False
        assert s4.needs_r3 is True

    def test_scan_bom_directory_with_extra_project_parts(self, tmp_path: Path) -> None:
        (tmp_path / "PLM_110C103NL0.xlsx").write_bytes(b"data")

        # Project has a part that is not yet downloaded anywhere
        extra = ["110C999US0"]
        results = scan_bom_directory(tmp_path, extra_parts=extra)
        res_map = {p.part_number: p for p in results}

        assert "110C999US0" in res_map
        s_unseen = res_map["110C999US0"]
        assert s_unseen.has_plm is False
        assert s_unseen.has_r3 is False
        assert s_unseen.is_complete is False


class TestBOMScanDialogUI:
    """Test BOMScanDialog table and filtering behaviors."""

    def test_scan_dialog_filtering_and_selection(self, qapp: QApplication, tmp_path: Path) -> None:
        (tmp_path / "PLM_110C103NL0.xlsx").write_bytes(b"123")
        (tmp_path / "R3_110C103NL1.xls").write_bytes(b"456")
        (tmp_path / "PLM_110C2MTTW0.xlsx").write_bytes(b"789")
        (tmp_path / "R3_110C2MTTW0.xlsx").write_bytes(b"789")

        dlg = BOMScanDialog(dest_dir=tmp_path, project_parts=["110C103NL0", "110C103NL1"])
        try:
            assert dlg.table.rowCount() == 3

            # 1. Filter: Only missing R3 -> should select 110C103NL0
            dlg._select_missing_r3()
            checked_r3 = [
                dlg.table.item(r, 1).text()
                for r in range(dlg.table.rowCount())
                if dlg.table.item(r, 0).checkState() == Qt.CheckState.Checked
            ]
            assert checked_r3 == ["110C103NL0"]

            # 2. Filter: Only missing PLM -> should select 110C103NL1
            dlg._select_missing_plm()
            checked_plm = [
                dlg.table.item(r, 1).text()
                for r in range(dlg.table.rowCount())
                if dlg.table.item(r, 0).checkState() == Qt.CheckState.Checked
            ]
            assert checked_plm == ["110C103NL1"]

            # 3. Filter: Complete only -> should select 110C2MTTW0
            dlg._select_complete()
            checked_comp = [
                dlg.table.item(r, 1).text()
                for r in range(dlg.table.rowCount())
                if dlg.table.item(r, 0).checkState() == Qt.CheckState.Checked
            ]
            assert checked_comp == ["110C2MTTW0"]

            # 4. Select All -> selects all 3
            dlg._select_all()
            checked_all = [
                dlg.table.item(r, 1).text()
                for r in range(dlg.table.rowCount())
                if dlg.table.item(r, 0).checkState() == Qt.CheckState.Checked
            ]
            assert len(checked_all) == 3

            # 5. Select None
            dlg._select_none()
            checked_none = [
                dlg.table.item(r, 1).text()
                for r in range(dlg.table.rowCount())
                if dlg.table.item(r, 0).checkState() == Qt.CheckState.Checked
            ]
            assert len(checked_none) == 0

            # 6. Confirm selection
            dlg._select_missing_r3()
            dlg._confirm_selection()
            assert dlg.get_selected_parts() == ["110C103NL0"]

        finally:
            dlg.close()
            dlg.deleteLater()
            qapp.processEvents()

    def test_scan_dialog_suggestion_preference_toggle(self, qapp: QApplication, tmp_path: Path) -> None:
        cfg = tmp_path / "settings.json"
        with open(cfg, "w", encoding="utf-8") as fp:
            json.dump({"paths": {"show_bom_scan_suggestion": True}}, fp)

        assert is_bom_scan_suggestion_enabled(cfg) is True

        set_bom_scan_suggestion_enabled(False, cfg)
        assert is_bom_scan_suggestion_enabled(cfg) is False

        set_bom_scan_suggestion_enabled(True, cfg)
        assert is_bom_scan_suggestion_enabled(cfg) is True


class TestSmartSuggestionBanner:
    """Test smart banner presentation and 1-click loading inside PLMDownloadDialog."""

    def test_smart_banner_detection_and_1click_load(self, qapp: QApplication, tmp_path: Path) -> None:
        # Create folder with 2 files that have PLM but no R3
        (tmp_path / "PLM_110C2MTTW0.xlsx").write_bytes(b"data1")
        (tmp_path / "PLM_110C2L9JP0.xlsx").write_bytes(b"data2")

        cfg = tmp_path / "settings.json"
        with open(cfg, "w", encoding="utf-8") as fp:
            json.dump({
                "paths": {
                    "last_download_dirs_by_model": {"Virgo": str(tmp_path)},
                    "last_download_dir": str(tmp_path),
                    "show_bom_scan_suggestion": True,
                }
            }, fp)

        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo", config_path=cfg)
        dlg.show()
        qapp.processEvents()

        try:
            # Banner should be visible because folder has 2 items needing R3 and txt_parts is empty
            assert dlg.banner_frame.isVisible() is True
            assert "chưa có BOM R3" in dlg.lbl_banner_msg.text()
            assert dlg.btn_banner_missing_r3.isVisible() is True

            # Click 1-Click Load Missing R3
            dlg._on_banner_load_missing_r3()

            # txt_parts should now contain both parts
            text = dlg.txt_parts.toPlainText()
            assert "110C2MTTW0" in text
            assert "110C2L9JP0" in text

            # Banner should automatically hide once text is filled
            assert dlg.banner_frame.isVisible() is False

            # Clear parts -> banner should reappear
            dlg._clear_parts()
            assert dlg.banner_frame.isVisible() is True

            # Click Dismiss -> banner should hide and remain hidden
            dlg._dismiss_smart_suggestion()
            assert dlg.banner_frame.isVisible() is False
            dlg._clear_parts()
            assert dlg.banner_frame.isVisible() is False

        finally:
            dlg.close()
            dlg.deleteLater()
            qapp.processEvents()

    def test_smart_banner_respects_disabled_setting(self, qapp: QApplication, tmp_path: Path) -> None:
        cfg = tmp_path / "settings.json"
        with open(cfg, "w", encoding="utf-8") as fp:
            json.dump({
                "paths": {
                    "last_download_dirs_by_model": {"Virgo": str(tmp_path)},
                    "last_download_dir": str(tmp_path),
                    "show_bom_scan_suggestion": False,
                }
            }, fp)

        (tmp_path / "PLM_110C2MTTW0.xlsx").write_bytes(b"data")

        dlg = PLMDownloadDialog(default_dir=tmp_path, initial_model="Virgo", config_path=cfg)
        dlg.show()
        qapp.processEvents()
        try:
            assert dlg.banner_frame.isVisible() is False
        finally:
            dlg.close()
            dlg.deleteLater()
            qapp.processEvents()
