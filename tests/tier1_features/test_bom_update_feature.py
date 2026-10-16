"""Unit tests for BOM Update Feature (modernizing VBA Buttons 7 & 8).

Verifies:
1. InheritanceEngine ham_match_index_mix logic preserves Col O, P, Q.
2. BOMUpdateDialog instantiation and file discovery.
3. BOMUpdateWorker execution logic.
"""

from __future__ import annotations

import os
from pathlib import Path
import tempfile

import openpyxl
import pandas as pd
import pytest
from PyQt6.QtWidgets import QApplication

from src.core.inheritance_engine import InheritanceEngine
from src.gui.bom_update_dialog import BOMUpdateDialog, BOMUpdateWorker


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_inheritance_engine_match_index_mix():
    """Verify legacy ham_match_index_mix preservation of explanations."""
    df_old = pd.DataFrame({
        "item_id": ["PART-001", "PART-002", "PART-003"],
        "giai_thich": ["Linh kiện tiêu chuẩn", "Đã thay đổi theo ECN", "Tạm thời không dùng"],
        "phu_trach": ["Son_mecha1", "Duy_mecha1", "Hieu_mecha1"],
        "quan_ly_check": ["OK", "OK", "PENDING"],
    })

    df_new = pd.DataFrame({
        "item_id": ["PART-002", "PART-004", "PART-001"],
        "quantity": [2, 1, 5],
    })

    res = InheritanceEngine.match_index_mix(df_new=df_new, df_old=df_old, key_column="item_id")

    assert len(res) == 3
    # PART-002 should inherit its explanation
    row_002 = res[res["item_id"] == "PART-002"].iloc[0]
    assert row_002["giai_thich"] == "Đã thay đổi theo ECN"
    assert row_002["phu_trach"] == "Duy_mecha1"
    assert row_002["quan_ly_check"] == "OK"

    # PART-001 should inherit its explanation
    row_001 = res[res["item_id"] == "PART-001"].iloc[0]
    assert row_001["giai_thich"] == "Linh kiện tiêu chuẩn"
    assert row_001["phu_trach"] == "Son_mecha1"
    assert row_001["quan_ly_check"] == "OK"

    # PART-004 is new, should be empty strings without NaN
    row_004 = res[res["item_id"] == "PART-004"].iloc[0]
    assert row_004["giai_thich"] == ""
    assert row_004["phu_trach"] == ""
    assert row_004["quan_ly_check"] == ""


def test_bom_update_dialog_init(qapp):
    """Test dialog initialization and folder scanning."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        bom_file = tmp_path / "BOM_Spica.xlsx"
        plm_file = tmp_path / "PLM_T10C0L2US0.xlsx"
        r3_file = tmp_path / "R3_3V2YJ58070.xlsx"

        # Create dummy excel files
        wb = openpyxl.Workbook()
        wb.save(bom_file)
        wb.save(plm_file)
        wb.save(r3_file)

        dlg = BOMUpdateDialog(initial_dir=tmp_path)
        assert dlg.current_bom_file == bom_file
        assert dlg.current_plm_file == plm_file
        assert dlg.current_r3_file == r3_file
        assert dlg.btn_run_plm.isEnabled()
        assert dlg.btn_run_r3.isEnabled()
        assert dlg.btn_run_both.isEnabled()
        dlg.close()


def test_bom_update_dialog_theme_and_bottom_buttons(qapp):
    """Verify dark mode contrast, status label contrast, and bottom buttons existence."""
    dlg = BOMUpdateDialog()

    # 1. Verify bottom bar buttons exist with required labels
    assert dlg.btn_open_target_folder is not None
    assert "Mở thư mục chứa file BOM" in dlg.btn_open_target_folder.text()
    assert dlg.btn_open_archive_folder is not None
    assert "Mở thư mục lưu trữ" in dlg.btn_open_archive_folder.text()
    assert dlg.btn_close is not None
    assert "Đóng" in dlg.btn_close.text()

    # 2. Simulate Dark Mode
    dlg.theme_mgr._effective_theme = "dark"
    dlg._apply_theme_styles()

    # Check status box has dark background
    assert "#1E293B" in dlg.status_box.styleSheet()
    assert "#F8FAFC" in dlg.lbl_title_bom.styleSheet()

    # When no BOM file is selected, status label should have high contrast red/coral in dark mode
    assert "#F87171" in dlg.lbl_detected_bom.styleSheet()

    # Check disabled buttons in dark mode have dark slate background, not bright yellow or glowing purple
    assert "#1E293B" in dlg.btn_run_plm.styleSheet()
    assert "#64748B" in dlg.btn_run_plm.styleSheet()
    assert "#1E293B" in dlg.btn_run_r3.styleSheet()
    assert "#1E293B" in dlg.btn_run_both.styleSheet()

    # 3. Simulate Light Mode
    dlg.theme_mgr._effective_theme = "light"
    dlg._apply_theme_styles()
    assert "#F8FAFC" in dlg.status_box.styleSheet()
    assert "#0F172A" in dlg.lbl_title_bom.styleSheet()

    dlg.close()
