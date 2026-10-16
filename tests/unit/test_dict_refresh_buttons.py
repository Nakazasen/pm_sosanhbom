"""Unit tests verifying all Refresh/Reload Excel buttons across Leader, Member, Dialogs."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.leader_view import LeaderWorkspaceView, ModelMachineManagerDialog
from src.gui.member_view import MemberWorkspaceView
from src.gui.member_roster_dialog import MemberRosterDialog
from src.gui.pcd_plan_dialog import PCDPlanScanDialog
from src.services.machine_dict_service import MachineDictService


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def mock_excel_file(tmp_path: Path) -> Path:
    f = tmp_path / "mock_dict.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(["Tên máy", "Kiểu", "Thương hiệu", "Mã máy", "Thông số kĩ thuật", "Nhóm", "Email"])
    ws.append(["Virgo", "MFP", "KDC", "02YJ; 02Z0", "250", "To", "pe_to@kdcf.com"])
    ws.append(["6th Next", "6th Next", "H", "0C0T", "", "CC1", "pe_cc1@kdcf.com"])
    wb.save(str(f))
    wb.close()
    return f


def test_leader_view_reload_dict_button(qapp, tmp_path: Path, mock_excel_file: Path):
    """Test reload dict button in LeaderWorkspaceView Step 1.1."""
    dict_svc = MachineDictService(excel_path=mock_excel_file)
    dict_svc.load()

    with patch("src.gui.leader_view.get_machine_dict_service", return_value=dict_svc):
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        assert hasattr(step1, "btn_reload_dict")
        assert "Tải lại từ điển" in step1.btn_reload_dict.text()

        # Modify mock excel file to add new model
        wb = openpyxl.load_workbook(str(mock_excel_file))
        ws = wb["Master"]
        ws.append(["NewModel2026", "MFP", "KDC", "0D99", "", "CC2", "pe_cc2@kdcf.com"])
        wb.save(str(mock_excel_file))
        wb.close()

        # Click reload with QMessageBox mocked
        with patch("PyQt6.QtWidgets.QMessageBox.information") as mock_info:
            step1.btn_reload_dict.click()
            assert mock_info.called

        # Verify model combo now contains the new model
        combo_items = [step1.model_combo.itemText(i) for i in range(step1.model_combo.count())]
        assert "NewModel2026" in combo_items


def test_model_machine_manager_dialog_reload_buttons(qapp, tmp_path: Path, mock_excel_file: Path):
    """Test reload buttons inside ModelMachineManagerDialog."""
    dict_svc = MachineDictService(excel_path=mock_excel_file)
    dict_svc.load()

    dlg = ModelMachineManagerDialog(dict_service=dict_svc)
    assert hasattr(dlg, "btn_quick_reload")
    assert "Tải lại Excel" in dlg.btn_quick_reload.text()
    assert hasattr(dlg, "btn_refresh_dict")
    assert "Tải Lại Từ Excel" in dlg.btn_refresh_dict.text()

    with patch("PyQt6.QtWidgets.QMessageBox.information") as mock_info:
        dlg._on_refresh_dict_clicked()
        assert mock_info.called

    dlg.close()


def test_member_view_reload_models_button(qapp, tmp_path: Path, mock_excel_file: Path):
    """Test reload models button on MemberWorkspaceView toolbar."""
    dict_svc = MachineDictService(excel_path=mock_excel_file)
    dict_svc.load()

    with patch("src.gui.member_view.get_machine_dict_service", return_value=dict_svc):
        view = MemberWorkspaceView()
        assert hasattr(view, "btn_reload_models")
        assert view.btn_reload_models is not None

        with patch("PyQt6.QtWidgets.QMessageBox.information") as mock_info:
            view._reload_models_from_excel()
            assert mock_info.called


def test_member_roster_dialog_refresh_button(qapp, tmp_path: Path):
    """Test refresh button in MemberRosterDialog."""
    dlg = MemberRosterDialog(base_dir=tmp_path)
    assert hasattr(dlg, "btn_import_excel")
    assert "Nạp lại từ Excel (Refresh)" in dlg.btn_import_excel.text()
    dlg.close()


def test_pcd_plan_dialog_reload_dict_button(qapp, mock_excel_file: Path):
    """Test reload dict button in PCDPlanScanDialog."""
    dict_svc = MachineDictService(excel_path=mock_excel_file)
    dict_svc.load()

    dlg = PCDPlanScanDialog(dict_service=dict_svc)
    assert hasattr(dlg, "btn_reload_dict")
    with patch("PyQt6.QtWidgets.QMessageBox.information") as mock_info:
        dlg._reload_dictionary()
        assert mock_info.called
    dlg.close()
