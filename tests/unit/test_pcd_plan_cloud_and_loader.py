"""Unit tests for PCD Plan cloud checks, file discovery, and multi-engine parsing."""

from __future__ import annotations

from pathlib import Path
import openpyxl
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.services.machine_dict_service import MachineDictService
from src.services.pcd_plan_service import (
    PCDPlanReadError,
    PCDPlanService,
    check_file_accessible,
    is_cloud_placeholder,
    is_onedrive_running,
)
from src.gui.pcd_plan_dialog import PCDPlanScanDialog


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def sample_dict_excel(tmp_path: Path) -> Path:
    f = tmp_path / "mock_dict.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(["Tên máy", "Kiểu", "Thương hiệu", "Mã máy", "Thông số", "Loại", "Email"])
    ws.append(["Virgo", "MFP", "KDC", "02YJ; 02Z0", "", "To", "test@kdcf.com"])
    wb.save(str(f))
    wb.close()
    return f


@pytest.fixture
def sample_plan_excel(tmp_path: Path) -> Path:
    f = tmp_path / "202609月度定期計画後明細計画.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "詳細日程"
    ws.append(["STT", "Plant", "Line", "Material", "Desc", "Date", "Qty", "履歴"])
    ws.append([1, "2200", "L1", "11002YJNL0", "Virgo MP", "2026-09-20", 100, "NEW"])
    ws.append([2, "2200", "L2", "T1002Z0NL0", "Virgo Trial", "2026-09-22", 20, "NEW"])
    ws.append([3, "2200", "L1", "11002YJNL0", "Virgo Old", "2026-09-25", 50, "OLD"])
    wb.save(str(f))
    wb.close()
    return f


class TestPCDPlanDiscoveryAndFiltering:
    def test_find_monthly_plan_file_excludes_past_year_and_quality_reports(self, tmp_path: Path):
        """Must strictly exclude past year (2024) and non-plan reports (品質状況) when searching 2026/09."""
        # Create non-plan quality status file from 2024
        quality_file = tmp_path / "2024年7月～9月機器製造 品質状況 Tinh hinh chat luong che tao may 2024(quý 2).xlsx"
        quality_file.write_text("dummy", encoding="utf-8")

        # Create kaizo file
        kaizo_file = tmp_path / "2026.09 Lưu trình kaizo.xlsx"
        kaizo_file.write_text("dummy", encoding="utf-8")

        svc = PCDPlanService(base_dir=tmp_path)
        # Search for year 2026, month 9 -> Quality file from 2024 must NOT be picked!
        found = svc.find_monthly_plan_file(year=2026, month=9)
        assert found is None

        # Now create the actual valid 2026 plan file
        plan_file = tmp_path / "202609月度定期計画後明細計画.xlsx"
        plan_file.write_text("dummy", encoding="utf-8")

        found2 = svc.find_monthly_plan_file(year=2026, month=9)
        assert found2 is not None
        matched_path, ftype = found2
        assert matched_path.name == "202609月度定期計画後明細計画.xlsx"
        assert ftype == "end_of_period"

    def test_check_file_accessible_validations(self, tmp_path: Path):
        """Validate accessibility checks for non-existent, url shortcuts, and 0-byte files."""
        # 1. Non-existent file
        ok, reason = check_file_accessible(tmp_path / "not_there.xlsx")
        assert ok is False
        assert "không tồn tại" in reason

        # 2. URL shortcut
        url_file = tmp_path / "test.xlsx.url"
        url_file.write_text("https://example.com", encoding="utf-8")
        ok_url, reason_url = check_file_accessible(url_file)
        assert ok_url is False
        assert "lối tắt Internet" in reason_url

        # 3. 0-byte file
        empty_file = tmp_path / "empty.xlsx"
        empty_file.touch()
        ok_empty, reason_empty = check_file_accessible(empty_file)
        assert ok_empty is False
        assert "0 byte" in reason_empty

    def test_parse_plan_file_success(self, sample_dict_excel: Path, sample_plan_excel: Path):
        """Parse valid plan file and extract NEW materials correctly."""
        dict_svc = MachineDictService(excel_path=sample_dict_excel)
        plan_svc = PCDPlanService(dict_service=dict_svc)

        result = plan_svc.parse_plan_file(sample_plan_excel)
        assert result.file_type == "end_of_period"
        assert len(result.items) == 2

        # Item 1: MP
        assert result.items[0].material_code == "11002YJNL0"
        assert result.items[0].machine_code_4char == "02YJ"
        assert result.items[0].suggested_phase == "MP"
        assert result.items[0].is_trial is False
        assert result.items[0].machine_name == "Virgo"

        # Item 2: Trial
        assert result.items[1].material_code == "T1002Z0NL0"
        assert result.items[1].machine_code_4char == "02Z0"
        assert result.items[1].suggested_phase == "DMT/PMT"
        assert result.items[1].is_trial is True
        assert result.items[1].machine_name == "Virgo"

    def test_dialog_auto_detect_resets_label_when_no_file_found(self, qapp: QApplication, tmp_path: Path):
        """Dialog clears input field and shows clear notice if no file matches."""
        plan_svc = PCDPlanService(base_dir=tmp_path)
        dlg = PCDPlanScanDialog(plan_service=plan_svc)

        # Force search for 2026/09 in empty tmp_path
        dlg.spin_year.setValue(2026)
        dlg.combo_month.setCurrentIndex(8)  # Month 9
        dlg._auto_detect_plan()

        assert dlg.edit_plan_file.text() == ""
        assert "Chưa tìm thấy file kế hoạch" in dlg.lbl_file_type.text()

    def test_dialog_smart_error_handling_when_cloud_placeholder(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        """Dialog detects cloud-only placeholder and prompts to start OneDrive."""
        from PyQt6.QtWidgets import QMessageBox

        plan_svc = PCDPlanService(base_dir=tmp_path)
        dlg = PCDPlanScanDialog(plan_service=plan_svc)

        fake_file = tmp_path / "cloud_plan.xlsx"
        fake_file.write_text("test", encoding="utf-8")
        dlg.edit_plan_file.setText(str(fake_file))

        import src.services.pcd_plan_service as pcd_mod
        monkeypatch.setattr(pcd_mod, "is_cloud_placeholder", lambda p: True)
        monkeypatch.setattr(pcd_mod, "is_onedrive_running", lambda: False)

        question_asked = []
        monkeypatch.setattr(
            QMessageBox,
            "question",
            lambda *args, **kwargs: (
                question_asked.append(args),
                QMessageBox.StandardButton.No,
            )[1],
        )

        dlg._on_scan_clicked()
        assert len(question_asked) == 1

    def test_official_pcd_sharepoint_url_and_open_helper(self, monkeypatch: pytest.MonkeyPatch):
        """Verify official SharePoint URL and browser open helper."""
        from src.services.pcd_plan_service import OFFICIAL_PCD_SHAREPOINT_URL, open_pcd_sharepoint_url
        import webbrowser

        assert "kdcf.sharepoint.com" in OFFICIAL_PCD_SHAREPOINT_URL
        assert "sites/kdtvn_PCD" in OFFICIAL_PCD_SHAREPOINT_URL
        assert "ProductionPlan" in OFFICIAL_PCD_SHAREPOINT_URL
        assert "Theo%20th%C3%A1ng" in OFFICIAL_PCD_SHAREPOINT_URL

        opened_urls = []
        monkeypatch.setattr(webbrowser, "open", lambda url: (opened_urls.append(url), True)[1])
        res = open_pcd_sharepoint_url()
        assert res is True
        assert opened_urls == [OFFICIAL_PCD_SHAREPOINT_URL]

    def test_dialog_has_sharepoint_button_and_hint(self, qapp: QApplication, tmp_path: Path):
        """Dialog UI contains '🌐 Mở SharePoint PCD...' button and helpful auto-sync hint."""
        plan_svc = PCDPlanService(base_dir=tmp_path)
        dlg = PCDPlanScanDialog(plan_service=plan_svc)

        assert hasattr(dlg, "btn_open_sharepoint")
        assert "SharePoint" in dlg.btn_open_sharepoint.text()

    def test_configured_pcd_base_dir_priority(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        """PCDPlanService uses pcd_base_dir from settings.json if present."""
        import json
        import src.services.pcd_plan_service as pcd_mod

        custom_dir = tmp_path / "custom_pcd_sharepoint"
        plan_sub = custom_dir / "2026" / "202609"
        plan_sub.mkdir(parents=True, exist_ok=True)
        valid_plan = plan_sub / "202609月度定期計画後明細計画.xlsx"
        valid_plan.write_text("test", encoding="utf-8")

        monkeypatch.setattr(pcd_mod, "_get_configured_pcd_base_dir", lambda: custom_dir)

        svc = PCDPlanService(base_dir=tmp_path / "empty_onedrive")
        found = svc.find_monthly_plan_file(year=2026, month=9)
        assert found is not None
        matched, ftype = found
        assert matched == valid_plan
        assert ftype == "end_of_period"
