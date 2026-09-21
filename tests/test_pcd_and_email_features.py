"""Unit tests for PCD Plan scanner, Machine Dictionary 2-way sync, and Outlook email templates."""

import os
import shutil
import tempfile
from pathlib import Path
import openpyxl
import pytest

from src.services.machine_dict_service import MachineDictService, MachineInfo
from src.services.pcd_plan_service import PCDPlanService, PCDPlanItem
from src.reporting.outlook_mailer import OutlookMailer


@pytest.fixture
def mock_dict_excel(tmp_path: Path) -> Path:
    """Create a mock file_loaimay_nhommail.xlsx file."""
    excel_file = tmp_path / "mock_file_loaimay_nhommail.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"

    # Header
    ws.append(["Tên máy", "Kiểu", "Thương hiệu", "Mã máy", "Thông số kĩ thuật", None, None])
    # Rows
    ws.append(["6th Next", "6th Next", "H", "0C0T; 0C0W; 0C0X", None, "To", "KDTVN-Production Engineering_Management(Local) <KDTVN-Production-Engineering_Management@kdcf.onmicrosoft.com>"])
    ws.append(["Libra 2", "PRT", "KDC", "0C15; 0C1F", None, "CC1(nhóm cơ 1)", "KDTVN-Production Engineering_Mecha 1(Local) <KDTVN-ProductionEngineering_Mecha1_Local_@kdcf.onmicrosoft.com>"])
    ws.append(["Virgo", "MFP", "KDC", "02YJ; 02Z0; 02Z1", "250", "CC2(nhóm cơ 2)", "KDTVN-Production Engineering_Mecha 2(Local) <KDTVN-ProductionEngineering_Mecha2_Local_@kdcf.onmicrosoft.com>"])

    wb.save(str(excel_file))
    wb.close()
    return excel_file


@pytest.fixture
def mock_pcd_plan_excel(tmp_path: Path) -> Path:
    """Create a mock 9月度定期計画後明細計画.xlsx file."""
    excel_file = tmp_path / "9月度定期計画後明細計画.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "詳細日程"

    # Header
    ws.append(["STT", "Plant", "Line", "Material", "Desc", "Date", "Qty", "履歴"])
    # Row 1: NEW with T1 (DMT/PMT)
    ws.append([1, "2200", "L1", "T10C0TZUS0", "6th Next Trial", "2026-09-25", 50, "NEW"])
    # Row 2: NEW with 11 (MP)
    ws.append([2, "2200", "L2", "110C153NL0", "Libra 2 Mass", "2026-09-28", 200, "new"])
    # Row 3: OLD with T1 (should be filtered out)
    ws.append([3, "2200", "L1", "T102YJ3AX0", "Virgo Old", "2026-09-20", 30, "OLD"])
    # Row 4: NEW with other prefix (should be filtered out)
    ws.append([4, "2200", "L3", "200C0T0010", "Part Item", "2026-09-21", 10, "NEW"])

    wb.save(str(excel_file))
    wb.close()
    return excel_file


class TestMachineDictService:
    def test_load_and_lookup(self, mock_dict_excel: Path):
        svc = MachineDictService(excel_path=mock_dict_excel)
        assert svc.load() is True

        # Test lookup by 4-char code
        m1 = svc.lookup_machine_code("0C0T")
        assert m1 is not None
        assert m1.machine_name == "6th Next"

        m2 = svc.lookup_machine_code("0C15")
        assert m2 is not None
        assert m2.machine_name == "Libra 2"

        # Non-existing code
        assert svc.lookup_machine_code("XXXX") is None

    def test_extract_and_lookup_material(self, mock_dict_excel: Path):
        svc = MachineDictService(excel_path=mock_dict_excel)
        svc.load()

        # Material with T1 prefix
        m1 = svc.extract_and_lookup_material("T10C0TZUS0")
        assert m1 is not None
        assert m1.machine_name == "6th Next"

        # Material with 11 prefix
        m2 = svc.extract_and_lookup_material("110C153NL0")
        assert m2 is not None
        assert m2.machine_name == "Libra 2"

    def test_2way_write_back(self, mock_dict_excel: Path):
        svc = MachineDictService(excel_path=mock_dict_excel)
        svc.load()

        # Append new code '0C0Z' to existing model '6th Next'
        success = svc.add_or_update_machine_code(machine_name="6th Next", new_code="0C0Z")
        assert success is True

        # Verify in memory
        m = svc.lookup_machine_code("0C0Z")
        assert m is not None
        assert m.machine_name == "6th Next"

        # Append entirely new model
        success2 = svc.add_or_update_machine_code(machine_name="Polaris Next", new_code="0C0V", variant="Polaris Next")
        assert success2 is True
        m_new = svc.lookup_machine_code("0C0V")
        assert m_new is not None
        assert m_new.machine_name == "Polaris Next"


class TestPCDPlanService:
    def test_parse_plan_file(self, mock_dict_excel: Path, mock_pcd_plan_excel: Path):
        dict_svc = MachineDictService(excel_path=mock_dict_excel)
        plan_svc = PCDPlanService(dict_service=dict_svc)

        res = plan_svc.parse_plan_file(mock_pcd_plan_excel)
        assert res.file_type == "end_of_period"
        assert len(res.items) == 2  # Only Row 1 and Row 2 match criteria

        item1 = res.items[0]
        assert item1.material_code == "T10C0TZUS0"
        assert item1.is_trial is True
        assert item1.machine_code_4char == "0C0T"
        assert item1.machine_name == "6th Next"
        assert item1.suggested_phase == "DMT/PMT"

        item2 = res.items[1]
        assert item2.material_code == "110C153NL0"
        assert item2.is_trial is False
        assert item2.machine_code_4char == "0C15"
        assert item2.machine_name == "Libra 2"
        assert item2.suggested_phase == "MP"


class TestOutlookMailerExtensions:
    def test_task_assignment_email_test_mode(self):
        mailer = OutlookMailer(test_mode=True, test_recipient="vinh.bd@dtvn.kyocera.com")
        preview = mailer.build_task_assignment_email(
            machine_type="Libra 2",
            start_date="29/07",
            quantity=7,
            phase="MP",
            deadline_copy="25.07.2024",
            deadline_verify="25.07.2024",
            attachment_path=r"\\fstvn01\Data\test_path",
        )

        assert preview.subject == 'So sánh BOM mã hàng mới "Libra 2"'
        assert preview.recipients_to == ["vinh.bd@dtvn.kyocera.com"]
        assert "CHẾ ĐỘ THỬ NGHIỆM" in preview.html_body
        assert "29/07" in preview.html_body
        assert "7" in preview.html_body
        assert "25.07.2024" in preview.html_body
        assert "18 điểm" in preview.html_body

    def test_management_review_email_test_mode(self):
        mailer = OutlookMailer(test_mode=True, test_recipient="vinh.bd@dtvn.kyocera.com")
        preview = mailer.build_management_review_email(
            machine_type="Libra 2",
            production_date="25.07.2024",
            attachment_path=r"\\fstvn01\Data\test_path",
        )

        assert preview.subject == 'So sánh BOM mã hàng mới "Libra 2"'
        assert preview.recipients_to == ["vinh.bd@dtvn.kyocera.com"]
        assert "CHẾ ĐỘ THỬ NGHIỆM" in preview.html_body
        assert "Dear các anh quản lý" in preview.html_body
        assert "25.07.2024" in preview.html_body

    def test_production_mode_recipients(self):
        mailer = OutlookMailer(test_mode=False)
        preview = mailer.build_task_assignment_email(
            machine_type="Libra 2",
            start_date="29/07",
            quantity=7,
            phase="MP",
            deadline_copy="25.07.2024",
            deadline_verify="25.07.2024",
            attachment_path=r"\\fstvn01\Data\test_path",
        )
        assert preview.recipients_to != ["vinh.bd@dtvn.kyocera.com"]
        assert any("Mecha" in r for r in preview.recipients_to)
        assert any("Management" in r for r in preview.recipients_cc)
        assert "CHẾ ĐỘ THỬ NGHIỆM" not in preview.html_body
