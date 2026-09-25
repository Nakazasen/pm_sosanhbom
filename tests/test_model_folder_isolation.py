"""Comprehensive unit tests for Strict Model & Folder Isolation.

Verifies:
1. Table 1.2 displays 'Loại Máy' and auto-detects canonical models via MachineDictService.
2. Machines of different models (PolarisNext, Virgo, Iris2024) are strictly isolated into
   separate model directories (<Base>/<Model>/...).
3. Package creation (_execute_create_folders_and_packages) generates workbooks in respective model folders.
4. Email dispatching in LeaderView and PCDPlanScanDialog segregates machines by model and
   generates separate email previews for each model with accurate quantities and paths.
5. Signal connection from PCDPlanScanDialog populates Table 1.2 with created machines and auto-detected models.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest
from PyQt6.QtWidgets import QApplication, QCheckBox, QMessageBox

from src.gui.leader_view import LeaderWorkspaceView, MachineTarget
from src.gui.pcd_plan_dialog import PCDPlanScanDialog
from src.services.machine_dict_service import MachineDictService
from src.services.pcd_plan_service import PCDPlanItem, PCDPlanService


@pytest.fixture
def mock_dict_service(tmp_path: Path) -> MachineDictService:
    excel_file = tmp_path / "test_dict.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(["Tên máy", "Kiểu", "Thương hiệu", "Mã máy", "Thông số kĩ thuật", "To", "Email"])
    ws.append(["Polaris Next", "H", "KDC", "0C10; 0C0Z", None, "To", "pe_polaris@kdcf.com"])
    ws.append(["Virgo", "MFP", "KDC", "02YJ; 02Z0", "250", "CC1", "pe_virgo@kdcf.com"])
    ws.append(["Iris 2024", "H", "KDC", "0C2G; 0C2H", None, "CC2", "pe_iris@kdcf.com"])
    wb.save(str(excel_file))
    wb.close()

    svc = MachineDictService(excel_path=excel_file)
    svc.load()
    return svc


class TestModelFolderIsolation:
    def test_table_1_2_columns_and_auto_model_detection(
        self,
        qapp: QApplication,
        tmp_path: Path,
        mock_dict_service: MachineDictService,
    ) -> None:
        """Verify Table 1.2 has 5 columns and automatically resolves model for each machine code."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.dict_service = mock_dict_service

        assert step1.machine_table.columnCount() == 5
        headers = [step1.machine_table.horizontalHeaderItem(c).text() for c in range(5)]
        assert headers == ["STT", "Mã Máy", "Loại Máy", "Bỏ qua (X)", "Ghi chú"]

        # Add 3 codes belonging to 3 different models
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C103NL0")  # Polaris Next
        step1._add_machine_row(code="11002YJ3NL0")  # Virgo
        step1._add_machine_row(code="1100C2G3NL0")  # Iris 2024

        assert step1.machine_table.item(0, 2).text() == "PolarisNext"
        assert step1.machine_table.item(1, 2).text() == "Virgo"
        assert step1.machine_table.item(2, 2).text() == "Iris2024"

    def test_sync_state_and_folder_isolation_per_model(
        self,
        qapp: QApplication,
        tmp_path: Path,
        mock_dict_service: MachineDictService,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Verify each machine is assigned to its own model directory and folders are created separately."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.dict_service = mock_dict_service

        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C103NL0")  # PolarisNext
        step1._add_machine_row(code="11002YJ3NL0")  # Virgo
        step1._add_machine_row(code="1100C2G3NL0")  # Iris2024

        step1.sync_state_from_ui()

        machines = view.state.machines
        assert len(machines) == 3

        m_polaris = machines[0]
        m_virgo = machines[1]
        m_iris = machines[2]

        assert m_polaris.model_name == "PolarisNext"
        assert m_polaris.folder_path == tmp_path / "PolarisNext" / "110C103NL0"

        assert m_virgo.model_name == "Virgo"
        assert m_virgo.folder_path == tmp_path / "Virgo" / "11002YJ3NL0"

        assert m_iris.model_name == "Iris2024"
        assert m_iris.folder_path == tmp_path / "Iris2024" / "1100C2G3NL0"

        # Execute folder and package creation
        step1._set_all_staff_checked(False)
        chk_w = step1.staff_table.cellWidget(0, 0)
        chk = chk_w.findChild(QCheckBox)
        chk.setChecked(True)
        step1.staff_table.setItem(0, 3, openpyxl.packaging if False else None)

        step1.execute_create_folders_and_packages()

        # Folders must exist in their respective model subdirectories, never combined into one
        assert (tmp_path / "PolarisNext" / "110C103NL0").is_dir()
        assert (tmp_path / "Virgo" / "11002YJ3NL0").is_dir()
        assert (tmp_path / "Iris2024" / "1100C2G3NL0").is_dir()

        # Negative checks: None should be created under wrong models
        assert not (tmp_path / "Virgo" / "110C103NL0").exists()
        assert not (tmp_path / "PolarisNext" / "11002YJ3NL0").exists()
        assert not (tmp_path / "PolarisNext" / "1100C2G3NL0").exists()

    def test_assignment_email_segregation_by_model(
        self,
        qapp: QApplication,
        tmp_path: Path,
        mock_dict_service: MachineDictService,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Verify Step 1 task assignment email opens separate previews per model with correct machine counts."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.dict_service = mock_dict_service

        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C103NL0")  # PolarisNext 1
        step1._add_machine_row(code="110C0Z3NL0")  # PolarisNext 2
        step1._add_machine_row(code="11002YJ3NL0")  # Virgo 1

        built_emails = []

        def mock_build(machine_type, start_date, quantity, phase, deadline_copy, deadline_verify, attachment_path):
            built_emails.append({
                "machine_type": machine_type,
                "quantity": quantity,
                "attachment_path": attachment_path,
            })
            mock_preview = MagicMock()
            mock_preview.to = ["eng@test.com"]
            mock_preview.to_string = "eng@test.com"
            mock_preview.cc_string = ""
            mock_preview.attachment_paths = []
            mock_preview.subject = f"Test {machine_type}"
            mock_preview.html_body = "<p>body</p>"
            mock_preview.body_plain = "body"
            return mock_preview

        monkeypatch.setattr(step1.mailer, "build_task_assignment_email", mock_build)
        monkeypatch.setattr("src.gui.leader_view.EmailPreviewDialog.exec", lambda self: None)

        step1._on_send_assignment_email()

        # Should generate 2 separate emails: 1 for PolarisNext (2 machines) and 1 for Virgo (1 machine)
        assert len(built_emails) == 2

        polaris_mail = next(m for m in built_emails if m["machine_type"] == "PolarisNext")
        assert polaris_mail["quantity"] == 2
        assert "PolarisNext" in str(polaris_mail["attachment_path"])

        virgo_mail = next(m for m in built_emails if m["machine_type"] == "Virgo")
        assert virgo_mail["quantity"] == 1
        assert "Virgo" in str(virgo_mail["attachment_path"])

    def test_pcd_plan_dialog_multi_model_isolation_and_signal(
        self,
        qapp: QApplication,
        tmp_path: Path,
        mock_dict_service: MachineDictService,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Verify PCDPlanScanDialog isolates created folders by model and signals LeaderView to populate them."""
        from src.services.pcd_plan_service import PCDPlanScanResult
        # Setup mock PCDPlanService
        mock_pcd_service = MagicMock(spec=PCDPlanService)
        item1 = PCDPlanItem(
            material_code="110C103NL0",
            machine_code_4char="0C10",
            history_status="NEW",
            is_trial=False,
            suggested_phase="MP",
            machine_name="Polaris Next",
            raw_data={"Plant": "2200", "Line": "L1", "Desc": "Polaris Mass", "Qty": 100, "Date": "2026-09-25"},
        )
        item2 = PCDPlanItem(
            material_code="11002YJ3NL0",
            machine_code_4char="02YJ",
            history_status="NEW",
            is_trial=False,
            suggested_phase="MP",
            machine_name="Virgo",
            raw_data={"Plant": "2200", "Line": "L2", "Desc": "Virgo Mass", "Qty": 50, "Date": "2026-09-26"},
        )
        dummy_plan_file = tmp_path / "plan.xlsx"
        dummy_plan_file.write_text("dummy")
        mock_pcd_service.find_monthly_plan_file.return_value = None
        mock_pcd_service.parse_plan_file.return_value = PCDPlanScanResult(
            year=2026,
            month=9,
            plan_file=dummy_plan_file,
            file_type="end_of_period",
            items=[item1, item2],
        )

        dlg = PCDPlanScanDialog(
            plan_service=mock_pcd_service,
            dict_service=mock_dict_service,
        )
        dlg.edit_plan_file.setText(str(dummy_plan_file))
        dlg._on_scan_clicked()

        assert dlg.table.rowCount() == 2
        assert "Polaris" in dlg.table.item(0, 4).text()
        assert dlg.table.item(1, 4).text() == "Virgo"

        # Check path preview and retarget to tmp_path for test execution
        for r in range(dlg.table.rowCount()):
            p_text = dlg.table.item(r, 6).text()
            parts = Path(p_text).parts
            # Parts end with <Model>, <Phase>, <YM>, <MachineCode>
            model_part, phase_part, ym_part, code_part = parts[-4:]
            canon_model = model_part.replace(" ", "")
            test_target = tmp_path / canon_model / phase_part / ym_part / code_part
            dlg.table.item(r, 6).setText(str(test_target))

        path0 = Path(dlg.table.item(0, 6).text())
        path1 = Path(dlg.table.item(1, 6).text())
        assert "PolarisNext" in path0.parts
        assert "Virgo" in path1.parts
        assert path0.name == "110C103NL0"
        assert path1.name == "11002YJ3NL0"

        # Mock execute and check signal emission
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        monkeypatch.setattr(QMessageBox, "question", lambda p, t, m, b, d: QMessageBox.StandardButton.Yes)
        dlg.chk_send_email.setChecked(False)

        signal_folders = []
        dlg.projects_created.connect(lambda folders: signal_folders.extend(folders))

        dlg._on_execute_clicked()

        assert len(signal_folders) == 2
        assert (tmp_path / "PolarisNext").exists()
        assert (tmp_path / "Virgo").exists()

        # Connect to LeaderView to verify automatic table update
        leader_view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = leader_view.step1_widget
        step1.dict_service = mock_dict_service
        step1.machine_table.setRowCount(0)

        step1._on_pcd_projects_created(signal_folders)

        assert step1.machine_table.rowCount() == 2
        assert step1.machine_table.item(0, 1).text() == "110C103NL0"
        assert step1.machine_table.item(0, 2).text() == "PolarisNext"
        assert step1.machine_table.item(1, 1).text() == "11002YJ3NL0"
        assert step1.machine_table.item(1, 2).text() == "Virgo"

    def test_staff_assignment_by_model_name_and_package_generation(
        self,
        qapp: QApplication,
        tmp_path: Path,
        mock_dict_service: MachineDictService,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Verify engineers assigned to model M only receive package files in model M machine folders."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.dict_service = mock_dict_service

        # 1. Setup 2 machines of 2 different models
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="11002YJ3NL0", model_name="Virgo")
        step1._add_machine_row(code="110C103NL0", model_name="PolarisNext")

        # 2. Setup 3 staff members
        # Staff 0: assigned to Virgo
        # Staff 1: assigned to PolarisNext
        # Staff 2: assigned to (Tất cả)
        step1._set_all_staff_checked(False)
        for r in range(3):
            chk_w = step1.staff_table.cellWidget(r, 0)
            chk = chk_w.findChild(QCheckBox)
            chk.setChecked(True)

        combo0 = step1.staff_table.cellWidget(0, 3)
        combo0.setCurrentText("Virgo")

        combo1 = step1.staff_table.cellWidget(1, 3)
        combo1.setCurrentText("PolarisNext")

        combo2 = step1.staff_table.cellWidget(2, 3)
        combo2.setCurrentText("(Tất cả)")

        step1.sync_state_from_ui()

        eng0_name = view.state.staff_roster[0].engineer_name
        eng1_name = view.state.staff_roster[1].engineer_name
        eng2_name = view.state.staff_roster[2].engineer_name

        # 3. Execute folder and package generation
        step1.execute_create_folders_and_packages()

        virgo_dir = tmp_path / "Virgo" / "11002YJ3NL0"
        polaris_dir = tmp_path / "PolarisNext" / "110C103NL0"

        assert virgo_dir.exists()
        assert polaris_dir.exists()

        # In Virgo machine folder:
        assert (virgo_dir / f"{eng0_name}.xlsm").exists(), "Virgo-assigned engineer package must exist in Virgo folder"
        assert (virgo_dir / f"{eng2_name}.xlsm").exists(), "All-model engineer package must exist in Virgo folder"
        assert not (virgo_dir / f"{eng1_name}.xlsm").exists(), "Polaris-assigned engineer package must NOT exist in Virgo folder"

        # In Polaris machine folder:
        assert (polaris_dir / f"{eng1_name}.xlsm").exists(), "Polaris-assigned engineer package must exist in Polaris folder"
        assert (polaris_dir / f"{eng2_name}.xlsm").exists(), "All-model engineer package must exist in Polaris folder"
        assert not (polaris_dir / f"{eng0_name}.xlsm").exists(), "Virgo-assigned engineer package must NOT exist in Polaris folder"

    def test_four_stage_folder_isolation_and_dialog_email_groups(
        self,
        qapp: QApplication,
        tmp_path: Path,
        mock_dict_service: MachineDictService,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Verify strict isolation across DMT, PMT, PP, and MP folders and dialog email table operations."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        monkeypatch.setattr(QMessageBox, "question", lambda p, t, m, b, d: QMessageBox.StandardButton.Yes)

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.dict_service = mock_dict_service

        now = datetime.datetime.now()
        year_month = f"{now.year}.{now.month:02d}"

        # 1. Test PMT folder isolation
        step1.stage_combo.setCurrentText("PMT (maT - tiền tố T10)")
        pmt_dir = step1.create_model_folder_structure(model_name="Virgo", force_date_hierarchy=True)
        assert pmt_dir == tmp_path / "Virgo" / "PMT" / year_month
        assert pmt_dir.exists()
        assert not (tmp_path / "Virgo" / "DMT" / year_month).exists(), "PMT folder must not create DMT folder"

        # 2. Test DMT folder isolation
        step1.stage_combo.setCurrentText("DMT (maT - tiền tố T10)")
        dmt_dir = step1.create_model_folder_structure(model_name="Virgo", force_date_hierarchy=True)
        assert dmt_dir == tmp_path / "Virgo" / "DMT" / year_month
        assert dmt_dir.exists()

        # 3. Test PP folder isolation
        step1.stage_combo.setCurrentText("PP (ma1 - tiền tố 110)")
        pp_dir = step1.create_model_folder_structure(model_name="Virgo", force_date_hierarchy=True)
        assert pp_dir == tmp_path / "Virgo" / "PP" / year_month
        assert pp_dir.exists()

        # 4. Test MP folder isolation
        step1.stage_combo.setCurrentText("MP (ma1 - tiền tố 110)")
        mp_dir = step1.create_model_folder_structure(model_name="Virgo", force_date_hierarchy=True)
        assert mp_dir == tmp_path / "Virgo" / "MP" / year_month
        assert mp_dir.exists()

        # 5. Test ModelMachineManagerDialog stage & email table
        from src.gui.leader_view import ModelMachineManagerDialog
        dlg = ModelMachineManagerDialog(
            dict_service=mock_dict_service,
            current_model="Virgo",
            current_stage="PMT (maT - tiền tố T10)",
        )
        assert dlg.combo_stage.currentText().startswith("PMT")
        assert dlg.table_emails.rowCount() >= 1

        # Test adding and saving email group
        dlg._add_email_row()
        new_row = dlg.table_emails.rowCount() - 1
        dlg.table_emails.item(new_row, 0).setText("CC5(nhóm KTCT Điện)")
        dlg.table_emails.item(new_row, 1).setText("pe_electrical@kdcf.com")
        dlg._save_email_groups()

        # Verify saved in dict service
        all_groups = mock_dict_service.get_all_email_groups()
        assert "CC5(nhóm KTCT Điện)" in all_groups
        assert all_groups["CC5(nhóm KTCT Điện)"].clean_email == "pe_electrical@kdcf.com"

        # Test populating codes with PMT -> prefix T10
        emitted_stage = []
        emitted_codes = []
        dlg.stage_selected.connect(lambda s: emitted_stage.append(s))
        dlg.populate_codes_requested.connect(lambda m, codes: emitted_codes.extend(codes))

        dlg._on_populate_table_clicked()
        assert emitted_stage == ["PMT"]
        assert len(emitted_codes) > 0
        assert all(c.startswith("T10") for c in emitted_codes)

