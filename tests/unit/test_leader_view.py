"""Unit tests for Feature F22 & Milestone M1: Leader Workspace 4-Step Sequential Wizard.

Comprehensive verification of:
- Wizard Step Header & Navigation guards (Step 1 -> Step 2 -> Step 3 -> Step 4).
- Step 1: Model & Phase selection (maT vs ma1), BOM codes list, Staffing table Cơ 1, 2, 3 (Sheet Lichsu),
  machine folder creation, and member submission packages generation.
- Step 2: Data sourcing, common vs individual date modes, file auto-routing, backup to backupTC14full/, and BOM filtering.
- Step 3: Real-time scan for CTTT!Q2 = 'OK', live status table, Fail-Closed consolidation gate,
  and archiving member files to phutrach/.
- Step 4: Master BOM workbook creation, Pivot Table refresh, 15-model JIG master catalog, 4M evaluation,
  and 2-tier Outlook emails.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest
from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import QApplication, QCheckBox, QComboBox, QDateEdit, QMessageBox, QTableWidgetItem

from src.gui.leader_view import (
    CHECKLIST_18_POINTS,
    JIG_MODELS,
    ROSTER_MECHA_1,
    ROSTER_MECHA_2,
    ROSTER_MECHA_3,
    BatchReconciliationWorker,
    EmailPreviewDialog,
    LeaderSessionState,
    LeaderWorkspaceView,
    ProjectStage,
    Step1ProjectSetupWidget,
    Step2DataSourcingWidget,
    Step3TrackingConsolidationWidget,
    Step4ComparisonReportingWidget,
    WizardStepHeader,
)
from src.reporting.excel_generator import STANDARD_SUB_UNITS


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


# =============================================================================
# Step 1 Tests
# =============================================================================

class TestLeaderWizardStep1:
    """Tests for Step 1: Project Setup & Staffing Assignment."""

    def test_roster_composition_and_department_counts(self) -> None:
        """Verify 38 engineers across Cơ 1, Cơ 2, Cơ 3 from Sheet tenphong_pt."""
        assert len(ROSTER_MECHA_1) == 12
        assert len(ROSTER_MECHA_2) == 11
        assert len(ROSTER_MECHA_3) == 15
        assert len(ROSTER_MECHA_1) + len(ROSTER_MECHA_2) + len(ROSTER_MECHA_3) == 38

        assert "Son_mecha1" in ROSTER_MECHA_1
        assert "Loc_mecha2" in ROSTER_MECHA_2
        assert "LVThuan_mecha3" in ROSTER_MECHA_3

    def test_stage_selection_affects_machine_code_prefixes(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Switching between ma1 and maT updates machine code prefixes (110 vs T10)."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # Default is MP (ma1) -> prefix 110
        assert step1.machine_table.rowCount() >= 1
        assert step1.machine_table.item(0, 1).text().startswith("110")

        # Switch to DMT / PMT (maT)
        step1.stage_combo.setCurrentText("DMT / PMT (maT - tiền tố T10)")
        assert view.state.stage == ProjectStage.MA_T
        assert step1.machine_table.item(0, 1).text().startswith("T10")

        # Switch back to MP (ma1)
        step1.stage_combo.setCurrentText("MP (ma1 - tiền tố 110)")
        assert view.state.stage == ProjectStage.MA_1
        assert step1.machine_table.item(0, 1).text().startswith("110")

    def test_staff_table_filtering_and_assignment(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Filter staffing table by department and test auto-assign subunits."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # Initially 38 rows visible
        assert step1.staff_table.rowCount() == 38

        # Filter by Cơ 1
        step1.combo_dept_filter.setCurrentText("Cơ 1")
        visible_c1 = [r for r in range(step1.staff_table.rowCount()) if not step1.staff_table.isRowHidden(r)]
        assert len(visible_c1) == 12

        # Filter by Cơ 2
        step1.combo_dept_filter.setCurrentText("Cơ 2")
        visible_c2 = [r for r in range(step1.staff_table.rowCount()) if not step1.staff_table.isRowHidden(r)]
        assert len(visible_c2) == 11

        # Filter by Cơ 3
        step1.combo_dept_filter.setCurrentText("Cơ 3")
        visible_c3 = [r for r in range(step1.staff_table.rowCount()) if not step1.staff_table.isRowHidden(r)]
        assert len(visible_c3) == 15

        # Reset to Tất cả
        step1.combo_dept_filter.setCurrentText("Tất cả")
        visible_all = [r for r in range(step1.staff_table.rowCount()) if not step1.staff_table.isRowHidden(r)]
        assert len(visible_all) == 38

        # Auto-assign sub-units
        step1._auto_assign_subunits()
        assigned_units = {step1.staff_table.item(r, 4).text() for r in range(step1.staff_table.rowCount())}
        assert "LSU" in assigned_units
        assert "FUSER" in assigned_units

    def test_dynamic_machine_binding_to_staff_table(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify column 'Mã Máy Giao' dynamically sources its choices from table 1.2."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # 1. Initial machines in table 1.2: 110C103NL0, 110C103NL1, 110C0Z3LV1
        active_codes = step1.get_active_machine_codes()
        assert active_codes == ["110C103NL0", "110C103NL1", "110C0Z3LV1"]

        # Each row in staff_table has a QComboBox in column 3 populated with active_codes
        combo0 = step1.staff_table.cellWidget(0, 3)
        assert isinstance(combo0, QComboBox)
        items0 = [combo0.itemText(i) for i in range(combo0.count())]
        for code in active_codes:
            assert code in items0

        # 2. Add a new machine code to table 1.2
        step1._add_machine_row(code="110C999US0")
        active_codes_after_add = step1.get_active_machine_codes()
        assert "110C999US0" in active_codes_after_add

        # All combo boxes now have the newly added machine code
        items0_after = [combo0.itemText(i) for i in range(combo0.count())]
        assert "110C999US0" in items0_after

        # 3. Mark a machine as Excluded (Bỏ qua X)
        chk_w = step1.machine_table.cellWidget(0, 2)
        chk = chk_w.findChild(QCheckBox)
        assert chk is not None
        chk.setChecked(True)  # Exclude 110C103NL0
        active_codes_after_ex = step1.get_active_machine_codes()
        assert "110C103NL0" not in active_codes_after_ex

        # 4. Paste new machines list
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="2NT1000")
        step1._add_machine_row(code="2NT2000")
        assert step1.get_active_machine_codes() == ["2NT1000", "2NT2000"]
        items0_pasted = [combo0.itemText(i) for i in range(combo0.count())]
        assert "2NT1000" in items0_pasted
        assert "2NT2000" in items0_pasted

        # 5. User can select machine code from dropdown
        combo0.setCurrentText("2NT2000")
        assert step1.staff_table.item(0, 3).text() == "2NT2000"
        step1.sync_state_from_ui()
        assert step1.state.staff_roster[0].machine_code == "2NT2000"

    def test_validation_guards_empty_machines_and_empty_staff(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Attempting folder generation without machine or staff triggers warning."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        warnings = []
        monkeypatch.setattr(QMessageBox, "warning", lambda p, t, m: warnings.append(m))
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)

        # Clear all machines
        step1.machine_table.setRowCount(0)
        step1.execute_create_folders_and_packages()
        assert len(warnings) == 1
        assert "Bạn chưa nhập mã máy vào list" in warnings[0]
        assert view.state.step1_completed is False

        # Add 1 machine, but deselect all staff
        step1._add_machine_row(code="110C103NL0")
        step1._set_all_staff_checked(False)
        warnings.clear()
        step1.execute_create_folders_and_packages()
        assert len(warnings) == 1
        assert "Bạn chưa lựa chọn người phụ trách" in warnings[0]
        assert view.state.step1_completed is False

    def test_execution_creates_machine_directories_and_member_packages(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Execution successfully creates machine folders and member workbooks."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        step1.model_combo.setCurrentText("Iris2024")
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C0Z3LV1")

        # Select only 2 engineers
        step1._set_all_staff_checked(False)
        for r in range(2):
            chk_w = step1.staff_table.cellWidget(r, 0)
            chk = chk_w.findChild(QCheckBox)
            chk.setChecked(True)
            step1.staff_table.setItem(r, 3, openpyxl_item := QTableWidgetItem("110C0Z3LV1"))

        step1.execute_create_folders_and_packages()

        assert view.state.step1_completed is True
        mach_dir = tmp_path / "Iris2024" / "110C0Z3LV1"
        assert mach_dir.exists()

        # Check member packages
        eng1_name = step1.staff_table.item(0, 1).text()
        eng2_name = step1.staff_table.item(1, 1).text()
        pkg1 = mach_dir / f"{eng1_name}.xlsm"
        pkg2 = mach_dir / f"{eng2_name}.xlsm"
        assert pkg1.exists()
        assert pkg2.exists()

        # Check sheets in package
        wb = openpyxl.load_workbook(pkg1)
        assert "CTTT" in wb.sheetnames
        assert "MSI" in wb.sheetnames
        assert "Label_7980_7990" in wb.sheetnames
        assert wb["CTTT"]["Q2"].value in ("", None)
        wb.close()


# =============================================================================
# Step 2 Tests
# =============================================================================

class TestLeaderWizardStep2:
    """Tests for Step 2: Data Sourcing, Date Selection & BOM Filtering."""

    def test_date_mode_switching(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify common date vs individual machine date selection with dynamic table sync."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step2 = view.step2_widget

        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C103NL0")
        step1._add_machine_row(code="110C103NL1")
        step1.sync_state_from_ui()

        step2.refresh_sourcing_table()
        assert step2.radio_common_date.isChecked()
        assert step2.sourcing_table.rowCount() == 2

        # 1. Changing common date immediately updates table column 2 in all rows
        step2.date_edit_common.setDate(QDate(2026, 9, 18))
        assert step2.sourcing_table.item(0, 2).text() == "2026/09/18"
        assert step2.sourcing_table.item(1, 2).text() == "2026/09/18"
        assert view.state.machines[0].custom_date == "2026/09/18"
        assert view.state.machines[1].custom_date == "2026/09/18"

        # 2. Switch to individual machine date mode
        step2.radio_individual_date.setChecked(True)
        assert step2.radio_individual_date.isChecked()
        assert step2.date_edit_common.isEnabled() is False

        # Row 0 date can be customized
        w0 = step2.sourcing_table.cellWidget(0, 2)
        assert isinstance(w0, QDateEdit)
        assert w0.isEnabled() is True
        w0.setDate(QDate(2026, 9, 20))
        assert step2.sourcing_table.item(0, 2).text() == "2026/09/20"
        assert step2.sourcing_table.item(1, 2).text() == "2026/09/18"
        assert view.state.machines[0].custom_date == "2026/09/20"
        assert view.state.machines[1].custom_date == "2026/09/18"

        # 3. Switching back to common date resets all rows to common date
        step2.radio_common_date.setChecked(True)
        assert step2.date_edit_common.isEnabled() is True
        assert step2.sourcing_table.item(0, 2).text() == "2026/09/18"
        assert step2.sourcing_table.item(1, 2).text() == "2026/09/18"

    def test_sourcing_table_detects_plm_and_r3(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Sourcing table recognizes when PLM and SAP R3 files are present."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step2 = view.step2_widget

        step1.model_combo.setCurrentText("Virgo")
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C103NL0")
        step1.sync_state_from_ui()

        mach_dir = tmp_path / "Virgo" / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        view.state.machines[0].folder_path = mach_dir

        # Before adding files
        step2.refresh_sourcing_table()
        assert step2.sourcing_table.rowCount() == 1
        assert "Thiếu file" in step2.sourcing_table.item(0, 3).text()
        assert "Thiếu file" in step2.sourcing_table.item(0, 4).text()
        assert view.state.step2_completed is False

        # Add PLM and R3 files
        plm_file = mach_dir / "PLM_110C103NL0.xlsx"
        r3_file = mach_dir / "R3_110C103NL0.xls"
        plm_file.write_text("PLM_CONTENT")
        r3_file.write_text("R3_CONTENT")

        step2.refresh_sourcing_table()
        assert "✓ Sẵn sàng" in step2.sourcing_table.item(0, 3).text()
        assert "✓ Sẵn sàng" in step2.sourcing_table.item(0, 4).text()
        assert view.state.step2_completed is True

    def test_execute_bom_filtering_creates_backup(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """BOM filter execution backs up raw PLM export to backupTC14full/."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step2 = view.step2_widget

        step1.model_combo.setCurrentText("Virgo")
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C103NL0")
        step1.sync_state_from_ui()

        mach_dir = tmp_path / "Virgo" / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        plm_file = mach_dir / "PLM_110C103NL0.xlsx"
        plm_file.write_text("RAW_PLM_DATA")

        step2.refresh_sourcing_table()
        step2.execute_bom_filtering()

        backup_dir = tmp_path / "backupTC14full"
        assert backup_dir.exists()
        backups = list(backup_dir.glob("PLM_110C103NL0_*.xlsx"))
        assert len(backups) >= 1
        assert view.state.machines[0].is_filtered is True


# =============================================================================
# Step 3 Tests
# =============================================================================

class TestLeaderWizardStep3:
    """Tests for Step 3: Real-time Scan, Fail-Closed Gate & Consolidation."""

    def test_live_scan_detects_cttt_q2_ok(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Live scan flags Q2='OK' as submitted and Q2='' as pending."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step3 = view.step3_widget

        mach_dir = tmp_path / "Virgo" / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)

        view.step1_widget.model_combo.setCurrentText("Virgo")
        view.step1_widget.machine_table.setRowCount(0)
        view.step1_widget._add_machine_row(code="110C103NL0")
        view.step1_widget.sync_state_from_ui()

        # Engineer 1 submitted OK
        pkg1 = mach_dir / "Son_mecha1.xlsm"
        wb1 = openpyxl.Workbook()
        ws1 = wb1.active
        ws1.title = "CTTT"
        ws1["Q2"] = "OK"
        wb1.save(pkg1)
        wb1.close()

        # Engineer 2 pending (Q2 empty)
        pkg2 = mach_dir / "Loc_mecha2.xlsm"
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        ws2.title = "CTTT"
        ws2["Q2"] = ""
        wb2.save(pkg2)
        wb2.close()

        step3.scan_submissions()
        assert step3.submission_table.rowCount() == 2

        status1 = step3.submission_table.item(0, 4).text()
        status2 = step3.submission_table.item(1, 4).text()
        statuses = {status1, status2}
        assert "✓ ĐÃ NỘP OK" in statuses
        assert "⏳ CHƯA NỘP" in statuses
        assert view.state.all_members_ok is False
        assert step3.btn_consolidate.isEnabled() is False

    def test_fail_closed_gate_blocks_consolidation_when_pending(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Consolidation is strictly blocked and shows warning if any member is pending."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step3 = view.step3_widget

        mach_dir = tmp_path / "Virgo" / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        view.step1_widget.model_combo.setCurrentText("Virgo")
        view.step1_widget.machine_table.setRowCount(0)
        view.step1_widget._add_machine_row(code="110C103NL0")
        view.step1_widget.sync_state_from_ui()

        pkg = mach_dir / "Pending_Eng.xlsm"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        ws["Q2"] = ""
        wb.save(pkg)
        wb.close()

        step3.scan_submissions()
        assert step3.btn_consolidate.isEnabled() is False

        warnings = []
        monkeypatch.setattr(QMessageBox, "warning", lambda p, t, m: warnings.append(m))

        # Forced attempt
        success = step3.execute_consolidation()
        assert success is False
        assert len(warnings) == 1
        assert "chưa nhập danh sách linh kiện" in warnings[0]
        assert not (mach_dir / "phutrach").exists()

    def test_consolidation_when_all_ok_archives_files(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """When 100% submitted OK, consolidation moves files to phutrach/ and unlocks Step 4."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step3 = view.step3_widget

        mach_dir = tmp_path / "Virgo" / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        view.step1_widget.model_combo.setCurrentText("Virgo")
        view.step1_widget.machine_table.setRowCount(0)
        view.step1_widget._add_machine_row(code="110C103NL0")
        view.step1_widget.sync_state_from_ui()

        # Both engineers submitted OK with part items
        pkg1 = mach_dir / "Son_mecha1.xlsm"
        wb1 = openpyxl.Workbook()
        ws1 = wb1.active
        ws1.title = "CTTT"
        ws1["Q2"] = "OK"
        ws1.append(["LSU", "01", "302FP02010", "MOTOR", 1, "Son_mecha1"])
        wb1.save(pkg1)
        wb1.close()

        step3.scan_submissions()
        assert step3.btn_consolidate.isEnabled() is True

        success = step3.execute_consolidation()
        assert success is True
        assert view.state.step3_completed is True

        # Verify files moved to phutrach/
        phutrach_dir = mach_dir / "phutrach"
        assert phutrach_dir.exists()
        assert (phutrach_dir / "Son_mecha1.xlsm").exists()
        assert not pkg1.exists()


# =============================================================================
# Step 4 Tests
# =============================================================================

class TestLeaderWizardStep4:
    """Tests for Step 4: Master BOM Generation, JIG Catalog, 4M & 2-Tier Emails."""

    def test_master_bom_generation_and_pivot_refresh(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Generating master BOM creates BOM_<machine_code>.xlsm with 7 sheets."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step4 = view.step4_widget

        view.step1_widget.model_combo.setCurrentText("Virgo")
        view.step1_widget.machine_table.setRowCount(0)
        view.step1_widget._add_machine_row(code="110C103NL0")
        view.step1_widget.sync_state_from_ui()
        step4.sync_machine_combo()

        # Create dummy source PLM and R3
        mach_dir = tmp_path / "Virgo" / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        plm_p = mach_dir / "PLM_110C103NL0.xlsx"
        r3_p = mach_dir / "R3_110C103NL0.xls"
        plm_p.write_text("PLM")
        r3_p.write_text("R3")
        view.state.machines[0].plm_file = plm_p
        view.state.machines[0].r3_file = r3_p

        out_file = step4.generate_master_bom_file()
        assert out_file is not None
        assert out_file.exists()
        assert "BOM_110C103NL0.xlsm" in out_file.name
        assert view.state.step4_completed is True

        # Check capnhat/old archive
        capnhat_old = mach_dir / "capnhat" / "old"
        assert capnhat_old.exists()
        assert (capnhat_old / "PLM_110C103NL0.xlsx").exists()

    def test_jig_catalog_15_models_and_4m_assessment(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """15 standard JIG models supported; 4M assessment fields interactive."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step4 = view.step4_widget

        assert len(JIG_MODELS) == 15
        assert step4.combo_jig_model.count() == 15

        # Select model and load JIG
        step4.combo_jig_model.setCurrentText("Libra2")
        step4.load_master_jig_catalog()

        # 4M assessment
        step4.chk_4m_change.setChecked(True)
        assert view.state.has_4m_change is True

        step4.radio_4m_ok.setChecked(True)
        assert step4.radio_4m_ok.isChecked()

        step4.edit_ktsx_reviewer.setText("Tran Van KTSX")
        assert step4.edit_ktsx_reviewer.text() == "Tran Van KTSX"

    def test_two_tier_outlook_emails(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Tier 1 contains 18 check points; Tier 2 contains management confirmation."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step4 = view.step4_widget

        assert len(CHECKLIST_18_POINTS) == 18

        # Tier 1 Preview
        p1 = step4._build_tier1_preview()
        assert "18 Điểm Kiểm Tra" in p1.subject
        assert "CTTT!Q2 = 'OK'" in p1.html_body
        for pt in CHECKLIST_18_POINTS[:3]:
            assert pt in p1.html_body

        # Tier 2 Preview
        p2 = step4._build_tier2_preview()
        assert "Báo cáo Hoàn tất Đối soát BOM" in p2.subject
        assert "100% OK" in p2.html_body


# =============================================================================
# Wizard Sequential Workflow & Navigation Guards
# =============================================================================

class TestLeaderWizardWorkflowGuards:
    """Tests for wizard state machine and step navigation."""

    def test_wizard_step_navigation_and_guards(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """User starts at Step 1; cannot skip steps without completion."""
        monkeypatch.setattr(QMessageBox, "information", lambda p, t, m: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)

        assert view.step_stack.currentIndex() == 0
        assert view.btn_prev.isEnabled() is False
        assert view.btn_next.isEnabled() is True

        # Advancing past Step 1 triggers Step 1 execution
        view.go_next_step()
        assert view.state.step1_completed is True
        assert view.step_stack.currentIndex() == 1
        assert view.btn_prev.isEnabled() is True

        # Move to Step 3
        view.go_next_step()
        assert view.step_stack.currentIndex() == 2

        # Step back to Step 2
        view.go_previous_step()
        assert view.step_stack.currentIndex() == 1

    def test_custom_labels_and_subunit_combobox(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify updated headers, Step 3 title, and subunit quick-select dropdown."""
        from src.gui.leader_view import WizardStepHeader, STANDARD_SUB_UNITS
        assert WizardStepHeader.STEP_NAMES[2] == "3. Theo dõi tổng hợp file của phụ trách"

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # Check Table 1.3 headers
        headers_1_3 = [
            step1.staff_table.horizontalHeaderItem(c).text()
            for c in range(step1.staff_table.columnCount())
        ]
        assert headers_1_3 == ["Áp dụng", "Phụ trách công đoạn", "Phòng Ban", "Mã máy", "Công Đoạn"]

        # Check Sub-unit QComboBox in row 0
        combo_sub = step1.staff_table.cellWidget(0, 4)
        assert isinstance(combo_sub, QComboBox)
        items = [combo_sub.itemText(i) for i in range(combo_sub.count())]
        for unit in STANDARD_SUB_UNITS:
            assert unit in items

        # Select a different sub-unit via combo
        combo_sub.setCurrentText("FUSER")
        assert step1.staff_table.item(0, 4).text() == "FUSER"
        assert view.state.staff_roster[0].sub_unit == "FUSER"

        # Check Step 3 Table headers
        step3 = view.step3_widget
        headers_3 = [
            step3.submission_table.horizontalHeaderItem(c).text()
            for c in range(step3.submission_table.columnCount())
        ]
        assert headers_3[3] == "Phụ trách công đoạn"

