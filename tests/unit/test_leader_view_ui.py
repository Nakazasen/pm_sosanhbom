"""Unit tests for Milestone 2 (M2): Leader View & Member View Data-Dense Upgrade.

Traceability:
- TEST-UI-03: Data-Dense layout verification, 32px row heights, gridline enforcement.
- SPEC_UI_UX_ENTERPRISE_DASHBOARD.md: Leader KPI Cards, Member Workflow Stepper,
  Lucide SVG QIcon usage, checkmark fatal flaw fix verification.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pandas as pd
import pytest
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QIcon
from PyQt6.QtWidgets import QApplication

from src.gui.leader_view import (
    KPICardWidget,
    LeaderSessionState,
    LeaderWorkspaceView,
    MachineTarget,
    MemberSubmissionStatus,
    Step1ProjectSetupWidget,
    Step2DataSourcingWidget,
    Step3TrackingConsolidationWidget,
    Step4ComparisonReportingWidget,
)
from src.gui.member_view import (
    MemberWorkflowStepper,
    MemberWorkspaceView,
)
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_GREEN_FONT_HEX,
    COLOR_RED_FILL_HEX,
    COLOR_RED_FONT_HEX,
)


# =============================================================================
# TEST-UI-03: Data-Dense Table Standardization
# =============================================================================

class TestDataDenseTableStandardization:
    """Verify RES-UI-01 & RES-UI-03: All 7 primary tables must have 32px row height and visible grids."""

    def test_leader_view_tables_have_32px_height_and_show_grid(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify Leader View's 4 tables (machine, staff, sourcing, submission) have 32px rows and visible grid."""
        view = LeaderWorkspaceView(base_dir=tmp_path)

        # 1. Machine Table (Step 1)
        machine_tbl = view.step1_widget.machine_table
        assert machine_tbl.verticalHeader().defaultSectionSize() == 32
        assert machine_tbl.verticalHeader().minimumSectionSize() == 28
        assert machine_tbl.showGrid() is True

        # 2. Staff Table (Step 1)
        staff_tbl = view.step1_widget.staff_table
        assert staff_tbl.verticalHeader().defaultSectionSize() == 32
        assert staff_tbl.verticalHeader().minimumSectionSize() == 28
        assert staff_tbl.showGrid() is True

        # 3. Sourcing Table (Step 2)
        sourcing_tbl = view.step2_widget.sourcing_table
        assert sourcing_tbl.verticalHeader().defaultSectionSize() == 32
        assert sourcing_tbl.verticalHeader().minimumSectionSize() == 28
        assert sourcing_tbl.showGrid() is True

        # 4. Submission Table (Step 3)
        sub_tbl = view.step3_widget.submission_table
        assert sub_tbl.verticalHeader().defaultSectionSize() == 32
        assert sub_tbl.verticalHeader().minimumSectionSize() == 28
        assert sub_tbl.showGrid() is True

    def test_member_view_tables_have_32px_height_and_show_grid(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify Member View's 3 tables (cttt, msi, label) have 32px rows and visible grid."""
        m_view = MemberWorkspaceView(base_dir=tmp_path)

        # 1. CTTT Table
        assert m_view.cttt_table.verticalHeader().defaultSectionSize() == 32
        assert m_view.cttt_table.verticalHeader().minimumSectionSize() == 28
        assert m_view.cttt_table.showGrid() is True

        # 2. MSI Table
        assert m_view.msi_table.verticalHeader().defaultSectionSize() == 32
        assert m_view.msi_table.verticalHeader().minimumSectionSize() == 28
        assert m_view.msi_table.showGrid() is True

        # 3. Label Table
        assert m_view.label_table.verticalHeader().defaultSectionSize() == 32
        assert m_view.label_table.verticalHeader().minimumSectionSize() == 28
        assert m_view.label_table.showGrid() is True


# =============================================================================
# Leader View KPI Cards & Quick Actions
# =============================================================================

class TestLeaderKPICardsAndQuickActions:
    """Verify 4 compact KPI cards and 3 SVG Quick Action buttons on Leader View."""

    def test_kpi_cards_initial_state(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify all 4 KPI cards are instantiated and display baseline counts."""
        view = LeaderWorkspaceView(base_dir=tmp_path)

        assert hasattr(view, "kpi_total_card")
        assert hasattr(view, "kpi_ready_card")
        assert hasattr(view, "kpi_progress_card")
        assert hasattr(view, "kpi_status_card")

        assert view.lbl_kpi_total_models.text() == "0"
        assert view.lbl_kpi_ready_models.text() == "0"
        assert view.lbl_kpi_cttt_progress.text() == "0%"
        assert view.lbl_kpi_recon_status.text() == "Chờ khởi tạo"

    def test_kpi_cards_dynamic_recalculation(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify KPI cards update correctly when models, BOM files, and submissions exist."""
        view = LeaderWorkspaceView(base_dir=tmp_path)

        # Set up 2 machines in session state
        m1 = MachineTarget(machine_code="2ND_Virgo")
        m2 = MachineTarget(machine_code="2ND_Libra")
        view.state.machines = [m1, m2]

        m1_dir = tmp_path / "2ND_Virgo"
        m1_dir.mkdir(parents=True, exist_ok=True)
        plm1 = m1_dir / "PLM_Virgo.xlsx"
        plm1.touch()
        r31 = m1_dir / "R3_Virgo.xls"
        r31.touch()
        m1.plm_file = plm1
        m1.r3_file = r31

        m2_dir = tmp_path / "2ND_Libra"
        m2_dir.mkdir(parents=True, exist_ok=True)
        plm2 = m2_dir / "PLM_Libra.xlsx"
        plm2.touch()
        m2.plm_file = plm2
        m2.r3_file = None  # Only 1 BOM ready

        # 2 submissions: 1 OK, 1 pending
        s1 = MemberSubmissionStatus(
            machine_code="2ND_Virgo", sub_unit="LSU", engineer_name="Son_mecha1", is_submitted_ok=True
        )
        s2 = MemberSubmissionStatus(
            machine_code="2ND_Libra", sub_unit="FUSER", engineer_name="Duy_mecha1", is_submitted_ok=False
        )
        view.state.submissions = [s1, s2]

        # Refresh KPI cards
        view.refresh_kpi_cards()

        assert view.lbl_kpi_total_models.text() == "2"
        assert view.lbl_kpi_ready_models.text() == "1/2"
        assert "50%" in view.lbl_kpi_cttt_progress.text()
        assert "1/2" in view.lbl_kpi_cttt_progress.text()
        assert view.lbl_kpi_recon_status.text() == "Đang chuẩn bị"

    def test_quick_action_buttons_have_valid_svg_icons(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify the 3 Quick Action buttons have valid non-null SVG QIcons."""
        view = LeaderWorkspaceView(base_dir=tmp_path)

        assert hasattr(view, "btn_quick_scan")
        assert hasattr(view, "btn_quick_report")
        assert hasattr(view, "btn_quick_reset")

        # Icons must not be null
        assert not view.btn_quick_scan.icon().isNull()
        assert not view.btn_quick_report.icon().isNull()
        assert not view.btn_quick_reset.icon().isNull()


# =============================================================================
# Member View Workflow Stepper & Diff View
# =============================================================================

class TestMemberWorkflowStepperAndDiffView:
    """Verify 3-step Workflow Stepper progression and WCAG AAA diff coloring."""

    def test_workflow_stepper_initial_state(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify Stepper starts at Step 0 (active) with steps 1 and 2 pending."""
        m_view = MemberWorkspaceView(base_dir=tmp_path)
        stepper = m_view.workflow_stepper

        assert stepper.current_step == 0
        assert stepper.step_completed_flags == [False, False, False]

    def test_workflow_stepper_progression_across_workflow(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify stepper transitions: Step 1 (BOM load) -> Step 2 (Self check) -> Step 3 (Submit)."""
        m_view = MemberWorkspaceView(base_dir=tmp_path)
        stepper = m_view.workflow_stepper

        # 1. Supply reference data -> Step 0 completed, Step 1 active
        plm_df = pd.DataFrame([{"Name": "302FP02010", "Quantity": 2.0, "Release Status": "A"}])
        r3_df = pd.DataFrame([{"material": "302FP02010", "quantity": 2.0, "revlev": "A"}])
        m_view.set_reference_data(plm_data=plm_df, r3_data=r3_df)

        assert stepper.step_completed_flags[0] is True
        assert stepper.current_step == 1

        # 2. Add CTTT row & run preliminary self check -> Step 1 completed, Step 2 active
        m_view.add_cttt_row(page="01", part_code="302FP02010", quantity=2.0)
        res = m_view.run_preliminary_self_check()
        assert res["status"] == "OK"

        assert stepper.step_completed_flags[0] is True
        assert stepper.step_completed_flags[1] is True
        assert stepper.current_step == 2

        # 3. Submit data -> Step 2 completed
        with patch.object(QApplication, "instance"):
            with patch("src.gui.member_view.QMessageBox.information"):
                target_file = m_view.submit_data()
                assert target_file is not None
                assert stepper.step_completed_flags[2] is True

        # 4. Unlock submission -> Step 2 revoked, Step 1 active
        with patch("src.gui.member_view.QMessageBox.information"):
            m_view.unlock_submission()
            assert stepper.step_completed_flags[2] is False
            assert stepper.current_step == 1

    def test_diff_view_wcag_aaa_contrast_coloring(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify WCAG AAA diff coloring across row while preserving Col 7 Excel fill colors."""
        m_view = MemberWorkspaceView(base_dir=tmp_path)

        plm_df = pd.DataFrame([
            {"Name": "302FP02010", "Quantity": 2.0, "Release Status": "A"},
            {"Name": "302FP02020", "Quantity": 4.0, "Release Status": "A"},
        ])
        r3_df = pd.DataFrame([
            {"material": "302FP02010", "quantity": 2.0, "revlev": "A"},
            {"material": "302FP02020", "quantity": 4.0, "revlev": "A"},
        ])
        m_view.set_reference_data(plm_data=plm_df, r3_data=r3_df)

        # Row 0: OK (qty 2.0 == 2.0)
        m_view.add_cttt_row(page="01", part_code="302FP02010", quantity=2.0)
        # Row 1: NG (qty 6.0 != 4.0)
        m_view.add_cttt_row(page="02", part_code="302FP02020", quantity=6.0)

        m_view.run_preliminary_self_check()

        # Row 0 (OK): Col 7 has Excel green fill; other cells have soft WCAG green #DCFCE7
        col7_ok = m_view.cttt_table.item(0, 7)
        assert col7_ok.text() == "OK"
        assert col7_ok.background().color().name().upper().endswith(COLOR_GREEN_FILL_HEX)

        col1_ok = m_view.cttt_table.item(0, 1)
        assert col1_ok.background().color().name().upper() == "#DCFCE7"
        assert col1_ok.foreground().color().name().upper() == "#064E3B"

        # Row 1 (NG): Col 7 has Excel red fill; other cells have soft WCAG red #FEE2E2
        col7_ng = m_view.cttt_table.item(1, 7)
        assert col7_ng.text() == "NG"
        assert col7_ng.background().color().name().upper().endswith(COLOR_RED_FILL_HEX)

        col1_ng = m_view.cttt_table.item(1, 1)
        assert col1_ng.background().color().name().upper() == "#FEE2E2"
        assert col1_ng.foreground().color().name().upper() == "#7F1D1D"


# =============================================================================
# Checkmark Fatal Flaw Fix & Clean Icons Verification
# =============================================================================

class TestCheckmarkFixAndCleanIcons:
    """Verify checkmark fatal flaw fix and elimination of raw emoji icons."""

    def test_checkmark_logic_uses_file_exists_not_string(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify refresh_sourcing_table directly checks plm_file.exists() and r3_file.exists()."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        step2 = view.step2_widget

        m_dir = tmp_path / "2ND_Virgo"
        m_dir.mkdir(parents=True, exist_ok=True)
        m = MachineTarget(machine_code="2ND_Virgo", folder_path=m_dir)
        view.state.machines = [m]

        # Initially no files
        step2.refresh_sourcing_table()
        assert m.plm_file is None
        assert m.r3_file is None

        # Create PLM & R3 files
        plm_f = m_dir / "PLM_Virgo.xlsx"
        plm_f.write_text("dummy plm")
        r3_f = m_dir / "R3_Virgo.xls"
        r3_f.write_text("dummy r3")

        step2.refresh_sourcing_table()
        # Verify genuine file existence check occurred
        assert m.plm_file is not None
        assert m.plm_file.exists()
        assert m.r3_file is not None
        assert m.r3_file.exists()

        # Both files confirmed ready
        ready_models = [
            mod for mod in view.state.machines
            if mod.plm_file is not None and mod.plm_file.exists()
            and mod.r3_file is not None and mod.r3_file.exists()
        ]
        assert len(ready_models) == 1

    def test_buttons_do_not_contain_raw_emojis(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify no raw emojis remain on action buttons, using QIcon instead."""
        raw_emojis = ["🚀", "🔍", "📁", "🔄", "🗑️", "💾", "➕", "➖", "🏷️", "🔒", "🔓", "📊", "👥", "⚙️"]

        leader_view = LeaderWorkspaceView(base_dir=tmp_path)
        member_view = MemberWorkspaceView(base_dir=tmp_path)

        # Check Leader buttons
        leader_buttons = [
            leader_view.step1_widget.btn_create_folders,
            leader_view.step1_widget.btn_add_mach,
            leader_view.step2_widget.btn_execute_sourcing,
            leader_view.step3_widget.btn_live_scan,
            leader_view.step3_widget.btn_consolidate,
            leader_view.step4_widget.btn_generate_master,
            leader_view.step4_widget.btn_preview_emails,
            leader_view.btn_quick_scan,
            leader_view.btn_quick_report,
            leader_view.btn_quick_reset,
        ]
        for btn in leader_buttons:
            for emoji in raw_emojis:
                assert emoji not in btn.text(), f"Found raw emoji {emoji} in {btn.text()}"

        # Check Member buttons
        member_buttons = [
            member_view.btn_add_row,
            member_view.btn_remove_row,
            member_view.btn_clear_table,
            member_view.btn_clear_msi,
            member_view.btn_apply_msi_edit,
            member_view.btn_add_label_row,
            member_view.btn_remove_label_row,
            member_view.btn_clear_label_table,
            member_view.btn_self_check,
            member_view.btn_unlock,
            member_view.btn_submit,
        ]
        for btn in member_buttons:
            for emoji in raw_emojis:
                assert emoji not in btn.text(), f"Found raw emoji {emoji} in {btn.text()}"
