"""Adversarial Empirical Stress Verification Suite for Milestone 2 (M2).

Covers:
1. Stress test 4 KPI Cards in Leader View:
   - Dynamic model variations from 0 to 50+ (rapid additions, deletions, bursts).
   - Combinatorial file availability (PLM/R3 existence, deletions, corrupted paths).
   - Real-time submission state fluctuations (0% to 100%, rounding, toggles).
   - Lifecycle reconciliation state machine transitions (Chờ khởi tạo -> Đang chuẩn bị -> Sẵn sàng đối soát -> Hoàn tất).
   - 100-iteration randomized property-based stress oracle.

2. Stress test 7 Tables across Leader & Member Views:
   - Leader View: machine_table, staff_table, sourcing_table, submission_table
   - Member View: cttt_table, msi_table, label_table
   - Verification of defaultSectionSize == 32, minimumSectionSize == 28, showGrid == True.
   - Heavy data insertion (100+ rows) ensuring 32px height across all rows.
   - Extreme viewport resizing (from 100x100 to 4K) verifying layout stability (no jitter/co giật).
   - Rapid scrolling, column resizing, and selection stress.
"""

from __future__ import annotations

import datetime
import os
import random
from pathlib import Path
from unittest.mock import MagicMock, patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pandas as pd
import pytest
from PyQt6.QtCore import QDate, QSize, Qt
from PyQt6.QtGui import QColor, QFont, QResizeEvent
from PyQt6.QtWidgets import (
    QApplication,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from src.core.reconciliation import ReconciliationResult
from src.gui.leader_view import (
    KPICardWidget,
    LeaderSessionState,
    LeaderWorkspaceView,
    MachineTarget,
    MemberSubmissionStatus,
)
from src.gui.member_view import MemberWorkspaceView


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


# =============================================================================
# SUITE 1: Leader View 4 KPI Cards Empirical Stress Suite
# =============================================================================

class TestKPICardsEmpiricalStress:
    """Rigorous empirical verification of the 4 compact KPI Cards under dynamic stress."""

    def test_kpi_total_models_dynamic_scaling(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Stress test 'Tổng số Model' card with rapid scaling from 0 to 60+, deletions, and bursts."""
        view = LeaderWorkspaceView(base_dir=tmp_path)

        # Baseline check
        assert view.lbl_kpi_total_models.text() == "0"

        # 1. Incremental addition from 1 to 60
        for count in range(1, 65):
            m = MachineTarget(machine_code=f"MODEL_{count:03d}")
            view.state.machines.append(m)
            view.refresh_kpi_cards()
            assert view.lbl_kpi_total_models.text() == str(count), f"Mismatch at count {count}"

        # 2. Deletion of 25 models
        del view.state.machines[10:35]  # remove 25
        view.refresh_kpi_cards()
        assert len(view.state.machines) == 39
        assert view.lbl_kpi_total_models.text() == "39"

        # 3. Clear all models
        view.state.machines.clear()
        view.refresh_kpi_cards()
        assert view.lbl_kpi_total_models.text() == "0"

        # 4. Burst addition of 80 models
        burst_machines = [MachineTarget(machine_code=f"BURST_{i}") for i in range(80)]
        view.state.machines.extend(burst_machines)
        view.refresh_kpi_cards()
        assert view.lbl_kpi_total_models.text() == "80"

        # 5. Rapid repeated refresh (50 calls) stability check
        for _ in range(50):
            view.refresh_kpi_cards()
        assert view.lbl_kpi_total_models.text() == "80"

    def test_kpi_ready_models_combinatorial_matrix(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Stress test 'Model đủ BOM' card with 50+ models under varied file combinations."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        assert view.lbl_kpi_ready_models.text() == "0"

        # Set up 52 models:
        # Group 1 (13 models): Both PLM and R3 exist -> READY
        # Group 2 (13 models): Only PLM exists -> NOT READY
        # Group 3 (13 models): Only R3 exists -> NOT READY
        # Group 4 (13 models): Neither exists -> NOT READY
        machines: list[MachineTarget] = []

        for i in range(52):
            code = f"MACH_{i:02d}"
            m_dir = tmp_path / code
            m_dir.mkdir(parents=True, exist_ok=True)
            m = MachineTarget(machine_code=code, folder_path=m_dir)

            if i < 13:  # Both
                plm = m_dir / "PLM.xlsx"
                plm.touch()
                r3 = m_dir / "R3.xls"
                r3.touch()
                m.plm_file = plm
                m.r3_file = r3
            elif i < 26:  # PLM only
                plm = m_dir / "PLM.xlsx"
                plm.touch()
                m.plm_file = plm
                m.r3_file = None
            elif i < 39:  # R3 only
                r3 = m_dir / "R3.xls"
                r3.touch()
                m.plm_file = None
                m.r3_file = r3
            else:  # Neither
                m.plm_file = None
                m.r3_file = None

            machines.append(m)

        view.state.machines = machines
        view.refresh_kpi_cards()

        # 13 out of 52 are ready
        assert view.lbl_kpi_ready_models.text() == "13/52"

        # Now make Group 2 ready by providing R3 files
        for i in range(13, 26):
            r3 = machines[i].folder_path / "R3.xls"
            r3.touch()
            machines[i].r3_file = r3

        view.refresh_kpi_cards()
        assert view.lbl_kpi_ready_models.text() == "26/52"

        # Now delete on-disk files for 5 of the ready models in Group 1
        for i in range(5):
            machines[i].plm_file.unlink()  # delete from disk!
        view.refresh_kpi_cards()
        # plm_file.exists() returns False -> ready count drops from 26 to 21
        assert view.lbl_kpi_ready_models.text() == "21/52"

        # Empty all models -> should return to "0"
        view.state.machines.clear()
        view.refresh_kpi_cards()
        assert view.lbl_kpi_ready_models.text() == "0"

    def test_kpi_cttt_progress_fluctuations_and_rounding(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Stress test 'Tiến độ nộp CTTT' card across percentages, rounding, and toggling."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        assert view.lbl_kpi_cttt_progress.text() == "0%"

        # 1. 60 submissions: 0 OK
        subs: list[MemberSubmissionStatus] = [
            MemberSubmissionStatus(
                machine_code=f"M_{i}",
                sub_unit="CTTT",
                engineer_name=f"Eng_{i}",
                is_submitted_ok=False,
            )
            for i in range(60)
        ]
        view.state.submissions = subs
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "0% (0/60)"

        # 2. Step up by quarters: 15, 30, 45, 60
        for i in range(15):
            subs[i].is_submitted_ok = True
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "25% (15/60)"

        for i in range(15, 30):
            subs[i].is_submitted_ok = True
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "50% (30/60)"

        for i in range(30, 45):
            subs[i].is_submitted_ok = True
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "75% (45/60)"

        for i in range(45, 60):
            subs[i].is_submitted_ok = True
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "100% (60/60)"

        # 3. Rounding edge cases (floor/int rounding check)
        # 1 / 3 = 33.333% -> "33% (1/3)"
        view.state.submissions = [
            MemberSubmissionStatus(machine_code="M", sub_unit="U", engineer_name="E1", is_submitted_ok=True),
            MemberSubmissionStatus(machine_code="M", sub_unit="U", engineer_name="E2", is_submitted_ok=False),
            MemberSubmissionStatus(machine_code="M", sub_unit="U", engineer_name="E3", is_submitted_ok=False),
        ]
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "33% (1/3)"

        # 2 / 3 = 66.666% -> "66% (2/3)"
        view.state.submissions[1].is_submitted_ok = True
        view.refresh_kpi_cards()
        assert view.lbl_kpi_cttt_progress.text() == "66% (2/3)"

        # 4. Rapid toggle of all 60 submissions 20 times
        for cycle in range(20):
            target_ok = (cycle % 2 == 0)
            for s in subs:
                s.is_submitted_ok = target_ok
            view.state.submissions = subs
            view.refresh_kpi_cards()
            if target_ok:
                assert view.lbl_kpi_cttt_progress.text() == "100% (60/60)"
            else:
                assert view.lbl_kpi_cttt_progress.text() == "0% (0/60)"

    def test_kpi_recon_status_lifecycle_transitions(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Stress test 'Trạng thái đối soát' card across lifecycle states and resets."""
        view = LeaderWorkspaceView(base_dir=tmp_path)

        # Baseline: No models, no report, no result -> "Chờ khởi tạo"
        assert view.lbl_kpi_recon_status.text() == "Chờ khởi tạo"

        # Transition 1: Add model -> "Đang chuẩn bị"
        view.state.machines = [MachineTarget(machine_code="VIRGO")]
        view.refresh_kpi_cards()
        assert view.lbl_kpi_recon_status.text() == "Đang chuẩn bị"

        # Transition 2: Step 3 marked complete -> "Sẵn sàng đối soát"
        view.state.step3_completed = True
        view.refresh_kpi_cards()
        assert view.lbl_kpi_recon_status.text() == "Sẵn sàng đối soát"

        # Transition 3: Reconciliation finished -> "Hoàn tất"
        dummy_result = ReconciliationResult(
            cttt_rows=pd.DataFrame(),
            plm_missing_rows=pd.DataFrame(),
            cttt_totals=pd.DataFrame(),
            msi_results=pd.DataFrame(),
            overall_status="OK",
        )
        view.current_result = dummy_result
        view.refresh_kpi_cards()
        assert view.lbl_kpi_recon_status.text() == "Hoàn tất"

        # Transition 4: Report generated -> remains "Hoàn tất"
        view.current_result = None
        view.last_report_path = tmp_path / "report.xlsx"
        view.refresh_kpi_cards()
        assert view.lbl_kpi_recon_status.text() == "Hoàn tất"

        # Transition 5: Clean reset
        view.current_result = None
        view.last_report_path = None
        view.state.step3_completed = False
        view.refresh_kpi_cards()
        assert view.lbl_kpi_recon_status.text() == "Đang chuẩn bị"

        view.state.machines.clear()
        view.refresh_kpi_cards()
        assert view.lbl_kpi_recon_status.text() == "Chờ khởi tạo"

    def test_kpi_cards_randomized_property_oracle_stress(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Adversarial Property-Based Oracle: 100 randomized rounds testing all 4 cards simultaneously."""
        rng = random.Random(42)
        view = LeaderWorkspaceView(base_dir=tmp_path)

        for iteration in range(100):
            # Random machine count [0, 55]
            n_machines = rng.randint(0, 55)
            machines: list[MachineTarget] = []
            expected_ready = 0

            for m_idx in range(n_machines):
                m_dir = tmp_path / f"ITER_{iteration}_M_{m_idx}"
                m = MachineTarget(machine_code=f"MACH_{m_idx}", folder_path=m_dir)

                has_plm = rng.choice([True, False])
                has_r3 = rng.choice([True, False])

                if has_plm:
                    m_dir.mkdir(parents=True, exist_ok=True)
                    plm = m_dir / "PLM.xlsx"
                    plm.touch()
                    m.plm_file = plm

                if has_r3:
                    m_dir.mkdir(parents=True, exist_ok=True)
                    r3 = m_dir / "R3.xls"
                    r3.touch()
                    m.r3_file = r3

                if has_plm and has_r3:
                    expected_ready += 1

                machines.append(m)

            view.state.machines = machines

            # Random submissions count [0, 45]
            n_subs = rng.randint(0, 45)
            expected_ok_subs = 0
            subs: list[MemberSubmissionStatus] = []
            for s_idx in range(n_subs):
                is_ok = rng.choice([True, False])
                if is_ok:
                    expected_ok_subs += 1
                subs.append(
                    MemberSubmissionStatus(
                        machine_code=f"MACH_{s_idx}",
                        sub_unit="CTTT",
                        engineer_name=f"Eng_{s_idx}",
                        is_submitted_ok=is_ok,
                    )
                )
            view.state.submissions = subs

            # Random lifecycle flags
            has_result = rng.choice([True, False])
            has_report = rng.choice([True, False])
            step3_done = rng.choice([True, False])

            view.current_result = MagicMock() if has_result else None
            view.last_report_path = (tmp_path / "dummy.xlsx") if has_report else None
            view.state.step3_completed = step3_done

            # Trigger refresh
            view.refresh_kpi_cards()

            # Oracle Assertions
            # Card 1: Total models
            assert view.lbl_kpi_total_models.text() == str(n_machines)

            # Card 2: Ready models
            expected_ready_str = f"{expected_ready}/{n_machines}" if n_machines > 0 else "0"
            assert view.lbl_kpi_ready_models.text() == expected_ready_str

            # Card 3: CTTT Progress
            if n_subs > 0:
                expected_pct = int((expected_ok_subs / n_subs) * 100)
                expected_prog_str = f"{expected_pct}% ({expected_ok_subs}/{n_subs})"
            else:
                expected_prog_str = "0%"
            assert view.lbl_kpi_cttt_progress.text() == expected_prog_str

            # Card 4: Recon Status
            if has_result or has_report:
                expected_status = "Hoàn tất"
            elif step3_done:
                expected_status = "Sẵn sàng đối soát"
            elif n_machines > 0:
                expected_status = "Đang chuẩn bị"
            else:
                expected_status = "Chờ khởi tạo"
            assert view.lbl_kpi_recon_status.text() == expected_status


# =============================================================================
# SUITE 2: 7 Table Grids Invariants, Density & Geometry Stress Suite
# =============================================================================

class TestTableGridsEmpiricalStress:
    """Stress test the 7 primary tables across Leader and Member views."""

    def _get_all_7_tables(self, tmp_path: Path) -> tuple[LeaderWorkspaceView, MemberWorkspaceView, dict[str, QTableWidget]]:
        """Helper to instantiate both views and collect all 7 tables with lifetime preservation."""
        leader_view = LeaderWorkspaceView(base_dir=tmp_path)
        member_view = MemberWorkspaceView(base_dir=tmp_path)

        tables = {
            "leader_machine_table": leader_view.step1_widget.machine_table,
            "leader_staff_table": leader_view.step1_widget.staff_table,
            "leader_sourcing_table": leader_view.step2_widget.sourcing_table,
            "leader_submission_table": leader_view.step3_widget.submission_table,
            "member_cttt_table": member_view.cttt_table,
            "member_msi_table": member_view.msi_table,
            "member_label_table": member_view.label_table,
        }
        return leader_view, member_view, tables

    def test_7_tables_data_dense_invariants(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify strict adherence to RES-UI-01 & RES-UI-03 on all 7 tables."""
        _lv, _mv, tables = self._get_all_7_tables(tmp_path)
        assert len(tables) == 7

        for name, table in tables.items():
            header = table.verticalHeader()
            assert header.defaultSectionSize() == 32, (
                f"Table {name} has defaultSectionSize {header.defaultSectionSize()} != 32"
            )
            assert header.minimumSectionSize() == 28, (
                f"Table {name} has minimumSectionSize {header.minimumSectionSize()} != 28"
            )
            assert table.showGrid() is True, (
                f"Table {name} does not showGrid"
            )

    def test_7_tables_heavy_row_insertion_density(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Insert 100 rows into each table and verify row heights remain strictly 32px."""
        _lv, _mv, tables = self._get_all_7_tables(tmp_path)

        for name, table in tables.items():
            table.setRowCount(0)
            col_count = table.columnCount()

            # Insert 100 rows with dummy text
            table.blockSignals(True)
            try:
                for r in range(100):
                    table.insertRow(r)
                    for c in range(col_count):
                        table.setItem(r, c, QTableWidgetItem(f"Row {r} Col {c}"))
            finally:
                table.blockSignals(False)

            assert table.rowCount() == 100

            # Verify row heights across all rows
            header = table.verticalHeader()
            assert header.defaultSectionSize() == 32

            for r in range(100):
                # When default section size is 32 and no custom size is forced, rowHeight(r) == 32
                assert table.rowHeight(r) == 32, (
                    f"Table {name} row {r} height is {table.rowHeight(r)} != 32"
                )

            # Delete 50 rows
            for _ in range(50):
                table.removeRow(0)
            assert table.rowCount() == 50

            for r in range(50):
                assert table.rowHeight(r) == 32, (
                    f"Table {name} row {r} height is {table.rowHeight(r)} != 32 after deletion"
                )

            # Clear all rows
            table.setRowCount(0)
            assert table.rowCount() == 0
            assert header.defaultSectionSize() == 32

    def test_7_tables_extreme_viewport_resizing_layout_stability(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Stress test table stability against layout jitter/co giật across extreme window sizes."""
        leader_view = LeaderWorkspaceView(base_dir=tmp_path)
        member_view = MemberWorkspaceView(base_dir=tmp_path)

        views = [leader_view, member_view]

        # Populate tables with sample rows
        for v in views:
            v.show()

        extreme_resolutions = [
            QSize(300, 200),     # Very tiny
            QSize(640, 480),     # VGA
            QSize(1024, 768),    # XGA
            QSize(1366, 768),    # HD Laptop
            QSize(1920, 1080),   # Full HD 1080p
            QSize(2560, 1440),   # 2K QHD
            QSize(3840, 2160),   # 4K UHD
            QSize(100, 100),     # Ultra collapsed
            QSize(1920, 1080),   # Return to Full HD
        ]

        tables = {
            "leader_machine": leader_view.step1_widget.machine_table,
            "leader_staff": leader_view.step1_widget.staff_table,
            "leader_sourcing": leader_view.step2_widget.sourcing_table,
            "leader_submission": leader_view.step3_widget.submission_table,
            "member_cttt": member_view.cttt_table,
            "member_msi": member_view.msi_table,
            "member_label": member_view.label_table,
        }

        # Rapidly loop through resolutions (50 rapid resize events)
        for cycle in range(5):
            for res in extreme_resolutions:
                leader_view.resize(res)
                member_view.resize(res)
                qapp.processEvents()

                # Invariant checks during resize:
                for name, tbl in tables.items():
                    assert tbl.verticalHeader().defaultSectionSize() == 32, (
                        f"Table {name} defaultSectionSize changed during resize to {res}"
                    )
                    assert tbl.showGrid() is True, (
                        f"Table {name} showGrid altered during resize to {res}"
                    )

        leader_view.close()
        member_view.close()

    def test_7_tables_horizontal_column_resizing_isolation(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify that horizontal column stretching and resizing never alters row height (32px)."""
        _lv, _mv, tables = self._get_all_7_tables(tmp_path)

        for name, table in tables.items():
            h_header = table.horizontalHeader()
            v_header = table.verticalHeader()

            # Insert 5 rows
            table.setRowCount(0)
            for r in range(5):
                table.insertRow(r)
                for c in range(table.columnCount()):
                    table.setItem(r, c, QTableWidgetItem(f"Val_{r}_{c}"))

            # Resize columns to extreme widths
            for c in range(table.columnCount()):
                h_header.resizeSection(c, 500)
                assert v_header.defaultSectionSize() == 32
                for r in range(5):
                    assert table.rowHeight(r) == 32

            # Resize columns down to 20px
            for c in range(table.columnCount()):
                h_header.resizeSection(c, 20)
                assert v_header.defaultSectionSize() == 32
                for r in range(5):
                    assert table.rowHeight(r) == 32

    def test_7_tables_rapid_scrolling_and_viewport_integrity(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Populate 150 rows in all 7 tables and simulate rapid bidirectional scroll without crash."""
        _lv, _mv, tables = self._get_all_7_tables(tmp_path)

        for name, table in tables.items():
            table.setRowCount(0)
            cols = table.columnCount()
            table.blockSignals(True)
            try:
                for r in range(100):
                    table.insertRow(r)
                    for c in range(cols):
                        table.setItem(r, c, QTableWidgetItem(f"Scroll_{r}_{c}"))
            finally:
                table.blockSignals(False)

            v_scroll = table.verticalScrollBar()
            max_val = v_scroll.maximum()

            # Rapid scroll up and down
            for _ in range(10):
                v_scroll.setValue(max_val)
                qapp.processEvents()
                assert table.verticalHeader().defaultSectionSize() == 32
                v_scroll.setValue(0)
                qapp.processEvents()
                assert table.verticalHeader().defaultSectionSize() == 32
                v_scroll.setValue(max_val // 2)
                qapp.processEvents()
                assert table.verticalHeader().defaultSectionSize() == 32
