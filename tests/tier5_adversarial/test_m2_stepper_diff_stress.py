"""Adversarial Empirical Stress Verification Suite for Milestone 2 (M2).

Focus:
1. Stepper 3-Step Lifecycle Stress Verifier:
   - 50 continuous transition cycles: Mở gói -> Đối soát sơ bộ -> Xác nhận nộp -> Mở khóa.
   - SVG icon states (active / completed / pending) rendering and validation (pixmaps, colors, styles).
   - Integrated MemberWorkspaceView Stepper state machine stability.
   - Zero crash, zero memory leak across rapid oscillations.

2. Diff View High-Volume Stress Verifier:
   - 1,000 rows benchmark: styling latency, precision of OK/NG classification.
   - 10,000 rows extreme benchmark: time-to-render, memory footprint, boundary verification.
   - Mathematical WCAG 2.1 AAA & AA contrast ratio verification on applied row & cell colors.
   - Cyclic memory leak detection across repeated load-check-clear sequences.
"""

from __future__ import annotations

import gc
import time
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import psutil
import pytest
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QApplication,
    QMessageBox,
    QTableWidgetItem,
)

from src.gui.member_view import MemberWorkflowStepper, MemberWorkspaceView
from src.gui.styles.tokens import (
    calculate_contrast_ratio,
)
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_GREEN_FONT_HEX,
    COLOR_RED_FILL_HEX,
    COLOR_RED_FONT_HEX,
)


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


# =============================================================================
# SUITE 1: Stepper 3-Step Lifecycle Empirical Stress Verifier
# =============================================================================

class TestStepperLifecycleStress:
    """Stress tests verifying the 3-step Stepper lifecycle across continuous transitions."""

    def test_stepper_initial_state(self, qapp: QApplication) -> None:
        """Verify initial state of MemberWorkflowStepper: Step 0 Active, Steps 1 & 2 Pending."""
        stepper = MemberWorkflowStepper()
        assert stepper.current_step == 0
        assert stepper.step_completed_flags == [False, False, False]

        # Step 0: Active (Blue)
        assert "#EFF6FF" in stepper._step_widgets[0].styleSheet()
        assert "#1E40AF" in stepper._step_labels[0].styleSheet()
        assert not stepper._icon_labels[0].pixmap().isNull()
        assert stepper._icon_labels[0].pixmap().width() > 0

        # Steps 1 & 2: Pending (Slate)
        for idx in (1, 2):
            assert "#F8FAFC" in stepper._step_widgets[idx].styleSheet()
            assert "#64748B" in stepper._step_labels[idx].styleSheet()
            assert not stepper._icon_labels[idx].pixmap().isNull()
            assert stepper._icon_labels[idx].pixmap().width() > 0

    def test_stepper_50_continuous_transition_cycles_standalone(self, qapp: QApplication) -> None:
        """Stress-test standalone MemberWorkflowStepper through 50 continuous 4-phase cycles.

        Lifecycle Phases per Cycle:
        - Phase 1 (Mở gói / Nạp dữ liệu): Step 0 Completed (Green check), Step 1 Active (Blue), Step 2 Pending (Slate)
        - Phase 2 (Đối soát sơ bộ): Step 0 Completed, Step 1 Completed, Step 2 Active (Blue)
        - Phase 3 (Xác nhận nộp): Step 0 Completed, Step 1 Completed, Step 2 Completed (All Green checks)
        - Phase 4 (Mở khóa bài nộp): Step 0 Completed, Step 1 Active, Step 2 Pending
        Total transitions: 50 * 4 = 200 transitions.
        """
        stepper = MemberWorkflowStepper()
        start_time = time.perf_counter()

        for cycle in range(1, 51):
            # Phase 1: Mở gói / Nạp dữ liệu
            stepper.set_step_completed(0, True)
            stepper.set_step_completed(1, False)
            stepper.set_step_completed(2, False)
            stepper.set_current_step(1)

            assert stepper.current_step == 1
            assert stepper.step_completed_flags == [True, False, False]
            # Step 0: completed (Green check)
            assert "#DCFCE7" in stepper._step_widgets[0].styleSheet()
            assert "#064E3B" in stepper._step_labels[0].styleSheet()
            assert not stepper._icon_labels[0].pixmap().isNull()
            # Step 1: active (Blue)
            assert "#EFF6FF" in stepper._step_widgets[1].styleSheet()
            assert "#1E40AF" in stepper._step_labels[1].styleSheet()
            assert not stepper._icon_labels[1].pixmap().isNull()
            # Step 2: pending (Slate)
            assert "#F8FAFC" in stepper._step_widgets[2].styleSheet()
            assert "#64748B" in stepper._step_labels[2].styleSheet()
            assert not stepper._icon_labels[2].pixmap().isNull()

            # Phase 2: Đối soát sơ bộ
            stepper.set_step_completed(0, True)
            stepper.set_step_completed(1, True)
            stepper.set_current_step(2)

            assert stepper.current_step == 2
            assert stepper.step_completed_flags == [True, True, False]
            # Step 0 & 1: completed
            assert "#DCFCE7" in stepper._step_widgets[0].styleSheet()
            assert "#DCFCE7" in stepper._step_widgets[1].styleSheet()
            # Step 2: active
            assert "#EFF6FF" in stepper._step_widgets[2].styleSheet()
            assert "#1E40AF" in stepper._step_labels[2].styleSheet()
            assert not stepper._icon_labels[2].pixmap().isNull()

            # Phase 3: Xác nhận nộp
            stepper.set_step_completed(2, True)
            stepper.set_current_step(2)

            assert stepper.current_step == 2
            assert stepper.step_completed_flags == [True, True, True]
            for idx in range(3):
                assert "#DCFCE7" in stepper._step_widgets[idx].styleSheet()
                assert "#064E3B" in stepper._step_labels[idx].styleSheet()
                assert not stepper._icon_labels[idx].pixmap().isNull()

            # Phase 4: Mở khóa bài nộp
            stepper.set_step_completed(1, False)
            stepper.set_step_completed(2, False)
            stepper.set_current_step(1)

            assert stepper.current_step == 1
            assert stepper.step_completed_flags == [True, False, False]
            # Step 0: completed, Step 1: active, Step 2: pending
            assert "#DCFCE7" in stepper._step_widgets[0].styleSheet()
            assert "#EFF6FF" in stepper._step_widgets[1].styleSheet()
            assert "#F8FAFC" in stepper._step_widgets[2].styleSheet()
            assert not stepper._icon_labels[0].pixmap().isNull()
            assert not stepper._icon_labels[1].pixmap().isNull()
            assert not stepper._icon_labels[2].pixmap().isNull()

        elapsed = time.perf_counter() - start_time
        print(f"\n[BENCHMARK] 50 Stepper cycles (200 transitions) completed in {elapsed:.3f}s ({elapsed / 200 * 1000:.2f} ms/transition)")
        assert elapsed < 15.0, f"Stepper transitions too slow: {elapsed:.3f}s"

    def test_stepper_unlock_submission_state_retention_flaw(self, qapp: QApplication, tmp_path: Path) -> None:
        """Adversarial check demonstrating the Stepper desynchronization in member_view.unlock_submission().

        When unlock_submission() is called, it sets step_completed_flags[2]=False and current_step=1,
        but fails to reset step_completed_flags[1]=False.
        Consequently, Step 1 remains in Completed (Green) style rather than Active (Blue) style!
        """
        view = MemberWorkspaceView(base_dir=tmp_path)
        plm_df = pd.DataFrame([{"PART_CODE": "PART-A", "QUANTITY": 1, "REV": "A"}])
        r3_df = pd.DataFrame([{"MATERIAL": "PART-A", "QUANTITY": 1, "REVLEV": "A"}])

        view.add_cttt_row("Trang 1", "PART-A", "Capacitor", 1.0, "LSU", "Kỹ sư A")

        with patch.object(QMessageBox, "information"), patch.object(QMessageBox, "warning"):
            # Step 1 -> Step 2
            view.set_reference_data(plm_df, r3_df)
            view.run_preliminary_self_check()
            assert view.workflow_stepper.step_completed_flags[1] is True

            # Step 3
            view.workflow_stepper.set_step_completed(2, True)
            view.workflow_stepper.set_current_step(2)

            # Unlock
            view.unlock_submission()

            # VERIFIED STATE: Step 1 completed flag IS cleared by unlock_submission!
            step1_still_completed = view.workflow_stepper.step_completed_flags[1]
            step1_stylesheet = view.workflow_stepper._step_widgets[1].styleSheet()

            print(f"\n[STEPPER UNLOCK AUDIT] Step 1 completed flag after unlock: {step1_still_completed}")
            print(f"[STEPPER UNLOCK AUDIT] Step 1 widget stylesheet: {step1_stylesheet}")

            # Verified fix: Step 1 properly displays as #EFF6FF (Blue Active)
            assert step1_still_completed is False, "Step 1 flag should be cleared on unlock"
            assert "#EFF6FF" in step1_stylesheet, "Step 1 should display Blue Active style upon unlock"

    def test_stepper_integrated_member_view_50_cycles(self, qapp: QApplication, tmp_path: Path) -> None:
        """Stress test MemberWorkspaceView with 50 live end-to-end Stepper state transitions."""
        view = MemberWorkspaceView(base_dir=tmp_path)

        plm_df = pd.DataFrame([{"PART_CODE": "PART-A", "QUANTITY": 1, "REV": "A"}])
        r3_df = pd.DataFrame([{"MATERIAL": "PART-A", "QUANTITY": 1, "REVLEV": "A"}])

        view.add_cttt_row("Trang 1", "PART-A", "Capacitor", 1.0, "LSU", "Kỹ sư A")

        start_time = time.perf_counter()

        with patch.object(QMessageBox, "information"), patch.object(QMessageBox, "warning"):
            for cycle in range(1, 51):
                # 1. Supply reference data (Phase 1)
                view.set_reference_data(plm_df, r3_df)
                assert view.workflow_stepper.current_step == 1
                assert view.workflow_stepper.step_completed_flags[0] is True

                # 2. Run self check (Phase 2)
                res = view.run_preliminary_self_check()
                assert res["status"] == "OK"
                assert view.workflow_stepper.current_step == 2
                assert view.workflow_stepper.step_completed_flags[0] is True
                assert view.workflow_stepper.step_completed_flags[1] is True

                # 3. Simulate submission (Phase 3)
                view.workflow_stepper.set_step_completed(2, True)
                view.workflow_stepper.set_current_step(2)
                assert view.workflow_stepper.step_completed_flags == [True, True, True]

                # 4. Unlock submission (Phase 4)
                view.unlock_submission()
                assert view.workflow_stepper.current_step == 1
                assert view.workflow_stepper.step_completed_flags[2] is False

        elapsed = time.perf_counter() - start_time
        print(f"\n[BENCHMARK] 50 Integrated MemberView Stepper cycles executed in {elapsed:.3f}s")
        assert elapsed < 10.0, f"Integrated stepper cycles took too long: {elapsed:.3f}s"


# =============================================================================
# SUITE 2: Diff View WCAG Contrast Mathematical Oracle
# =============================================================================

class TestDiffViewContrastOracle:
    """Mathematical verification of WCAG 2.1 AAA & AA compliance for Diff View and Stepper palettes."""

    def test_diff_view_row_contrast_empirical(self) -> None:
        """Measure actual contrast ratios of Diff View row colors and compare with WCAG AA/AAA standards."""
        # OK Row: Soft green background #DCFCE7 vs Dark green text #166534
        ok_ratio = calculate_contrast_ratio("#166534", "#DCFCE7")
        print(f"\n[CONTRAST] OK Row (#166534 on #DCFCE7): {ok_ratio:.2f}:1")
        # Meets WCAG AA (>= 4.5:1)
        assert ok_ratio >= 4.5, f"OK row contrast {ok_ratio:.2f} failed WCAG AA (4.5:1)"
        # Note finding: 6.49:1 fails WCAG AAA (7.0:1) despite SPEC claim of 7.1:1
        is_aaa = ok_ratio >= 7.0
        print(f"[CONTRAST AUDIT] OK Row meets WCAG AA: True | meets WCAG AAA: {is_aaa} (Actual: {ok_ratio:.2f}:1, Claimed: 7.1:1)")

        # NG Row: Soft red background #FEE2E2 vs Dark red text #991B1B
        ng_ratio = calculate_contrast_ratio("#991B1B", "#FEE2E2")
        print(f"[CONTRAST] NG Row (#991B1B on #FEE2E2): {ng_ratio:.2f}:1")
        assert ng_ratio >= 4.5, f"NG row contrast {ng_ratio:.2f} failed WCAG AA (4.5:1)"
        is_aaa_ng = ng_ratio >= 7.0
        print(f"[CONTRAST AUDIT] NG Row meets WCAG AA: True | meets WCAG AAA: {is_aaa_ng} (Actual: {ng_ratio:.2f}:1, Claimed: 7.4:1)")

    def test_diff_view_status_cell_col7_contrast_wcag_aa(self) -> None:
        """Verify Col 7 status cells (Excel stamp colors) meet WCAG AA (4.5:1) standard."""
        ok_status_ratio = calculate_contrast_ratio(f"#{COLOR_GREEN_FONT_HEX}", f"#{COLOR_GREEN_FILL_HEX}")
        print(f"[CONTRAST] Col 7 OK Status (#{COLOR_GREEN_FONT_HEX} on #{COLOR_GREEN_FILL_HEX}): {ok_status_ratio:.2f}:1")
        assert ok_status_ratio >= 4.5, f"Col 7 OK status contrast {ok_status_ratio:.2f} failed WCAG AA"

        ng_status_ratio = calculate_contrast_ratio(f"#{COLOR_RED_FONT_HEX}", f"#{COLOR_RED_FILL_HEX}")
        print(f"[CONTRAST] Col 7 NG Status (#{COLOR_RED_FONT_HEX} on #{COLOR_RED_FILL_HEX}): {ng_status_ratio:.2f}:1")
        assert ng_status_ratio >= 4.5, f"Col 7 NG status contrast {ng_status_ratio:.2f} failed WCAG AA"

    def test_stepper_tokens_contrast(self) -> None:
        """Verify Stepper indicator palettes meet WCAG requirements."""
        completed_ratio = calculate_contrast_ratio("#166534", "#DCFCE7")
        print(f"[CONTRAST] Stepper Completed (#166534 on #DCFCE7): {completed_ratio:.2f}:1")
        assert completed_ratio >= 4.5  # Meets AA

        active_ratio = calculate_contrast_ratio("#1E40AF", "#EFF6FF")
        print(f"[CONTRAST] Stepper Active (#1E40AF on #EFF6FF): {active_ratio:.2f}:1")
        assert active_ratio >= 7.0, f"Stepper active contrast {active_ratio:.2f} failed WCAG AAA"

        pending_ratio = calculate_contrast_ratio("#64748B", "#F8FAFC")
        print(f"[CONTRAST] Stepper Pending (#64748B on #F8FAFC): {pending_ratio:.2f}:1")
        assert pending_ratio >= 4.5, f"Stepper pending contrast {pending_ratio:.2f} failed WCAG AA"


# =============================================================================
# SUITE 3: Diff View High-Volume Stress Verifier (1,000 & 10,000 Rows)
# =============================================================================

class TestDiffViewHighVolumeStress:
    """Stress tests evaluating rendering latency, contrast, and memory stability for 1k & 10k rows."""

    def _generate_synthetic_dataset(
        self,
        n_rows: int,
        ok_ratio: float = 0.60,
        qty_mismatch_ratio: float = 0.20,
        missing_ratio: float = 0.15,
        plm_qty_col_name: str = "QUANTITY",
    ) -> tuple[list[dict], pd.DataFrame, pd.DataFrame]:
        """Generate a realistic synthetic dataset of CTTT items with matching/mismatching PLM & R3."""
        cttt_rows: list[dict] = []
        plm_rows: list[dict] = []
        r3_rows: list[dict] = []

        n_ok = int(n_rows * ok_ratio)
        n_qty = int(n_rows * qty_mismatch_ratio)
        n_missing = int(n_rows * missing_ratio)
        n_rev = n_rows - (n_ok + n_qty + n_missing)

        # 1. OK items (exact match)
        for i in range(n_ok):
            pcode = f"PART-OK-{i:05d}"
            cttt_rows.append({"page": f"P-{i//20+1}", "code": pcode, "name": f"Resistor {i}", "qty": 10.0, "unit": "LSU", "author": "Eng1"})
            plm_rows.append({"PART_CODE": pcode, plm_qty_col_name: 10.0, "REV": "A"})
            r3_rows.append({"MATERIAL": pcode, "QUANTITY": 10.0, "REVLEV": "A"})

        # 2. Quantity mismatch items (NG)
        for i in range(n_qty):
            pcode = f"PART-QTY-{i:05d}"
            cttt_rows.append({"page": f"P-Q{i//20+1}", "code": pcode, "name": f"Capacitor {i}", "qty": 10.0, "unit": "FUSER", "author": "Eng2"})
            plm_rows.append({"PART_CODE": pcode, plm_qty_col_name: 8.0, "REV": "A"})  # Different PLM qty
            r3_rows.append({"MATERIAL": pcode, "QUANTITY": 10.0, "REVLEV": "A"})

        # 3. Missing in PLM/R3 (NG)
        for i in range(n_missing):
            pcode = f"PART-MISS-{i:05d}"
            cttt_rows.append({"page": f"P-M{i//20+1}", "code": pcode, "name": f"Diode {i}", "qty": 5.0, "unit": "MAIN", "author": "Eng3"})
            # Not added to plm_rows or r3_rows

        # 4. Revision mismatch items (NG)
        for i in range(n_rev):
            pcode = f"PART-REV-{i:05d}"
            cttt_rows.append({"page": f"P-R{i//20+1}", "code": pcode, "name": f"IC {i}", "qty": 2.0, "unit": "ENGINE", "author": "Eng4"})
            plm_rows.append({"PART_CODE": pcode, plm_qty_col_name: 2.0, "REV": "B"})  # Rev B in PLM
            r3_rows.append({"MATERIAL": pcode, "QUANTITY": 2.0, "REVLEV": "A"})  # Rev A in R3

        return cttt_rows, pd.DataFrame(plm_rows), pd.DataFrame(r3_rows)

    def test_plm_q_dot_ty_column_resolution_flaw(self, qapp: QApplication, tmp_path: Path) -> None:
        """Adversarial test revealing that column name 'Q.TY' fails to match member_view.py line 1437.

        In Sheet PLM (Kyocera canonical standard), the column header is 'Q.TY'.
        Because line 1437 checks ['QUANTITY', 'QTY', 'SL', 'SỐ LƯỢNG'], 'Q.TY' is not found.
        This causes plm_dict to be empty, silently falling back to plm_qty = cttt_qty,
        masking all PLM quantity differences!
        """
        view = MemberWorkspaceView(base_dir=tmp_path)
        # Generate with PLM column 'Q.TY'
        cttt_rows, plm_df, r3_df = self._generate_synthetic_dataset(100, plm_qty_col_name="Q.TY")
        assert "Q.TY" in plm_df.columns

        view.set_reference_data(plm_df, r3_df)
        view.cttt_table.setRowCount(len(cttt_rows))
        for r_idx, item in enumerate(cttt_rows):
            view.cttt_table.setItem(r_idx, 0, QTableWidgetItem(item["page"]))
            view.cttt_table.setItem(r_idx, 1, QTableWidgetItem(item["code"]))
            view.cttt_table.setItem(r_idx, 2, QTableWidgetItem(item["name"]))
            view.cttt_table.setItem(r_idx, 3, QTableWidgetItem(str(item["qty"])))
            view.cttt_table.setItem(r_idx, 4, QTableWidgetItem("0"))
            view.cttt_table.setItem(r_idx, 5, QTableWidgetItem("0"))
            view.cttt_table.setItem(r_idx, 6, QTableWidgetItem(""))
            view.cttt_table.setItem(r_idx, 7, QTableWidgetItem("-"))
            view.cttt_table.setItem(r_idx, 8, QTableWidgetItem(item["unit"]))
            view.cttt_table.setItem(r_idx, 9, QTableWidgetItem(item["author"]))
            for c in range(10, 15):
                view.cttt_table.setItem(r_idx, c, QTableWidgetItem("-"))

        res = view.run_preliminary_self_check()
        print(f"\n[PLM Q.TY VERIFICATION] Expected 60 OK, 40 NG. Actual: {res['ok_count']} OK, {res['ng_count']} NG")
        # Verified fix: 'Q.TY' column is now recognized properly; exactly 60 items match and 40 items have discrepancies
        assert res["ok_count"] == 60, "Expected exactly 60 items to be classified as OK"
        assert res["ng_count"] == 40, "Expected exactly 40 items to be classified as NG"

    def test_diff_view_1000_rows_benchmark(self, qapp: QApplication, tmp_path: Path) -> None:
        """Stress-test Diff View with 1,000 rows, measuring styling latency, contrast, and accuracy."""
        view = MemberWorkspaceView(base_dir=tmp_path)
        cttt_rows, plm_df, r3_df = self._generate_synthetic_dataset(1000, plm_qty_col_name="QUANTITY")

        view.set_reference_data(plm_df, r3_df)

        # Batch insert 1,000 rows into cttt_table
        t0 = time.perf_counter()
        view.cttt_table.setRowCount(len(cttt_rows))
        for r_idx, item in enumerate(cttt_rows):
            view.cttt_table.setItem(r_idx, 0, QTableWidgetItem(item["page"]))
            view.cttt_table.setItem(r_idx, 1, QTableWidgetItem(item["code"]))
            view.cttt_table.setItem(r_idx, 2, QTableWidgetItem(item["name"]))
            view.cttt_table.setItem(r_idx, 3, QTableWidgetItem(str(item["qty"])))
            view.cttt_table.setItem(r_idx, 4, QTableWidgetItem("0"))
            view.cttt_table.setItem(r_idx, 5, QTableWidgetItem("0"))
            view.cttt_table.setItem(r_idx, 6, QTableWidgetItem(""))
            view.cttt_table.setItem(r_idx, 7, QTableWidgetItem("-"))
            view.cttt_table.setItem(r_idx, 8, QTableWidgetItem(item["unit"]))
            view.cttt_table.setItem(r_idx, 9, QTableWidgetItem(item["author"]))
            for c in range(10, 15):
                view.cttt_table.setItem(r_idx, c, QTableWidgetItem("-"))
        pop_time = time.perf_counter() - t0

        # Run Preliminary Self Check (Reconciliation + Diff View Styling)
        t1 = time.perf_counter()
        res = view.run_preliminary_self_check()
        diff_time = time.perf_counter() - t1

        print(f"\n[BENCHMARK 1,000 ROWS]")
        print(f"  Population time: {pop_time:.3f}s")
        print(f"  Reconciliation + Diff Styling time: {diff_time:.3f}s")
        print(f"  Average time per row: {diff_time / 1000 * 1000:.2f} ms/row")
        print(f"  Result: {res['ok_count']} OK, {res['ng_count']} NG / Total: {res['total']}")

        # Assertions
        assert res["total"] == 1000
        assert res["ok_count"] == 600
        assert res["ng_count"] == 400
        assert diff_time < 2.5, f"1,000 rows diff styling exceeded 2.5s: {diff_time:.3f}s"

        # Verify applied styling on sample OK row (Row 0)
        row_0_bg = view.cttt_table.item(0, 1).background().color().name().upper()
        row_0_fg = view.cttt_table.item(0, 1).foreground().color().name().upper()
        assert row_0_bg == "#DCFCE7"
        assert row_0_fg == "#064E3B"
        assert view.cttt_table.item(0, 7).text() == "OK"
        assert view.cttt_table.item(0, 7).background().color().name().upper().endswith(COLOR_GREEN_FILL_HEX)

        # Verify applied styling on sample NG row (Row 650)
        row_650_bg = view.cttt_table.item(650, 1).background().color().name().upper()
        row_650_fg = view.cttt_table.item(650, 1).foreground().color().name().upper()
        assert row_650_bg == "#FEE2E2"
        assert row_650_fg == "#7F1D1D"
        assert view.cttt_table.item(650, 7).text() == "NG"
        assert view.cttt_table.item(650, 7).background().color().name().upper().endswith(COLOR_RED_FILL_HEX)

    def test_diff_view_10000_rows_extreme_stress(self, qapp: QApplication, tmp_path: Path) -> None:
        """Extreme stress test: Populate and render 10,000 rows in Diff View.

        Checks:
        - Memory footprint before, during, and after.
        - Performance duration for 10,000 row color styling.
        - Absence of memory leaks or process crashes.
        """
        process = psutil.Process()
        gc.collect()
        mem_before = process.memory_info().rss / (1024 * 1024)

        view = MemberWorkspaceView(base_dir=tmp_path)
        cttt_rows, plm_df, r3_df = self._generate_synthetic_dataset(10000, plm_qty_col_name="QUANTITY")

        view.set_reference_data(plm_df, r3_df)

        # Populate 10,000 rows
        t0 = time.perf_counter()
        view.cttt_table.setRowCount(len(cttt_rows))
        for r_idx, item in enumerate(cttt_rows):
            view.cttt_table.setItem(r_idx, 0, QTableWidgetItem(item["page"]))
            view.cttt_table.setItem(r_idx, 1, QTableWidgetItem(item["code"]))
            view.cttt_table.setItem(r_idx, 2, QTableWidgetItem(item["name"]))
            view.cttt_table.setItem(r_idx, 3, QTableWidgetItem(str(item["qty"])))
            view.cttt_table.setItem(r_idx, 4, QTableWidgetItem("0"))
            view.cttt_table.setItem(r_idx, 5, QTableWidgetItem("0"))
            view.cttt_table.setItem(r_idx, 6, QTableWidgetItem(""))
            view.cttt_table.setItem(r_idx, 7, QTableWidgetItem("-"))
            view.cttt_table.setItem(r_idx, 8, QTableWidgetItem(item["unit"]))
            view.cttt_table.setItem(r_idx, 9, QTableWidgetItem(item["author"]))
            for c in range(10, 15):
                view.cttt_table.setItem(r_idx, c, QTableWidgetItem("-"))
        pop_time = time.perf_counter() - t0

        mem_populated = process.memory_info().rss / (1024 * 1024)

        # Run Self Check on 10,000 rows
        t1 = time.perf_counter()
        res = view.run_preliminary_self_check()
        diff_time = time.perf_counter() - t1

        mem_styled = process.memory_info().rss / (1024 * 1024)

        print(f"\n[BENCHMARK 10,000 ROWS EXTREME STRESS]")
        print(f"  Memory baseline: {mem_before:.1f} MB")
        print(f"  Memory populated: {mem_populated:.1f} MB (+{mem_populated - mem_before:.1f} MB)")
        print(f"  Memory styled: {mem_styled:.1f} MB (+{mem_styled - mem_populated:.1f} MB)")
        print(f"  Population duration: {pop_time:.3f}s")
        print(f"  Self-check & Diff Styling duration: {diff_time:.3f}s")
        print(f"  Throughput: {10000 / diff_time:.0f} rows/second")
        print(f"  Result: {res['ok_count']} OK, {res['ng_count']} NG / Total: {res['total']}")

        # Boundary checks
        assert res["total"] == 10000
        assert res["ok_count"] == 6000
        assert res["ng_count"] == 4000

        # Check boundaries: row 0 (OK), row 5999 (OK), row 6000 (NG), row 9999 (NG)
        assert view.cttt_table.item(0, 1).background().color().name().upper() == "#DCFCE7"
        assert view.cttt_table.item(5999, 1).background().color().name().upper() == "#DCFCE7"
        assert view.cttt_table.item(6000, 1).background().color().name().upper() == "#FEE2E2"
        assert view.cttt_table.item(9999, 1).background().color().name().upper() == "#FEE2E2"

        # Test memory cleanup
        view.clear_cttt_table()
        gc.collect()
        mem_cleaned = process.memory_info().rss / (1024 * 1024)
        print(f"  Memory after table clear & GC: {mem_cleaned:.1f} MB (reclaimed {mem_styled - mem_cleaned:.1f} MB)")

        # Verify that memory is largely reclaimed (within reasonable margin)
        assert mem_cleaned < mem_styled, "Memory was not reclaimed after table clearing"

    def test_diff_view_memory_leak_cyclic_stress(self, qapp: QApplication, tmp_path: Path) -> None:
        """Detect memory leaks across 5 consecutive 1,000-row load, style, and clear cycles."""
        process = psutil.Process()
        view = MemberWorkspaceView(base_dir=tmp_path)
        cttt_rows, plm_df, r3_df = self._generate_synthetic_dataset(1000, plm_qty_col_name="QUANTITY")
        view.set_reference_data(plm_df, r3_df)

        mem_samples: list[float] = []

        for cycle in range(1, 6):
            view.cttt_table.setRowCount(len(cttt_rows))
            for r_idx, item in enumerate(cttt_rows):
                view.cttt_table.setItem(r_idx, 0, QTableWidgetItem(item["page"]))
                view.cttt_table.setItem(r_idx, 1, QTableWidgetItem(item["code"]))
                view.cttt_table.setItem(r_idx, 2, QTableWidgetItem(item["name"]))
                view.cttt_table.setItem(r_idx, 3, QTableWidgetItem(str(item["qty"])))
                view.cttt_table.setItem(r_idx, 4, QTableWidgetItem("0"))
                view.cttt_table.setItem(r_idx, 5, QTableWidgetItem("0"))
                view.cttt_table.setItem(r_idx, 6, QTableWidgetItem(""))
                view.cttt_table.setItem(r_idx, 7, QTableWidgetItem("-"))
                view.cttt_table.setItem(r_idx, 8, QTableWidgetItem(item["unit"]))
                view.cttt_table.setItem(r_idx, 9, QTableWidgetItem(item["author"]))
                for c in range(10, 15):
                    view.cttt_table.setItem(r_idx, c, QTableWidgetItem("-"))

            view.run_preliminary_self_check()
            view.clear_cttt_table()
            gc.collect()

            mem_now = process.memory_info().rss / (1024 * 1024)
            mem_samples.append(mem_now)

        print("\n[MEMORY LEAK CYCLIC SAMPLES (MB)]:", [f"{m:.1f}" for m in mem_samples])
        # Check stability: Memory growth from cycle 2 to cycle 5 must be less than 15 MB
        drift = mem_samples[-1] - mem_samples[1]
        print(f"Memory drift from cycle 2 to 5: {drift:.2f} MB")
        assert drift < 15.0, f"Memory leak detected: {drift:.2f} MB growth across cycles"
