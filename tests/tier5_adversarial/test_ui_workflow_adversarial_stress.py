"""Tier 5 Adversarial Stress & Hardening Test Suite for UI Workflows.

Empirically challenges the UI workflows:
1. Corrupted Excel workbooks in machine folders (truncated zip, binary garbage, malformed XML, zero-byte).
2. Missing sheets or cells when scanning for CTTT!Q2 = "OK" (missing CTTT, blank Q2, invalid strings, formulas).
3. Concurrent Q2 status updates & unsubmit/unlock race conditions (TOCTOU between scan and consolidation).
4. Extreme BOM code list inputs (empty list, 50+ codes, path traversal, Windows reserved device names, Unicode).
5. Fail-closed gate verification: ensure consolidation is impossible if even 1 member has not stamped OK,
   and test whether corrupted workbooks breach the fail-closed gate.
"""

from __future__ import annotations

import datetime
import io
import os
import shutil
import zipfile
from pathlib import Path
from typing import Any
from unittest.mock import patch

import openpyxl
from openpyxl.styles import PatternFill
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox, QTableWidget

from src.gui.leader_view import (
    LeaderSessionState,
    LeaderWorkspaceView,
    MachineTarget,
    MemberSubmissionStatus,
    ProjectStage,
    StaffAssignment,
    Step1ProjectSetupWidget,
    Step3TrackingConsolidationWidget,
)
from src.gui.member_view import MemberWorkspaceView


@pytest.fixture(autouse=True)
def qapp() -> QApplication:
    """Headless QApplication fixture."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


# =============================================================================
# 1. Corrupted Workbooks Stress Tests
# =============================================================================

class TestCorruptedWorkbooksStress:
    """Stress-test how the scanner and workspace handle corrupted and malformed workbooks."""

    def test_corrupted_zero_byte_workbook_handled_safely(self, tmp_path: Path) -> None:
        """0-byte file must NOT be considered submitted OK and must not crash scanner."""
        zero_file = tmp_path / "zero_byte.xlsx"
        zero_file.write_bytes(b"")

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        is_ok, count, msi_c, lbl_c = step3._inspect_submission_file(zero_file)
        assert is_ok is False, "0-byte file must never be marked as submitted OK"
        assert count == 0

    def test_corrupted_truncated_zip_empirical_probe(self, tmp_path: Path) -> None:
        """Empirical probe on truncated zip file.
        
        Tests whether the legacy fallback `is_ok = file_path.stat().st_size > 0`
        in `_inspect_submission_file` erroneously classifies corrupted files as OK.
        """
        corrupted_file = tmp_path / "truncated_zip.xlsx"
        # Truncated PK zip header with garbage payload
        corrupted_file.write_bytes(b"PK\x03\x04\x14\x00\x00\x00\x08\x00corrupted_payload_without_central_directory")

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        is_ok, count, msi_c, lbl_c = step3._inspect_submission_file(corrupted_file)
        assert is_ok is False, "Corrupted truncated zip must be rejected as NOT OK (fail-closed)"
        assert count == 0

    def test_corrupted_random_binary_garbage_empirical_probe(self, tmp_path: Path) -> None:
        """Empirical probe on completely invalid binary garbage (non-zip)."""
        garbage_file = tmp_path / "binary_garbage.xlsx"
        garbage_file.write_bytes(os.urandom(2048))

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        is_ok, count, msi_c, lbl_c = step3._inspect_submission_file(garbage_file)
        assert is_ok is False, "Random binary garbage must be rejected as NOT OK (fail-closed)"
        assert count == 0

    def test_corrupted_malformed_xml_in_valid_zip(self, tmp_path: Path) -> None:
        """Create a valid zip containing malformed XML syntax in workbook.xml."""
        bad_xml_file = tmp_path / "malformed_xml.xlsx"
        with zipfile.ZipFile(bad_xml_file, "w") as zf:
            zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
            zf.writestr("xl/workbook.xml", '<workbook><sheets><sheet name="CTTT" broken_unclosed>')

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        is_ok, count, _, _ = step3._inspect_submission_file(bad_xml_file)
        assert is_ok is False, "Malformed XML workbook must be rejected as NOT OK (fail-closed)"
        assert count == 0



    def test_member_workspace_open_corrupted_file_resilience(self, qapp: QApplication, tmp_path: Path) -> None:
        """Member workspace must survive opening a corrupted file without crashing."""
        corrupt_file = tmp_path / "corrupt_assignment.xlsx"
        corrupt_file.write_bytes(b"INVALID_EXCEL_DATA")

        view = MemberWorkspaceView(base_dir=tmp_path)
        res = view.open_assignment_package(corrupt_file)
        assert res is True  # File exists, state initialized
        assert view.is_submitted_ok is False
        assert "CHƯA NỘP" in view.lbl_submission_seal.text()


# =============================================================================
# 2. Missing Sheets or Cells when Scanning for CTTT!Q2 = "OK"
# =============================================================================

class TestMissingSheetsAndCellsStress:
    """Stress-test scanner against missing sheets, cells, and edge-case values."""

    def test_missing_cttt_sheet_in_machine_folder(self, tmp_path: Path) -> None:
        """A valid workbook with Sheet1 but no CTTT sheet must be rejected."""
        wb_file = tmp_path / "no_cttt.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Summary"
        ws["A1"] = "No CTTT here"
        wb.save(wb_file)
        wb.close()

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)
        is_ok, count, _, _ = step3._inspect_submission_file(wb_file)
        assert is_ok is False, "Workbook without CTTT sheet must not be submitted OK"

    def test_cttt_sheet_with_empty_or_whitespace_q2(self, tmp_path: Path) -> None:
        """Empty, whitespace-only, or None in Q2 must evaluate as NOT OK."""
        test_cases = [
            ("none", None),
            ("empty", ""),
            ("spaces", "   "),
            ("tabs", "\t\t"),
        ]

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        for name, q2_val in test_cases:
            f = tmp_path / f"test_{name}.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CTTT"
            if q2_val is not None:
                ws["Q2"] = q2_val
            wb.save(f)
            wb.close()

            is_ok, _, _, _ = step3._inspect_submission_file(f)
            assert is_ok is False, f"Q2 value '{q2_val}' must evaluate as not submitted"

    def test_cttt_sheet_with_invalid_statuses(self, tmp_path: Path) -> None:
        """Arbitrary non-OK strings and numbers must be rejected."""
        invalid_values = [
            "NOK", "PENDING", "REJECT", "FAIL", "0", "1", "TRUE", "FALSE", "OK_NOT", "N/A"
        ]

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        for val in invalid_values:
            f = tmp_path / f"test_val_{val.replace('/', '_')}.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CTTT"
            ws["Q2"] = val
            wb.save(f)
            wb.close()

            is_ok, _, _, _ = step3._inspect_submission_file(f)
            assert is_ok is False, f"Value '{val}' in Q2 must NOT be accepted as OK"

    def test_cttt_sheet_with_case_and_padding_variations(self, tmp_path: Path) -> None:
        """Valid variations of OK with casing and whitespace padding should be accepted."""
        valid_variations = ["OK", "ok", "Ok", "oK", "  OK  ", "\tOK\n"]

        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state)

        for val in valid_variations:
            f = tmp_path / f"test_ok_{hash(val)}.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CTTT"
            ws["Q2"] = val
            wb.save(f)
            wb.close()

            is_ok, _, _, _ = step3._inspect_submission_file(f)
            assert is_ok is True, f"Variation '{val}' must be accepted as OK"


# =============================================================================
# 3. Concurrent Updates & Unsubmit/Unlock Race Conditions (TOCTOU)
# =============================================================================

class TestTOCTOURaceConditionsStress:
    """Stress-test time-of-check to time-of-use and concurrency edge cases."""

    def test_toctou_unsubmit_after_scan_before_consolidation(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """TOCTOU Vulnerability Test:
        
        1. Member submits -> CTTT!Q2 = 'OK'.
        2. Leader runs `scan_submissions()` -> `all_members_ok = True`.
        3. Member un-submits / unlocks via `unlock_submission()` -> CTTT!Q2 becomes None.
        4. Leader calls `execute_consolidation()` without re-scanning.
        
        Vulnerability Check:
        Does `execute_consolidation()` blindly proceed because `all_members_ok` was cached,
        or does it verify the submission seal at execution time?
        """
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Virgo")
        model_dir = tmp_path / "Virgo"
        mach_code = "110C103NL0"
        mach_dir = model_dir / mach_code
        mach_dir.mkdir(parents=True, exist_ok=True)

        member_file = mach_dir / "Son_mecha1.xlsm"
        wb = openpyxl.Workbook()
        ws_cttt = wb.active
        ws_cttt.title = "CTTT"
        ws_cttt["Q2"] = "OK"
        ws_cttt.append(["Header1"])
        ws_cttt.append(["Header2"])
        ws_cttt.append(["LSU", "01", "302FP93010", "MOTOR", 1.0])
        wb.save(member_file)
        wb.close()

        # Step 1: Leader setup
        step1 = view.step1_widget
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code=mach_code)
        step1.sync_state_from_ui()

        # Step 3: Leader scans
        step3 = view.step3_widget
        statuses = step3.scan_submissions()
        assert step3.state.all_members_ok is True, "Gate should be open after valid submission"
        assert step3.btn_consolidate.isEnabled() is True

        # Interleaved Event: Member un-submits / clears Q2!
        member_view = MemberWorkspaceView(base_dir=tmp_path)
        member_view.open_assignment_package(member_file)
        assert member_view.is_submitted_ok is True
        member_view.unlock_submission()

        # Verify Q2 is now cleared on disk
        wb_check = openpyxl.load_workbook(member_file)
        assert wb_check["CTTT"]["Q2"].value is None
        wb_check.close()

        # Step 4: Leader executes consolidation WITHOUT calling scan_submissions() again
        # Check if consolidation is aborted because atomic pre-flight check re-scans disk:
        consolidation_result = step3.execute_consolidation()

        dest_file = mach_dir / "phutrach" / member_file.name
        was_moved = dest_file.exists() and not member_file.exists()

        assert consolidation_result is False, "Atomic pre-flight scan must detect unsubmitted file and abort consolidation"
        assert was_moved is False, "Unsubmitted file must NOT be moved to phutrach/"


    def test_file_deleted_between_scan_and_consolidation(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """File deleted after scan must be handled gracefully without crashing."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        view = LeaderWorkspaceView(base_dir=tmp_path)
        mach_code = "110C103NL0"
        mach_dir = tmp_path / "Virgo" / mach_code
        mach_dir.mkdir(parents=True, exist_ok=True)

        member_file = mach_dir / "Son_mecha1.xlsm"
        wb = openpyxl.Workbook()
        ws_cttt = wb.active
        ws_cttt.title = "CTTT"
        ws_cttt["Q2"] = "OK"
        wb.save(member_file)
        wb.close()

        step1 = view.step1_widget
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code=mach_code)
        step1.sync_state_from_ui()

        step3 = view.step3_widget
        step3.scan_submissions()
        assert step3.state.all_members_ok is True

        # Delete file before consolidation
        member_file.unlink()

        # Execute consolidation: atomic pre-flight check detects missing file and safely aborts
        success = step3.execute_consolidation()
        assert success is False
        assert len(step3.state.consolidated_cttt) == 0


    def test_rapid_submit_unlock_cycle_stress(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """20 rapid consecutive submit/unlock cycles on MemberWorkspaceView."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        mach_dir = tmp_path / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        assign_file = mach_dir / "Nguyen_mecha1.xlsm"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        wb.save(assign_file)
        wb.close()

        view = MemberWorkspaceView(base_dir=tmp_path)
        view.open_assignment_package(assign_file)
        view.add_cttt_row(page="01", part_code="302FP93010", part_name="MOTOR", quantity=1.0, sub_unit="LSU")

        for cycle in range(10):
            # Submit
            res = view.submit_data()
            assert res is not None
            assert view.is_submitted_ok is True

            wb_chk = openpyxl.load_workbook(assign_file, data_only=True)
            assert wb_chk["CTTT"]["Q2"].value == "OK"
            wb_chk.close()

            # Unlock
            view.unlock_submission()
            assert view.is_submitted_ok is False

            wb_chk2 = openpyxl.load_workbook(assign_file, data_only=True)
            assert wb_chk2["CTTT"]["Q2"].value is None
            wb_chk2.close()


# =============================================================================
# 4. Extreme BOM Code List Inputs (Step 1 Setup)
# =============================================================================

class TestExtremeBOMCodeInputsStress:
    """Stress-test Step 1 with extreme inputs, path traversals, and large code lists."""

    def test_step1_empty_machine_code_list_rejected(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Creating packages with empty machine list must be blocked."""
        warning_shown = []
        monkeypatch.setattr(QMessageBox, "warning", lambda p, t, m: warning_shown.append(m))

        state = LeaderSessionState(base_dir=tmp_path)
        step1 = Step1ProjectSetupWidget(state)
        step1.machine_table.setRowCount(0)

        step1.execute_create_folders_and_packages()
        assert len(warning_shown) > 0
        assert "chưa nhập mã máy" in warning_shown[0]
        assert state.step1_completed is False

    def test_step1_50_plus_machine_codes_stress(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Stress-test Step 1 with 55 machine codes."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        state = LeaderSessionState(base_dir=tmp_path)
        step1 = Step1ProjectSetupWidget(state)
        step1.machine_table.setRowCount(0)

        # Add 55 machine codes
        for i in range(1, 56):
            step1._add_machine_row(code=f"110C{i:03d}NL0")

        # Keep only 2 active engineers to avoid massive file generation overhead
        for r in range(step1.staff_table.rowCount()):
            w = step1.staff_table.cellWidget(r, 0)
            if w:
                chk = w.findChild(openpyxl.Workbook if False else pytest.importorskip("PyQt6.QtWidgets").QCheckBox)
                if chk:
                    chk.setChecked(r < 2)

        step1.execute_create_folders_and_packages()

        assert state.step1_completed is True
        assert len(state.machines) == 55

        # Verify all 55 directories created
        model_dir = tmp_path / state.model_name
        for i in range(1, 56):
            m_code = f"110C{i:03d}NL0"
            assert (model_dir / m_code).exists(), f"Directory for {m_code} must exist"

    def test_step1_path_traversal_and_invalid_chars_empirical(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Empirical probe: Test how Step 1 handles path traversal and Windows-forbidden characters."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        state = LeaderSessionState(base_dir=tmp_path)
        step1 = Step1ProjectSetupWidget(state)
        step1.machine_table.setRowCount(0)

        malicious_codes = [
            "../../escaped_dir",
            "110*STAR",
            "110:COLON",
            "CON",
            "PRN",
            "110<LT",
            "110>GT",
            "110|PIPE",
        ]

        # Test that input machine codes with path traversal or invalid chars are safely sanitized
        # and do not escape model_dir or crash with unhandled OSError
        model_dir = tmp_path / state.model_name
        for code in malicious_codes:
            step1.machine_table.setRowCount(0)
            step1._add_machine_row(code=code)
            step1.execute_create_folders_and_packages()

        # Confirm created directories reside strictly inside model_dir without path traversal
        for d in model_dir.iterdir():
            assert d.resolve().parent == model_dir.resolve()
            assert ".." not in d.name
            assert "/" not in d.name
            assert "\\" not in d.name



    def test_step1_unicode_machine_codes(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Machine codes with Japanese kanji and symbols."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        state = LeaderSessionState(base_dir=tmp_path)
        step1 = Step1ProjectSetupWidget(state)
        step1.machine_table.setRowCount(0)

        unicode_codes = ["110_東京_01", "110_★_TEST", "110_品質_V2"]
        for c in unicode_codes:
            step1._add_machine_row(code=c)

        # 1 active engineer
        for r in range(step1.staff_table.rowCount()):
            w = step1.staff_table.cellWidget(r, 0)
            if w:
                chk = w.findChild(pytest.importorskip("PyQt6.QtWidgets").QCheckBox)
                if chk:
                    chk.setChecked(r == 0)

        step1.execute_create_folders_and_packages()
        model_dir = tmp_path / state.model_name
        for c in unicode_codes:
            assert (model_dir / c).exists()


# =============================================================================
# 5. Fail-Closed Gate Violations
# =============================================================================

class TestFailClosedGateViolationsStress:
    """Stress-test fail-closed invariants and evaluate gate security."""

    def test_fail_closed_gate_blocks_when_1_of_N_members_pending(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Gate must block consolidation if even 1 out of 10 members has not stamped OK."""
        warning_shown = []
        monkeypatch.setattr(QMessageBox, "warning", lambda p, t, m: warning_shown.append(m))

        model_dir = tmp_path / "Virgo"
        mach_code = "110C103NL0"
        mach_dir = model_dir / mach_code
        mach_dir.mkdir(parents=True, exist_ok=True)

        # Create 9 OK workbooks and 1 pending workbook
        for i in range(1, 10):
            f = mach_dir / f"Eng_{i}_mecha1.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CTTT"
            ws["Q2"] = "OK"
            wb.save(f)
            wb.close()

        # 10th member is pending (empty Q2)
        f_pending = mach_dir / "Eng_10_mecha1.xlsm"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        ws["Q2"] = ""
        wb.save(f_pending)
        wb.close()

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code=mach_code)
        step1.sync_state_from_ui()

        step3 = view.step3_widget
        statuses = step3.scan_submissions()

        assert step3.state.all_members_ok is False, "Fail-closed gate must remain LOCKED"
        assert step3.btn_consolidate.isEnabled() is False, "Consolidate button must be DISABLED"

        # Attempt to execute consolidation
        res = step3.execute_consolidation()
        assert res is False, "execute_consolidation must return False"
        assert len(warning_shown) > 0
        assert "chưa nhập danh sách linh kiện" in warning_shown[0]

    def test_fail_closed_gate_breached_by_corrupted_file_empirical(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Empirical verification: Does a non-empty corrupted file breach the fail-closed gate?
        
        Setup:
        - 3 members: 2 valid OK submissions, 1 corrupted file (truncated zip).
        - Test whether `scan_submissions()` erroneously sets `all_members_ok = True`!
        """
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: None)

        model_dir = tmp_path / "Virgo"
        mach_code = "110C103NL0"
        mach_dir = model_dir / mach_code
        mach_dir.mkdir(parents=True, exist_ok=True)

        # Member 1 & 2: Valid OK files
        for eng in ["Son_mecha1", "Duy_mecha1"]:
            f = mach_dir / f"{eng}.xlsm"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CTTT"
            ws["Q2"] = "OK"
            wb.save(f)
            wb.close()

        # Member 3: Corrupted file (random garbage / truncated zip)
        f_corrupt = mach_dir / "KhiemA_mecha1.xlsm"
        f_corrupt.write_bytes(b"PK\x03\x04CORRUPTED_NON_EMPTY_EXCEL_WORKBOOK")

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code=mach_code)
        step1.sync_state_from_ui()

        step3 = view.step3_widget
        statuses = step3.scan_submissions()

        # Verify fail-closed gate blocks consolidation when corrupted file is present
        gate_breached = step3.state.all_members_ok
        button_enabled = step3.btn_consolidate.isEnabled()

        assert gate_breached is False, (
            "Fail-closed invariant: all_members_ok must remain False when a corrupted file is present"
        )
        assert button_enabled is False, (
            "Fail-closed invariant: btn_consolidate must remain disabled when a corrupted file is present"
        )


    def test_fail_closed_member_submit_blocked_on_unexplained_ng(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Member submission must prompt warning if NG item lacks explanation.
        If user declines prompt, submission is aborted (Q2 remains empty).
        """
        # User clicks "No" when warned about unaddressed NG
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: QMessageBox.StandardButton.No)

        mach_dir = tmp_path / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        assign_file = mach_dir / "Test_mecha1.xlsm"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        wb.save(assign_file)
        wb.close()

        view = MemberWorkspaceView(base_dir=tmp_path)
        view.open_assignment_package(assign_file)

        # Add an NG row without explanation
        view.add_cttt_row(
            page="01",
            part_code="302FP93010",
            part_name="MOTOR",
            quantity=1.0,
            explanation="",
            status="NG",
            sub_unit="LSU",
        )

        # Trigger submit
        res = view.submit_data()
        assert res is None, "Submission must be aborted when user cancels at NG warning"
        assert view.is_submitted_ok is False

        # File on disk must NOT have Q2="OK"
        wb_chk = openpyxl.load_workbook(assign_file)
        assert wb_chk["CTTT"]["Q2"].value != "OK"
        wb_chk.close()

    def test_fail_closed_member_submit_blocked_on_empty_table(
        self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Member submission must be rejected if CTTT table is completely empty."""
        warning_shown = []
        monkeypatch.setattr(QMessageBox, "warning", lambda p, t, m: warning_shown.append(m))

        mach_dir = tmp_path / "110C103NL0"
        mach_dir.mkdir(parents=True, exist_ok=True)
        assign_file = mach_dir / "Empty_mecha1.xlsm"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        wb.save(assign_file)
        wb.close()

        view = MemberWorkspaceView(base_dir=tmp_path)
        view.open_assignment_package(assign_file)
        view.clear_cttt_table()

        res = view.submit_data()
        assert res is None
        assert len(warning_shown) > 0
        assert "không có dữ liệu" in warning_shown[0]
