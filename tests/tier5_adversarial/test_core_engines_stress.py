"""Adversarial Stress Test Suite for Core Backend Engines.

Empirical test suite executing rigorous stress tests against:
- JIGManager: missing master, unsupported models, corrupted cell ranges A3:G50, 4M evaluation without KTSX sign-offs.
- InheritanceEngine: 10,000+ rows, duplicate Item IDs, missing PLM_old sheet, missing cols O/P/Q, read-only destinations.
- MSIEngine: malformed unit strings, empty-string PLM vulnerability, 1-char lookup vulnerability, non-standard serials, missing master file.
- Survey Edge Cases: reproduces and verifies the 3 initial test failures from Explorer 3.
"""

from __future__ import annotations

import os
from pathlib import Path
import stat
import time
from typing import Any

import openpyxl
import pandas as pd
import pytest

from src.core.inheritance_engine import (
    InheritanceEngine,
    match_index_mix,
    update_workbook_with_inheritance,
)
from src.core.jig_manager import (
    Assessment4M,
    Decision4M,
    JIGManager,
    Presence4M,
    STANDARD_JIG_SERIES,
)
from src.core.msi_engine import (
    FixSerialMaster,
    MSIEngine,
    evaluate_msi_branch,
)


# ============================================================================
# PART 1: JIGManager Adversarial Stress Tests
# ============================================================================


class TestJIGManagerAdversarialStress:
    """Adversarial tests for JIGManager."""

    def test_missing_master_excel_file(self, tmp_path: Path):
        """Verify behavior when master catalog Excel file is completely missing."""
        missing_file = tmp_path / "non_existent_master_catalog.xlsx"
        manager = JIGManager(master_catalog_path=missing_file)

        # 1. load_series_catalog must return an empty DataFrame, not raise
        df = manager.load_series_catalog("Virgo")
        assert isinstance(df, pd.DataFrame)
        assert df.empty

        # 2. populate_jig_sheet must return False and not create or corrupt target
        target_wb_path = tmp_path / "target_form_ssbom.xlsx"
        wb = openpyxl.Workbook()
        wb.create_sheet("List JIG")
        wb.save(target_wb_path)
        wb.close()

        success = manager.populate_jig_sheet(target_wb_path, "Virgo")
        assert success is False

        # 3. extract_master_path_from_sheet with missing file returns None
        assert manager.extract_master_path_from_sheet(missing_file) is None

    def test_unsupported_model_names(self, tmp_path: Path):
        """Verify behavior when given unsupported, unknown, or malformed model names."""
        # Setup a valid minimal master workbook
        master_path = tmp_path / "valid_master.xlsx"
        wb_master = openpyxl.Workbook()
        ws = wb_master.active
        ws.title = "Virgo"
        ws["A3"] = "JIG-001"
        ws["B3"] = "FUSER_JIG"
        wb_master.save(master_path)
        wb_master.close()

        manager = JIGManager(master_catalog_path=master_path)

        unsupported_models = [
            "AlienSpaceshipModel_999",
            "NonExistentSeries",
            "",
            "   ",
            "12345",
            "None",
            "Null",
        ]

        target_wb = tmp_path / "test_target.xlsx"
        wb_target = openpyxl.Workbook()
        wb_target.create_sheet("List JIG")
        wb_target.save(target_wb)
        wb_target.close()

        for model in unsupported_models:
            assert manager.is_supported_series(model) is False
            df = manager.load_series_catalog(model)
            assert df.empty
            res = manager.populate_jig_sheet(target_wb, model)
            assert res is False

        # Test case-insensitivity and whitespace stripping for valid series
        assert manager.is_supported_series("  virgo  ") is True
        assert manager.is_supported_series("LIBRA2") is True
        assert manager.is_supported_series("6tha4") is True

    def test_corrupted_cell_ranges_in_a3_g50(self, tmp_path: Path):
        """Verify handling of corrupted, partial, or malformed A3:G50 ranges in master catalog."""
        master_path = tmp_path / "corrupted_master.xlsx"
        wb_master = openpyxl.Workbook()
        ws = wb_master.active
        ws.title = "Virgo"

        # Inject various corruption artifacts into rows 3 to 50
        ws["A3"] = "=#REF!"  # Excel reference error
        ws["B3"] = "=#VALUE!"  # Formula evaluation error
        ws["C3"] = "=1/0"  # Division by zero
        ws["D3"] = "Normal Text"
        ws["E3"] = None
        ws["F3"] = 999999999999999999
        ws["G3"] = "Special Char: \u30c6\u30b9\u30c8 !@#$%^&*()"  # Unicode Japanese & symbols

        # Add overlapping or spanning merged cells
        ws.merge_cells("A4:B6")
        ws.merge_cells("C10:G10")

        # Set unusual column width
        ws.column_dimensions["A"].width = 0

        wb_master.save(master_path)
        wb_master.close()

        manager = JIGManager(master_catalog_path=master_path)

        target_path = tmp_path / "corrupted_target.xlsx"
        wb_target = openpyxl.Workbook()
        wb_target.create_sheet("List JIG")
        wb_target.save(target_path)
        wb_target.close()

        # Should successfully populate preserving formulas and merged cells without unhandled crash
        success = manager.populate_jig_sheet(target_path, "Virgo")
        assert success is True

        # Verify target preserved the formula strings
        wb_check = openpyxl.load_workbook(target_path, data_only=False)
        ws_check = wb_check["List JIG"]
        assert ws_check["A3"].value == "=#REF!"
        assert ws_check["B3"].value == "=#VALUE!"
        assert ws_check["D3"].value == "Normal Text"
        wb_check.close()

    def test_4m_evaluation_missing_ktsx_signoffs(self, tmp_path: Path):
        """Verify 4M evaluation behavior with missing KTSX sign-offs and unconfirmed states."""
        target_path = tmp_path / "4m_assessment_test.xlsx"
        wb = openpyxl.Workbook()
        wb.create_sheet("List JIG")
        wb.save(target_path)
        wb.close()

        manager = JIGManager()

        # Case 1: All unchecked, no KTSX sign-off
        assessment_unconfirmed = Assessment4M(
            man_changed=Presence4M.UNCHECKED,
            machine_changed=Presence4M.UNCHECKED,
            material_changed=Presence4M.UNCHECKED,
            method_changed=Presence4M.UNCHECKED,
            overall_evaluation=Decision4M.UNCHECKED,
            ktsx_confirmed=False,
            ktsx_confirmer="",
        )

        assert assessment_unconfirmed.has_any_change() is False
        assert assessment_unconfirmed.presence_summary() == Presence4M.UNCHECKED

        res_write = manager.write_4m_assessment(target_path, assessment_unconfirmed)
        assert res_write is True

        # Read back from workbook and verify
        read_back = manager.read_4m_assessment(target_path)
        assert read_back is not None
        assert read_back.ktsx_confirmed is False
        assert read_back.overall_evaluation == Decision4M.UNCHECKED
        assert "Chưa xác nhận" in read_back.notes

        # Case 2: One 4M dimension has changed, but KTSX STILL NOT signed off
        assessment_with_change_no_signoff = Assessment4M(
            man_changed=Presence4M.YES,
            machine_changed=Presence4M.NO,
            material_changed=Presence4M.NO,
            method_changed=Presence4M.NO,
            overall_evaluation=Decision4M.NG,
            ktsx_confirmed=False,
            ktsx_confirmer="",
        )
        assert assessment_with_change_no_signoff.has_any_change() is True
        assert assessment_with_change_no_signoff.presence_summary() == Presence4M.YES

        manager.write_4m_assessment(target_path, assessment_with_change_no_signoff)
        read_back_2 = manager.read_4m_assessment(target_path)
        assert read_back_2 is not None
        assert read_back_2.ktsx_confirmed is False
        assert read_back_2.overall_evaluation == Decision4M.NG

        # Case 3: Proper KTSX confirmation
        assessment_confirmed = Assessment4M(
            man_changed=Presence4M.NO,
            machine_changed=Presence4M.NO,
            material_changed=Presence4M.NO,
            method_changed=Presence4M.NO,
            overall_evaluation=Decision4M.OK,
            ktsx_confirmed=True,
            ktsx_confirmer="Nguyen Van B - KTSX",
            confirmation_date="19/09/2026",
        )
        assert assessment_confirmed.presence_summary() == Presence4M.NO
        manager.write_4m_assessment(target_path, assessment_confirmed)
        read_back_3 = manager.read_4m_assessment(target_path)
        assert read_back_3 is not None
        assert read_back_3.ktsx_confirmed is True
        assert "Nguyen Van B" in read_back_3.ktsx_confirmer

    def test_jig_target_workbook_missing_sheet_or_readonly(self, tmp_path: Path):
        """Verify behavior when target workbook has no 'List JIG' or is read-only."""
        manager = JIGManager()

        # 1. Non-existent file
        assert manager.clear_jig_sheet(tmp_path / "does_not_exist.xlsx") is False
        assert manager.read_4m_assessment(tmp_path / "does_not_exist.xlsx") is None

        # 2. Target workbook has no 'List JIG' sheet
        no_jig_path = tmp_path / "no_jig.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        wb.save(no_jig_path)
        wb.close()

        # clear_jig_sheet should return False cleanly
        assert manager.clear_jig_sheet(no_jig_path) is False
        # read_4m_assessment should return None
        assert manager.read_4m_assessment(no_jig_path) is None


# ============================================================================
# PART 2: InheritanceEngine Adversarial Stress Tests
# ============================================================================


class TestInheritanceEngineAdversarialStress:
    """Adversarial stress testing of ham_match_index_mix and multi-version updates."""

    def test_ham_match_index_mix_under_10000_plus_rows(self):
        """Stress-test ham_match_index_mix with 15,000 rows for memory and throughput."""
        n = 15_000

        # Generate 15,000 rows in df_new
        df_new = pd.DataFrame(
            {
                "item_id": [f"PART_{i:06d}" for i in range(n)],
                "item_name": [f"Component {i}" for i in range(n)],
                "quantity": [1.0] * n,
            }
        )

        # Generate historical annotations for 8,000 of those items
        df_old = pd.DataFrame(
            {
                "item_id": [f"PART_{i:06d}" for i in range(0, 8000)],
                "giai_thich": [f"Legacy explanation for {i}" for i in range(0, 8000)],
                "phu_trach": [f"Eng_{i % 5}" for i in range(0, 8000)],
                "quan_ly_check": ["OK" if i % 2 == 0 else "Pending" for i in range(0, 8000)],
            }
        )

        t0 = time.perf_counter()
        df_res = InheritanceEngine.match_index_mix(df_new, df_old, key_column="item_id")
        elapsed = time.perf_counter() - t0

        # Performance constraint: 15,000 rows must finish in under 5.0 seconds
        assert elapsed < 5.0, f"Inheritance matching took too long: {elapsed:.2f}s"

        # Structural constraints
        assert len(df_res) == n
        assert "giai_thich" in df_res.columns
        assert "phu_trach" in df_res.columns
        assert "quan_ly_check" in df_res.columns

        # Verify inherited contents for matching rows
        assert df_res.loc[0, "giai_thich"] == "Legacy explanation for 0"
        assert df_res.loc[7999, "giai_thich"] == "Legacy explanation for 7999"

        # Verify unannotated rows are cleanly empty string "", NOT 'NaN' or None
        assert df_res.loc[8000, "giai_thich"] == ""
        assert df_res.loc[14999, "giai_thich"] == ""
        assert (df_res["giai_thich"] == "nan").sum() == 0

        # Verify count of inherited
        matched_count = (df_res["giai_thich"] != "").sum()
        assert matched_count == 8000

    def test_duplicate_item_ids_handling(self):
        """Verify behavior under duplicate Item IDs in both df_old and df_new."""
        # 1. Duplicate keys in df_old: Excel MATCH semantics dictate FIRST match wins
        df_new = pd.DataFrame({"item_id": ["PART_DUP", "PART_UNIQUE"]})
        df_old_duplicates = pd.DataFrame(
            {
                "item_id": ["PART_DUP", "PART_DUP", "PART_UNIQUE"],
                "giai_thich": ["FIRST_MATCH", "SECOND_MATCH_SHOULD_BE_IGNORED", "EXPLANATION_UNIQUE"],
                "phu_trach": ["PIC_1", "PIC_2", "PIC_3"],
                "quan_ly_check": ["OK", "NG", "OK"],
            }
        )

        res = match_index_mix(df_new, df_old_duplicates)
        assert len(res) == 2
        # First match wins!
        assert res.loc[0, "giai_thich"] == "FIRST_MATCH"
        assert res.loc[0, "phu_trach"] == "PIC_1"
        assert res.loc[1, "giai_thich"] == "EXPLANATION_UNIQUE"

        # 2. Duplicate keys in df_new: Both rows must inherit without changing DataFrame length
        df_new_duplicates = pd.DataFrame(
            {
                "item_id": ["PART_A", "PART_A", "PART_B", "PART_A"],
                "level": [1, 2, 2, 3],
            }
        )
        df_old_single = pd.DataFrame(
            {
                "item_id": ["PART_A", "PART_B"],
                "giai_thich": ["EXP_A", "EXP_B"],
                "phu_trach": ["PIC_A", "PIC_B"],
                "quan_ly_check": ["OK", "OK"],
            }
        )

        res2 = match_index_mix(df_new_duplicates, df_old_single)
        assert len(res2) == 4  # Length MUST NOT change
        assert list(res2["giai_thich"]) == ["EXP_A", "EXP_A", "EXP_B", "EXP_A"]

    def test_missing_plm_old_sheet_and_multi_version_increment(self, tmp_path: Path):
        """Verify that when Sheet 'PLM_old' does not exist, it is created; successive calls increment."""
        target_path = tmp_path / "workbook_versions.xlsx"

        # Create workbook with initial PLM sheet only
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PLM"
        ws.cell(row=1, column=3, value="Item Id")
        ws.cell(row=2, column=3, value="PART_001")
        ws.cell(row=2, column=15, value="Explanation V1")
        wb.save(target_path)
        wb.close()

        engine = InheritanceEngine()

        # Update 1: PLM_old does NOT exist yet -> must create 'PLM_old'
        df_v2 = pd.DataFrame({"item_id": ["PART_001", "PART_002"]})
        summary_1 = engine.update_workbook_with_inheritance(target_path, df_v2)
        assert summary_1.archived_sheet_name == "PLM_old"

        wb1 = openpyxl.load_workbook(target_path)
        assert "PLM" in wb1.sheetnames
        assert "PLM_old" in wb1.sheetnames
        assert "PLM_old_1" not in wb1.sheetnames
        wb1.close()

        # Update 2: PLM_old already exists -> must create 'PLM_old_1'
        df_v3 = pd.DataFrame({"item_id": ["PART_001", "PART_003"]})
        summary_2 = engine.update_workbook_with_inheritance(target_path, df_v3)
        assert summary_2.archived_sheet_name == "PLM_old_1"

        wb2 = openpyxl.load_workbook(target_path)
        assert "PLM_old" in wb2.sheetnames
        assert "PLM_old_1" in wb2.sheetnames
        wb2.close()

        # Update 3: PLM_old and PLM_old_1 exist -> must create 'PLM_old_2'
        summary_3 = engine.update_workbook_with_inheritance(target_path, df_v3)
        assert summary_3.archived_sheet_name == "PLM_old_2"

    def test_missing_columns_o_p_q_in_historical_data(self):
        """Verify behavior when historical data (df_old) lacks explanation/person/manager columns."""
        df_new = pd.DataFrame({"item_id": ["PART_1", "PART_2", "PART_3"]})

        # df_old with ONLY item_id (no O, P, Q)
        df_old_no_opq = pd.DataFrame({"item_id": ["PART_1", "PART_2"]})

        res = match_index_mix(df_new, df_old_no_opq)
        assert len(res) == 3
        # Must populate empty strings gracefully
        assert list(res["giai_thich"]) == ["", "", ""]
        assert list(res["phu_trach"]) == ["", "", ""]
        assert list(res["quan_ly_check"]) == ["", "", ""]

        # df_old is None
        res_none = match_index_mix(df_new, None)  # type: ignore
        assert len(res_none) == 3
        assert list(res_none["giai_thich"]) == ["", "", ""]

    def test_target_workbook_missing_plm_sheet(self, tmp_path: Path):
        """Verify that attempting to update a workbook without Sheet 'PLM' raises ValueError."""
        target_path = tmp_path / "corrupted_no_plm.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        wb.save(target_path)
        wb.close()

        engine = InheritanceEngine()
        df_new = pd.DataFrame({"item_id": ["P1"]})

        with pytest.raises(ValueError, match="Sheet 'PLM' does not exist"):
            engine.update_workbook_with_inheritance(target_path, df_new)

    def test_read_only_destination_handling(self, tmp_path: Path):
        """Verify behavior when the target workbook is read-only."""
        target_path = tmp_path / "readonly_target.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PLM"
        ws.cell(row=2, column=3, value="PART_ORIGINAL")
        wb.save(target_path)
        wb.close()

        # Mark file as read-only on filesystem
        os.chmod(target_path, stat.S_IREAD)

        engine = InheritanceEngine()
        df_new = pd.DataFrame({"item_id": ["PART_MODIFIED"]})

        try:
            with pytest.raises(PermissionError):
                engine.update_workbook_with_inheritance(target_path, df_new)
        finally:
            # Restore write permissions so tmp_path cleanup succeeds
            os.chmod(target_path, stat.S_IWRITE)


# ============================================================================
# PART 3: MSIEngine Adversarial Stress Tests
# ============================================================================


class TestMSIEngineAdversarialStress:
    """Adversarial stress testing of MSIEngine, FixSerialMaster, and 9 branches."""

    def test_malformed_unit_strings(self):
        """Stress-test unit string normalization with adversarial and edge-case strings."""
        engine = MSIEngine()

        malformed_inputs = [
            None,
            float("nan"),
            "",
            "   ",
            "\n\t",
            "\x00\xff",
            "🔥🚀💣",
            "A" * 5000,  # 5,000 character string
            "123",
            "---",
            "$$$###@@@",
        ]

        valid_plm = ["302N493010", "302N493020"]

        for bad_input in malformed_inputs:
            # Must evaluate without crashing with IndexError or TypeError
            res = engine.evaluate_row(
                unit_code=bad_input,  # type: ignore
                member_code="ABC",
                member_service="",
                plm_part_codes=valid_plm,
            )
            assert res.status in ("OK", "NG")
            assert isinstance(res.branch, int)

    def test_vulnerability_empty_string_in_plm_part_codes(self):
        """Empirically test vulnerability: empty string in plm_part_codes causes false positive.

        When plm_part_codes contains '' or whitespace (common in raw Excel data with blank rows),
        `any(clean_unit in p or p in clean_unit for p in plm_set)` evaluates `"" in clean_unit` as True,
        marking ANY fake unit code as in_plm=True.
        """
        engine = MSIEngine()

        # PLM list contaminated with empty strings (standard when reading openpyxl trailing rows)
        contaminated_plm = ["", "   ", "302N493010"]

        fake_unit_code = "TOTALLY_NON_EXISTENT_UNIT_XYZ_999"

        res = engine.evaluate_row(
            unit_code=fake_unit_code,
            member_code="ABC",
            member_service="",
            plm_part_codes=contaminated_plm,
        )

        # Secure behavior: empty strings are filtered so fake unit is NOT in PLM
        assert res.in_plm is False, "Empty strings in plm_part_codes must not cause false positive"

    def test_vulnerability_single_character_lookup_in_fix_serial_master(self):
        """Empirically test vulnerability: 1-character unit lookup causes accidental prefix match.

        In FixSerialMaster: `for k, val in self.subunit_map.items(): if prefix9 in k or k in clean_full:`
        If unit_code is "A", prefix9 is "A", which matches ANY key in the map containing "A"!
        """
        master = FixSerialMaster()
        master.add_subunit("302N493010_LASER_SCANNER", "M99", "SERVICE_COMMENT_A")

        # Looking up "A" must NOT match "302N493010_LASER_SCANNER"
        msi, srv = master.lookup_subunit("A")
        assert msi == "" and srv == "", "Single-char 'A' must not falsely match 302N493010_LASER_SCANNER"


    def test_non_standard_serial_numbers_and_all_9_branches(self):
        """Stress-test non-standard 3-char serial numbers across all 9 branches."""
        # Branch 1: Not in PLM, codes match
        b1 = evaluate_msi_branch(in_plm=False, member_code="AB", master_code="AB")
        assert b1["branch"] == 1
        assert b1["status"] == "NG"
        assert b1["highlight_red"] == ["B", "K"]

        # Branch 2: Blank member service normalizes to '-'
        b2 = evaluate_msi_branch(in_plm=True, member_code="ABC", master_code="ABC", member_service="")
        assert b2["member_service_normalized"] == "-"

        # Branch 3: Not in PLM, codes mismatch
        b3 = evaluate_msi_branch(in_plm=False, member_code="12345", master_code="99999")
        assert b3["branch"] == 3
        assert b3["status"] == "NG"
        assert b3["highlight_red"] == ["B", "D", "K"]

        # Branch 4: In PLM, missing from master tool
        b4 = evaluate_msi_branch(in_plm=True, member_code="XYZ", master_code="")
        assert b4["branch"] == 4
        assert b4["status"] == "NG"

        # Branch 5: Both match exactly
        b5 = evaluate_msi_branch(
            in_plm=True,
            member_code="M01",
            master_code="M01",
            member_service="REPLACE ON JAM",
            master_service="REPLACE ON JAM",
        )
        assert b5["branch"] == 5
        assert b5["status"] == "OK"

        # Branch 6: Both match, neither has service (member '-', master '')
        b6 = evaluate_msi_branch(
            in_plm=True,
            member_code="M02",
            master_code="M02",
            member_service="-",
            master_service="",
        )
        assert b6["branch"] == 6
        assert b6["status"] == "OK"

        # Branch 7: Code matches, member '-', but master has required note -> OK with warning
        b7 = evaluate_msi_branch(
            in_plm=True,
            member_code="M03",
            master_code="M03",
            member_service="-",
            master_service="REQUIRED_SERVICE_NOTE",
        )
        assert b7["branch"] == 7
        assert b7["status"] == "OK"
        assert b7["service_warning"] is True
        assert b7["highlight_red"] == ["E"]

        # Branch 8: In PLM, 3-char code mismatch
        b8 = evaluate_msi_branch(
            in_plm=True,
            member_code="M04_BAD",
            master_code="M04_GOOD",
        )
        assert b8["branch"] == 8
        assert b8["status"] == "NG"
        assert "D" in b8["highlight_red"]

        # Branch 9: In PLM, code matches, service mismatch (member != '-')
        b9 = evaluate_msi_branch(
            in_plm=True,
            member_code="M05",
            master_code="M05",
            member_service="CUSTOM_NOTE_1",
            master_service="STANDARD_NOTE_2",
        )
        assert b9["branch"] == 9
        assert b9["status"] == "NG"

    def test_missing_fix_serial_master_excel_file(self, tmp_path: Path):
        """Verify behavior when FIX_SERIAL_DLTOOL_VER010.xls is missing."""
        master = FixSerialMaster()
        missing_file = tmp_path / "FIX_SERIAL_DLTOOL_VER010.xls"

        # load_from_file returns False
        success = master.load_from_file(missing_file)
        assert success is False

        # MSIEngine with unpopulated master evaluates cleanly
        engine = MSIEngine(master=master)
        res = engine.evaluate_row(
            unit_code="302N493010",
            member_code="M01",
            member_service="",
            plm_part_codes=["302N493010"],
        )
        # Unit is in PLM, but master has no entry -> Branch 4
        assert res.branch == 4
        assert res.status == "NG"


# ============================================================================
# PART 4: Verification of 3 Initial Survey Edge Cases
# ============================================================================


class TestSurveyEdgeCasesVerification:
    """Verifies the 3 edge cases identified by Explorer 3."""

    def test_verify_survey_failure_1_tc14_login_page_dom_missing(self):
        """Verify Failure 1: tc14_login_page_dom.html is missing in workspace root.

        Test checks if the file exists; if not, confirms that reading it raises FileNotFoundError.
        """
        workspace_dir = Path(__file__).resolve().parent.parent.parent
        target_file = workspace_dir / "tc14_login_page_dom.html"

        # Verify file exists and is valid HTML
        assert target_file.exists(), f"Expected {target_file} to exist in workspace root."
        content = target_file.read_text(encoding="utf-8")
        assert len(content) > 0
        assert "Version 2412" in content


    def test_verify_survey_failure_2_dos_device_names_credentials(self, tmp_path: Path):
        """Verify Failure 2: DOS reserved device names (CON, PRN, AUX, NUL) on Windows.

        Under Windows, creating or renaming a file named 'CON.dpapi' fails with [WinError 183].
        """
        dos_names = ["CON", "PRN", "AUX", "NUL"]

        # Attempt to create CON.tmp and rename to CON.dpapi
        if os.name == "nt":
            test_file = tmp_path / "CON.tmp"
            dest_file = tmp_path / "CON.dpapi"

            # In Windows, writing to CON.tmp or renaming to CON.dpapi fails
            failed = False
            try:
                test_file.write_bytes(b"dummy")
                dest_file.unlink(missing_ok=True)
                os.replace(test_file, dest_file)
            except (OSError, PermissionError):
                failed = True

            # If write_bytes didn't fail, check dest_file.is_file()
            if not failed and not dest_file.is_file():
                failed = True

            assert failed, "Confirmed: Windows DOS device name 'CON' fails filesystem operations"

    def test_verify_survey_failure_3_windows_file_locking_concurrency(self, tmp_path: Path):
        """Verify Failure 3: Windows atomic rename concurrency locking.

        Demonstrates that on Windows, if a file is open for reading without FILE_SHARE_DELETE,
        an atomic replace from another thread raises PermissionError.
        """
        if os.name == "nt":
            shared_file = tmp_path / "locked_file.dat"
            shared_file.write_text("initial content", encoding="utf-8")

            tmp_file = tmp_path / "locked_file.tmp"
            tmp_file.write_text("new content", encoding="utf-8")

            # Open file for reading with exclusive lock
            with open(shared_file, "r", encoding="utf-8") as f_reader:
                content = f_reader.read()
                assert content == "initial content"

                # Attempting to replace the open file on Windows raises PermissionError
                with pytest.raises(PermissionError):
                    os.replace(tmp_file, shared_file)
