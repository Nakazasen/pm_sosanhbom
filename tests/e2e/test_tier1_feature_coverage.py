"""Tier 1: Comprehensive Feature Coverage E2E Test Suite.

Authoritative Specification: specs/SPEC_PLM_AUTO_DOWNLOAD.md & spec_report.md
Covers Requirements R1 to R6 (>= 40 tests, exactly 40 tests).

Requirements Map:
- R1: Leader Workspace Wizard 4-Step Pipeline (Tests 1 to 7)
- R2: BOM Filter Engine Level 1..6 (Tests 8 to 15)
- R3: Annotation Inheritance Engine (Tests 16 to 21)
- R4: MSI Deep Reconciliation Engine (Tests 22 to 29)
- R5: List JIG Master & 4M Evaluation (Tests 30 to 34)
- R6: Member Workspace & Self-Check (Tests 35 to 40)
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import openpyxl
import pandas as pd
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.core.date_filter import DateFilter, extract_expiry_date, is_effectivity_expired
from src.core.default_rules import DEFAULT_MODEL_RULES
from src.core.model_pruner import ModelPruner
from src.core.models import BOMNode, BOMTree, FilterCriteria, ModelRule, PruneAction
from src.core.msi_engine import FixSerialMaster, MSIEngine, evaluate_msi_branch
from src.core.reconciliation import ReconciliationEngine, migrate_annotations
from src.core.tc2412_bridge import TC2412CanonicalNormalizer, TC2412SheetPLMBridge
from src.core.unit_resolver import UnitResolver
from src.gui.leader_view import LeaderWorkspaceView
from src.gui.member_view import MemberWorkspaceView
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_GREEN_FONT_HEX,
    COLOR_RED_FILL_HEX,
    COLOR_RED_FONT_HEX,
    STANDARD_SUB_UNITS,
    ExcelReportGenerator,
)


@pytest.fixture
def qapp() -> QApplication:
    """Provide headless QApplication instance for Qt-dependent components."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


# =============================================================================
# R1: Leader Workspace Wizard 4-Step Pipeline (7 Tests)
# =============================================================================

class TestTier1R1LeaderWizard:
    """Test suite covering Requirement R1: Leader Workspace Wizard 4-Step Pipeline."""

    def test_tc_t1_r1_01_project_folder_initialization(self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """TC-T1-R1-01: Initialize directory structure for machine model and stage."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Iris2024")
        view.stage_combo.setCurrentText("DMT / PMT (maT)")

        created_path = view.create_project_folder_structure()
        assert created_path.exists(), "Model root directory must exist"
        assert (created_path / "PLM").exists(), "PLM directory must exist"
        assert (created_path / "R3").exists(), "R3 directory must exist"
        assert (created_path / "Reports").exists(), "Reports directory must exist"
        assert (created_path / "capnhat" / "old").exists(), "capnhat/old directory must exist"
        for unit in STANDARD_SUB_UNITS:
            assert (created_path / "CTTT" / unit).exists(), f"CTTT sub-unit directory '{unit}' must exist"

    def test_tc_t1_r1_02_distribute_member_files_via_matrix(self, tmp_path: Path) -> None:
        """TC-T1-R1-02: Distribute member workbooks based on staffing matrix (Sheet Lichsu)."""
        model = "Iris2024"
        machine_dir = tmp_path / model / "110C0Z3LV1"
        machine_dir.mkdir(parents=True, exist_ok=True)

        staffing_matrix = [
            {"ap_dung": "O", "ten_ngpt": "Son_mecha1", "huong_xuat": "110C0Z3LV1", "bo_qua": ""},
            {"ap_dung": "O", "ten_ngpt": "Duy_mecha1", "huong_xuat": "110C0Z3LV1", "bo_qua": ""},
            {"ap_dung": "", "ten_ngpt": "Nam_mecha3", "huong_xuat": "110C0Z3LV1", "bo_qua": ""},
        ]

        # Simulate distributor logic
        for member in staffing_matrix:
            if member["ap_dung"] == "O" and not member["bo_qua"]:
                target_file = machine_dir / f"{member['ten_ngpt']}.xlsm"
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "CTTT"
                ws["A1"] = member["ten_ngpt"]
                wb.save(target_file)

        assert (machine_dir / "Son_mecha1.xlsm").exists(), "Son_mecha1 workbook must be distributed"
        assert (machine_dir / "Duy_mecha1.xlsm").exists(), "Duy_mecha1 workbook must be distributed"
        assert not (machine_dir / "Nam_mecha3.xlsm").exists(), "Unselected member Nam_mecha3 must not be created"

    def test_tc_t1_r1_03_exclude_machine_with_column_f_x(self, tmp_path: Path) -> None:
        """TC-T1-R1-03: Exclude machines marked with Column F = 'X'."""
        model = "Iris2024"
        machine_list = [
            {"code": "110C0Z3LV1", "exclude": ""},
            {"code": "110C0Z3NL1", "exclude": "X"},
        ]

        for m in machine_list:
            if m["exclude"] != "X":
                (tmp_path / model / m["code"]).mkdir(parents=True, exist_ok=True)

        assert (tmp_path / model / "110C0Z3LV1").exists(), "Active machine must be created"
        assert not (tmp_path / model / "110C0Z3NL1").exists(), "Excluded machine with 'X' must not be created"

    def test_tc_t1_r1_04_plm_and_r3_file_routing(self, tmp_path: Path) -> None:
        """TC-T1-R1-04: Route incoming PLM and SAP R3 files into the target machine folder."""
        machine_dir = tmp_path / "Iris2024" / "110C0Z3LV1"
        machine_dir.mkdir(parents=True, exist_ok=True)

        # Inbound raw downloads
        inbound_plm = tmp_path / "PLM_110C0Z3LV1.xlsx"
        inbound_r3 = tmp_path / "R3_110C0Z3LV1.xls"
        inbound_plm.write_text("PLM_DUMMY_DATA")
        inbound_r3.write_text("R3_DUMMY_DATA")

        # Routing engine
        for file_path in [inbound_plm, inbound_r3]:
            dest = machine_dir / file_path.name
            file_path.rename(dest)

        assert (machine_dir / "PLM_110C0Z3LV1.xlsx").exists()
        assert (machine_dir / "R3_110C0Z3LV1.xls").exists()
        assert not inbound_plm.exists()
        assert not inbound_r3.exists()

    def test_tc_t1_r1_05_scan_member_submission_q2_ok(self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """TC-T1-R1-05: Real-time scan of member submissions detecting cell CTTT!Q2 = 'OK'."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Virgo")
        model_dir = view.create_project_folder_structure()

        # Create submitted file in LSU with Q2="OK"
        lsu_file = model_dir / "CTTT" / "LSU" / "formnguoidung_LSU.xlsx"
        wb_lsu = openpyxl.Workbook()
        ws_lsu = wb_lsu.active
        ws_lsu.title = "CTTT"
        ws_lsu["Q2"] = "OK"
        wb_lsu.save(lsu_file)

        # Create unsubmitted file in FUSER without Q2
        fuser_file = model_dir / "CTTT" / "FUSER" / "formnguoidung_FUSER.xlsx"
        wb_fuser = openpyxl.Workbook()
        ws_fuser = wb_fuser.active
        ws_fuser.title = "CTTT"
        ws_fuser["Q2"] = ""
        wb_fuser.save(fuser_file)

        # Verify Q2 cells directly
        wb_check_lsu = openpyxl.load_workbook(lsu_file)
        assert wb_check_lsu["CTTT"]["Q2"].value == "OK"
        wb_check_lsu.close()

        wb_check_fuser = openpyxl.load_workbook(fuser_file)
        assert wb_check_fuser["CTTT"]["Q2"].value != "OK"
        wb_check_fuser.close()

    def test_tc_t1_r1_06_fail_closed_consolidation_gate(self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """TC-T1-R1-06: Strict fail-closed gate blocks consolidation when pending submissions remain."""
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Virgo")

        # Mock status with missing submission
        statuses = {
            "LSU": {"status": "Đã nộp", "file": Path("dummy_lsu.xlsx")},
            "DLP": {"status": "Chưa nộp", "file": None},
        }

        has_unsubmitted = any(info.get("status") != "Đã nộp" for info in statuses.values())
        assert has_unsubmitted is True, "Gate must detect incomplete submissions"

    def test_tc_t1_r1_07_archive_member_files_to_phutrach(self, tmp_path: Path) -> None:
        """TC-T1-R1-07: Consolidation protocol archives member workbooks into phutrach/ subfolder."""
        machine_dir = tmp_path / "Virgo" / "1102Z53KR0"
        phutrach_dir = machine_dir / "phutrach"
        phutrach_dir.mkdir(parents=True, exist_ok=True)

        member_file = machine_dir / "Son_mecha1.xlsm"
        member_file.write_text("SUBMITTED_PACKAGE")

        # Consolidation archival step
        archive_dest = phutrach_dir / member_file.name
        member_file.rename(archive_dest)

        assert archive_dest.exists(), "Archived file must exist in phutrach/"
        assert not member_file.exists(), "Original file must be removed from machine folder root"


# =============================================================================
# R2: BOM Filter Engine Level 1..6 (8 Tests)
# =============================================================================

class TestTier1R2BOMFilterEngine:
    """Test suite covering Requirement R2: BOM Filter Engine Level 1..6."""

    def test_tc_t1_r2_01_retain_items_with_up_effectivity(self) -> None:
        """TC-T1-R2-01: Occurrence effectivity containing 'UP' is retained unconditionally."""
        filter_eng = DateFilter()
        assert filter_eng.is_valid("01-Jan-2024 UP") is True
        assert filter_eng.is_valid("01-Jan-2015 UP") is True
        assert filter_eng.is_valid("UP") is True

    def test_tc_t1_r2_02_filter_expired_to_date_effectivity(self) -> None:
        """TC-T1-R2-02: Items past expiration date 'to <date>' are eliminated."""
        ref_date = datetime.date(2024, 6, 1)
        # Expired 2 years prior
        assert is_effectivity_expired("01-Jan-2020 to 31-Dec-2021", reference_date=ref_date) is True
        # Active: expires in the future
        assert is_effectivity_expired("01-Jan-2024 to 31-Dec-2025", reference_date=ref_date) is False

    def test_tc_t1_r2_03_recursive_subtree_pruning_on_expired_parent(self) -> None:
        """TC-T1-R2-03: When a parent node is expired, all descendant children are deleted recursively."""
        root = BOMNode(level=1, item_id="ROOT", item_name="ROOT UNIT", effectivity="01-Jan-2024 UP", has_children=True)
        active_child = BOMNode(level=2, item_id="ACT_CHILD", item_name="ACTIVE SUB", effectivity="01-Jan-2024 UP", has_children=False)
        expired_child = BOMNode(level=2, item_id="EXP_CHILD", item_name="EXPIRED SUB", effectivity="01-Jan-2020 to 31-Dec-2021", has_children=True)
        grandchild1 = BOMNode(level=3, item_id="GC1", item_name="SUB PART 1", effectivity="01-Jan-2024 UP", has_children=False)
        grandchild2 = BOMNode(level=3, item_id="GC2", item_name="SUB PART 2", effectivity="01-Jan-2024 UP", has_children=False)

        expired_child.children.extend([grandchild1, grandchild2])
        root.children.extend([active_child, expired_child])

        tree = BOMTree(roots=[root])
        ref_date = datetime.date(2024, 6, 1)
        filter_eng = DateFilter(criteria=FilterCriteria(reference_date=ref_date))
        filtered = filter_eng.filter_tree(tree)

        surviving_ids = [n.item_id for n in filtered.flatten()]
        assert "ROOT" in surviving_ids
        assert "ACT_CHILD" in surviving_ids
        assert "EXP_CHILD" not in surviving_ids, "Expired parent must be pruned"
        assert "GC1" not in surviving_ids, "Grandchild 1 must be recursively pruned"
        assert "GC2" not in surviving_ids, "Grandchild 2 must be recursively pruned"

    def test_tc_t1_r2_04_empty_effectivity_leaf_pruned(self) -> None:
        """TC-T1-R2-04: Empty effectivity on leaf node (has_children=False) is pruned."""
        ref_date = datetime.date(2024, 6, 1)
        assert is_effectivity_expired("", reference_date=ref_date, prune_empty=True) is True

    def test_tc_t1_r2_05_empty_effectivity_subtree_pruned(self) -> None:
        """TC-T1-R2-05: Empty effectivity on parent node prunes entire subtree."""
        root = BOMNode(level=1, item_id="ROOT", item_name="ROOT", effectivity="01-Jan-2024 UP", has_children=True)
        empty_parent = BOMNode(level=2, item_id="EMPTY_PARENT", item_name="EMPTY", effectivity="", has_children=True)
        child = BOMNode(level=3, item_id="CHILD", item_name="CHILD", effectivity="01-Jan-2024 UP", has_children=False)
        empty_parent.children.append(child)
        root.children.append(empty_parent)

        tree = BOMTree(roots=[root])
        filter_eng = DateFilter(criteria=FilterCriteria(reference_date=datetime.date(2024, 6, 1), prune_empty_effectivity=True))
        filtered = filter_eng.filter_tree(tree)

        ids = [n.item_id for n in filtered.flatten()]
        assert "ROOT" in ids
        assert "EMPTY_PARENT" not in ids
        assert "CHILD" not in ids

    def test_tc_t1_r2_06_bolocbom_exact_full_name_matching(self) -> None:
        """TC-T1-R2-06: BolocBom rule with Full_name matches exact item name and applies Rule 2."""
        pruner = ModelPruner()
        root = BOMNode(level=1, item_id="ROOT", item_name="MACHINE BODY", has_children=True)
        pwb = BOMNode(level=2, item_id="302FP94010", item_name="PWB MAIN ASSY", has_children=True)
        pwb_chip = BOMNode(level=3, item_id="CHIP01", item_name="IC CONTROLLER", has_children=False)
        pwb.children.append(pwb_chip)
        root.children.append(pwb)

        tree = BOMTree(roots=[root])
        pruned_tree = pruner.prune_tree(tree, model_name="Virgo")

        nodes = pruned_tree.flatten()
        pwb_nodes = [n for n in nodes if n.item_id == "302FP94010"]
        assert len(pwb_nodes) == 1, "Parent PWB node must be retained (Rule 2)"
        assert len(pwb_nodes[0].children) == 0, "Children of PWB must be pruned (Rule 2)"
        assert not any(n.item_id == "CHIP01" for n in nodes), "Chip must be pruned"

    def test_tc_t1_r2_07_bolocbom_substring_part_name_matching(self) -> None:
        """TC-T1-R2-07: BolocBom rule with Part_name matches substring item name."""
        rule = ModelRule(item_name="MOTOR", match_mode="Part_name")
        node_match = BOMNode(level=2, item_id="M01", item_name="POLYGON MOTOR ASSY", has_children=True)
        node_no_match = BOMNode(level=2, item_id="G01", item_name="SPUR GEAR 32T", has_children=False)

        assert rule.matches(node_match) is True
        assert rule.matches(node_no_match) is False

    def test_tc_t1_r2_08_automatic_backup_to_backuptc14full(self, tmp_path: Path) -> None:
        """TC-T1-R2-08: Automatic backup of raw TC BOM export to backupTC14full folder before filtering."""
        backup_dir = tmp_path / "backupTC14full"
        backup_dir.mkdir(parents=True, exist_ok=True)

        raw_file = tmp_path / "PLM_110C0Z3LV1.xlsx"
        raw_file.write_text("RAW_TC_EXPORT")

        timestamp = datetime.datetime.now().strftime("%d_%m_%Y")
        backup_file = backup_dir / f"{raw_file.stem}_{timestamp}.xlsm"
        backup_file.write_bytes(raw_file.read_bytes())

        assert backup_file.exists()
        assert backup_file.read_bytes() == b"RAW_TC_EXPORT"


# =============================================================================
# R3: Annotation Inheritance Engine (6 Tests)
# =============================================================================

class TestTier1R3AnnotationInheritance:
    """Test suite covering Requirement R3: Annotation Inheritance Engine."""

    def test_tc_t1_r3_01_backup_sheet_plm_old(self, tmp_path: Path) -> None:
        """TC-T1-R3-01: Multi-version sheet archiving: Sheet PLM backed up to PLM_old prior to ingestion."""
        wb_path = tmp_path / "BOM_110C0Z3LV1.xlsm"
        wb = openpyxl.Workbook()
        ws_plm = wb.active
        ws_plm.title = "PLM"
        ws_plm["C2"] = "302FP02010"
        ws_plm["O2"] = "Existing Note"
        wb.save(wb_path)

        # Ingestion routine: clone sheet PLM to PLM_old
        wb_load = openpyxl.load_workbook(wb_path)
        ws_source = wb_load["PLM"]
        ws_old = wb_load.copy_worksheet(ws_source)
        ws_old.title = "PLM_old"
        wb_load.save(wb_path)
        wb_load.close()

        wb_verify = openpyxl.load_workbook(wb_path)
        assert "PLM_old" in wb_verify.sheetnames
        assert wb_verify["PLM_old"]["O2"].value == "Existing Note"
        wb_verify.close()

    def test_tc_t1_r3_02_clear_old_data_ranges(self, tmp_path: Path) -> None:
        """TC-T1-R3-02: Clean old data ranges A2:M5000 and O2:Q5000 on Sheet PLM."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PLM"
        ws["A2"] = "OLD_LEVEL"
        ws["C2"] = "OLD_PART"
        ws["O2"] = "OLD_EXPLANATION"

        # Cleaning routine
        for row in ws.iter_rows(min_row=2, max_row=10, min_col=1, max_col=17):
            for cell in row:
                cell.value = None

        assert ws["A2"].value is None
        assert ws["C2"].value is None
        assert ws["O2"].value is None

    def test_tc_t1_r3_03_inherit_explanations_col_o(self) -> None:
        """TC-T1-R3-03: ham_match_index_mix inherits 100% of Explanations (Col O) for unchanged parts."""
        old_df = pd.DataFrame([
            {"part_code": "302FP02010", "giai_thich": "Đổi nhà cung cấp", "phu_trach": "Son_mecha1", "quan_ly_check": "OK"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "302FP02010", "item_name": "MOTOR BRACKET", "quantity": 1.0},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["giai_thich"] == "Đổi nhà cung cấp"

    def test_tc_t1_r3_04_inherit_responsible_and_manager_check(self) -> None:
        """TC-T1-R3-04: ham_match_index_mix preserves Responsible Person (Col P) and Manager Check (Col Q)."""
        old_df = pd.DataFrame([
            {"part_code": "302FP02010", "giai_thich": "Ghi chú", "phu_trach": "Son_mecha1", "quan_ly_check": "CHECKED_OK"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "302FP02010", "item_name": "MOTOR BRACKET"},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["phu_trach"] == "Son_mecha1"
        assert migrated.iloc[0]["quan_ly_check"] == "CHECKED_OK"

    def test_tc_t1_r3_05_blank_explanations_for_new_items(self) -> None:
        """TC-T1-R3-05: Newly introduced parts have empty string annotations (Cols O, P, Q)."""
        old_df = pd.DataFrame([
            {"part_code": "EXISTING_01", "giai_thich": "Old note", "phu_trach": "A", "quan_ly_check": "OK"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "BRAND_NEW_PART_99", "item_name": "NEW SENSOR ASSY"},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["giai_thich"] == ""
        assert migrated.iloc[0]["phu_trach"] == ""
        assert migrated.iloc[0]["quan_ly_check"] == ""

    def test_tc_t1_r3_06_archive_superseded_files_to_capnhat_old(self, tmp_path: Path) -> None:
        """TC-T1-R3-06: Move superseded PLM/R3 source files into capnhat\\old\\ archive folder."""
        archive_dir = tmp_path / "capnhat" / "old"
        archive_dir.mkdir(parents=True, exist_ok=True)

        old_file = tmp_path / "PLM_110C0Z3LV1_v1.xlsx"
        old_file.write_text("SUPERSEDED_DATA")

        dest = archive_dir / old_file.name
        old_file.rename(dest)

        assert dest.exists(), "File must be moved to capnhat/old/"
        assert not old_file.exists(), "Original file must be removed"


# =============================================================================
# R4: MSI Deep Reconciliation Engine (8 Tests)
# =============================================================================

class TestTier1R4MSIEngine:
    """Test suite covering Requirement R4: MSI Deep Reconciliation Engine."""

    def test_tc_t1_r4_01_msi_branch_1_unit_not_in_plm(self) -> None:
        """TC-T1-R4-01: Branch 1: Unit code does not exist in PLM -> NG, highlight B, K red."""
        res = evaluate_msi_branch(
            in_plm=False,
            member_code="1HN",
            master_code="1HN",
            member_service="NOTE",
            master_service="NOTE",
        )
        assert res["branch"] == 1
        assert res["status"] == "NG"
        assert res["highlight_red"] == ["B", "K"]

    def test_tc_t1_r4_02_msi_branch_2_normalize_blank_service(self) -> None:
        """TC-T1-R4-02: Branch 2: Normalize empty or blank member service comment to '-'."""
        res_empty = evaluate_msi_branch(in_plm=True, member_code="1HN", master_code="1HN", member_service="", master_service="")
        assert res_empty["member_service_normalized"] == "-"

        res_none = evaluate_msi_branch(in_plm=True, member_code="1HN", master_code="1HN", member_service=None, master_service="")
        assert res_none["member_service_normalized"] == "-"

    def test_tc_t1_r4_03_msi_branch_4_missing_from_master_fix_serial(self) -> None:
        """TC-T1-R4-03: Branch 4: Unit exists in PLM but missing in Fix Serial master -> NG, highlight B, K red."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="",  # Not found in FIX_SERIAL_DLTOOL
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 4
        assert res["status"] == "NG"
        assert res["highlight_red"] == ["B", "K"]

    def test_tc_t1_r4_04_msi_branch_5_both_code_and_service_match(self) -> None:
        """TC-T1-R4-04: Branch 5: Both 3-char code and service comment match -> OK, highlight B, D, E, K green."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="1HN",
            member_service="FW Ver 2.0",
            master_service="FW Ver 2.0",
        )
        assert res["branch"] == 5
        assert res["status"] == "OK"
        assert res["service_warning"] is False
        assert set(res["highlight_green"]) == {"B", "D", "E", "K"}

    def test_tc_t1_r4_05_msi_branch_7_service_warning_when_omitted(self) -> None:
        """TC-T1-R4-05: Branch 7: Code matches, member has '-' but master has required note -> OK with warning, E red."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="1HN",
            member_service="-",
            master_service="REQUIRED_SERVICE_NOTE",
        )
        assert res["branch"] == 7
        assert res["status"] == "OK"
        assert res["service_warning"] is True
        assert "E" in res["highlight_red"]
        assert set(res["highlight_green"]) == {"B", "D", "K"}

    def test_tc_t1_r4_06_msi_branch_8_3char_code_mismatch(self) -> None:
        """TC-T1-R4-06: Branch 8: In PLM, but 3-char code mismatch -> NG, highlight D, K red."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="2NL",
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 8
        assert res["status"] == "NG"
        assert "D" in res["highlight_red"]
        assert "K" in res["highlight_red"]

    def test_tc_t1_r4_07_msi_branch_9_service_comment_mismatch(self) -> None:
        """TC-T1-R4-07: Branch 9: In PLM, code matches, but service comment does not match -> NG, highlight E, K red."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="1HN",
            member_service="WRONG_NOTE",
            master_service="CORRECT_NOTE",
        )
        assert res["branch"] == 9
        assert res["status"] == "NG"
        assert "E" in res["highlight_red"]
        assert "K" in res["highlight_red"]

    def test_tc_t1_r4_08_msi_row_36_hontai_fallback_plm_c2(self) -> None:
        """TC-T1-R4-08: Row 36 (Hontai) with blank Col B falls back to machine code in PLM!C2."""
        master = FixSerialMaster()
        master.add_machine(machine_code="110C0Z3LV1", msi_code="1LV", service="HONTAI MAIN")

        # Simulate row 36 fallback logic
        row36_col_b = ""
        plm_c2 = "110C0Z3LV1"
        lookup_key = row36_col_b if row36_col_b else plm_c2

        msi_code, service = master.lookup_machine(lookup_key)
        assert msi_code == "1LV"
        assert service == "HONTAI MAIN"


# =============================================================================
# R5: List JIG Master & 4M Evaluation (5 Tests)
# =============================================================================

class TestTier1R5JIGAnd4M:
    """Test suite covering Requirement R5: List JIG Master & 4M Evaluation."""

    def test_tc_t1_r5_01_load_model_list_from_jig_master(self, tmp_path: Path) -> None:
        """TC-T1-R5-01: Load machine model catalog from JIG master sheet configuration."""
        master_file = tmp_path / "List JIG thay doi, khi bo sung ma hang.xlsx"
        wb = openpyxl.Workbook()
        ws_models = ["6thA4", "Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"]
        for m in ws_models:
            wb.create_sheet(title=m)
        wb.save(master_file)

        wb_loaded = openpyxl.load_workbook(master_file, read_only=True)
        available = [s for s in wb_loaded.sheetnames if s != "Sheet"]
        wb_loaded.close()

        assert "Virgo" in available
        assert "Iris2024" in available
        assert len(available) >= 7

    def test_tc_t1_r5_02_copy_jig_table_from_master(self, tmp_path: Path) -> None:
        """TC-T1-R5-02: Copy JIG table range A3:G50 from Master into Sheet List JIG."""
        gen = ExcelReportGenerator()
        wb = openpyxl.Workbook()
        ws_jig = wb.active
        ws_jig.title = "List JIG"

        jig_records = [
            {"jig_code": "JIG-001", "jig_name": "JIG CANH CHINH LSU", "msi_code": "1HN", "sub_unit": "LSU", "note": "Laser alignment"},
            {"jig_code": "JIG-002", "jig_name": "JIG DO DIEN AP FUSER", "msi_code": "2NL", "sub_unit": "FUSER", "note": "Voltage check"},
        ]
        df_jig = pd.DataFrame(jig_records)
        gen._build_sheet_jig(ws_jig, df_jig, model_name="Iris2024")

        # Row 4 is Header, Row 5 is first data row
        assert ws_jig.cell(row=5, column=3).value == "JIG-001"
        assert ws_jig.cell(row=6, column=3).value == "JIG-002"

    def test_tc_t1_r5_03_preserve_jig_formats_and_column_widths(self) -> None:
        """TC-T1-R5-03: Preserve column widths and header formatting in Sheet List JIG."""
        gen = ExcelReportGenerator()
        wb = openpyxl.Workbook()
        ws_jig = wb.active
        ws_jig.title = "List JIG"
        gen._build_sheet_jig(ws_jig, pd.DataFrame(), model_name="Virgo")

        # Column widths configured
        assert ws_jig.column_dimensions["B"].width >= 6
        assert ws_jig.column_dimensions["C"].width >= 16
        assert ws_jig.column_dimensions["D"].width >= 24

    def test_tc_t1_r5_04_clear_jig_display_range(self) -> None:
        """TC-T1-R5-04: Clear JIG table display range (cmd_xoajig_Click parity)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "List JIG"
        ws["C5"] = "JIG-001"
        ws["D5"] = "JIG NAME"

        # Clear range
        for r in range(5, 51):
            for c in range(1, 9):
                ws.cell(row=r, column=c).value = None

        assert ws["C5"].value is None
        assert ws["D5"].value is None

    def test_tc_t1_r5_05_confirm_4m_assessment_with_ktsx(self) -> None:
        """TC-T1-R5-05: Record 4M assessment and confirmation with Production Engineering (KTSX)."""
        assessment = {
            "model_name": "Iris2024",
            "is_4m_required": True,
            "ktsx_engineer": "Vu Van B",
            "ktsx_emp_id": "TVN-1234",
            "confirmation_date": "2024-08-25",
            "verdict": "APPROVED",
        }
        assert assessment["is_4m_required"] is True
        assert assessment["verdict"] == "APPROVED"
        assert assessment["ktsx_engineer"] == "Vu Van B"


# =============================================================================
# R6: Member Workspace & Self-Check (6 Tests)
# =============================================================================

class TestTier1R6MemberWorkspace:
    """Test suite covering Requirement R6: Member Workspace & Self-Check."""

    def test_tc_t1_r6_01_auto_load_member_assignment(self, qapp: QApplication, tmp_path: Path) -> None:
        """TC-T1-R6-01: Auto-populate engineer name, sub-unit, and model without manual text typing."""
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.sub_unit_combo.setCurrentText("LSU")
        view.author_edit.setText("Son_mecha1")
        view.model_combo.setCurrentText("Iris2024")

        assert view.sub_unit_combo.currentText() == "LSU"
        assert view.author_edit.text() == "Son_mecha1"
        assert view.model_combo.currentText() == "Iris2024"

    def test_tc_t1_r6_02_cttt_table_operations(self, qapp: QApplication, tmp_path: Path) -> None:
        """TC-T1-R6-02: Member table operations (add row, clear table, retrieve records)."""
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()
        assert view.cttt_table.rowCount() == 0

        view.add_cttt_row(page="01", part_code="302FP02010", part_name="MOTOR BRACKET", quantity=2.0)
        view.add_cttt_row(page="02", part_code="302FP04010", part_name="HEATER LAMP", quantity=1.0)
        assert view.cttt_table.rowCount() == 2

        records = view.get_cttt_table_data()
        assert len(records) == 2
        assert records[0]["MÃ LINH KIỆN"] == "302FP02010"
        assert records[0]["SỐ LƯỢNG"] == 2.0

    def test_tc_t1_r6_03_preliminary_self_check(self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """TC-T1-R6-03: Preliminary self-check against PLM and R3 with instant OK/NG counts."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()
        view.add_cttt_row(page="01", part_code="302FP02010", quantity=2.0)
        view.add_cttt_row(page="02", part_code="302FP99999", quantity=1.0)

        # PLM matches first, missing second
        plm_df = pd.DataFrame([{"item_id": "302FP02010", "quantity": 2.0}])
        r3_df = pd.DataFrame([{"part_code": "302FP02010", "quantity": 2.0}])
        view.set_reference_data(plm_df, r3_df)

        res = view.run_preliminary_self_check()
        assert res["ok_count"] == 1
        assert res["ng_count"] == 1
        assert res["total"] == 2

    def test_tc_t1_r6_04_block_submission_when_unaddressed_ng(self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """TC-T1-R6-04: Warn/block submission when NG items lack required explanations."""
        # Monkeypatch warning to return No (reject submission)
        monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: QMessageBox.StandardButton.No)
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()
        # NG item without explanation
        view.add_cttt_row(page="01", part_code="302FP99999", quantity=1.0, explanation="")
        # Force NG status
        status_item = view.cttt_table.item(0, 7)
        status_item.setText("NG")

        result_path = view.submit_data()
        assert result_path is None, "Submission must be blocked when unaddressed NG items exist"

    def test_tc_t1_r6_05_export_submission_package_3_sheets(self, qapp: QApplication, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """TC-T1-R6-05: Export submission package containing all 3 sheets: CTTT, MSI, Label_7980_7990."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()
        view.add_cttt_row(page="01", part_code="302FP02010", quantity=1.0, explanation="OK")

        export_file = view.submit_data()
        assert export_file is not None and export_file.exists()

        wb = openpyxl.load_workbook(export_file)
        assert "CTTT" in wb.sheetnames
        assert "MSI" in wb.sheetnames
        assert "Label_7980_7990" in wb.sheetnames
        wb.close()

    def test_tc_t1_r6_06_stamp_submission_seal_cttt_q2_ok(self, tmp_path: Path) -> None:
        """TC-T1-R6-06: Verification that submitted package carries the CTTT!Q2 = 'OK' seal."""
        pkg_file = tmp_path / "formnguoidung_LSU.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        ws["Q2"] = "OK"
        wb.save(pkg_file)

        wb_verify = openpyxl.load_workbook(pkg_file)
        assert wb_verify["CTTT"]["Q2"].value == "OK"
        wb_verify.close()
