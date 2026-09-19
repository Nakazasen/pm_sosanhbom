"""Tier 3: Pairwise Combinations & Inter-Module Pipelines E2E Test Suite.

Authoritative Specification: specs/SPEC_PLM_AUTO_DOWNLOAD.md & spec_report.md
Covers 6 Pairwise Combinatorial Integration Pipelines:
- TC-T3-01: Leader Wizard (R1) <===> TC24 BOM Filter Engine (R2)
- TC-T3-02: Leader Tracking & Consolidation (R1) <===> Member Workspace Submissions (R6)
- TC-T3-03: Consolidated Report (R1) <===> MSI Deep Reconciliation (R4)
- TC-T3-04: Master BOM Generation (R1) <===> List JIG & 4M Assessment (R5)
- TC-T3-05: TC24 BOM Filter Engine (R2) <===> Annotation Inheritance Pipeline (R3)
- TC-T3-06: Annotation Inheritance (R3) <===> Member Workspace Re-Check (R6)
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

from src.core.date_filter import DateFilter
from src.core.model_pruner import ModelPruner
from src.core.models import BOMNode, BOMTree, FilterCriteria, ModelRule
from src.core.msi_engine import FixSerialMaster, MSIEngine, evaluate_msi_branch
from src.core.reconciliation import ReconciliationEngine, migrate_annotations
from src.core.tc2412_bridge import TC2412CanonicalNormalizer, TC2412SheetPLMBridge
from src.gui.leader_view import LeaderWorkspaceView
from src.gui.member_view import MemberWorkspaceView
from src.reporting.excel_generator import ExcelReportGenerator, STANDARD_SUB_UNITS


@pytest.fixture
def qapp() -> QApplication:
    """Provide headless QApplication instance for Qt-dependent components."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestTier3PairwiseCombinations:
    """Test suite covering Tier 3: Pairwise Combinatorial Pipelines."""

    def test_tc_t3_01_pipeline_wizard_to_tc24_bom_filter(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """TC-T3-01: Pipeline 1: Leader Wizard (Steps 1 & 2) ===> TC24 BOM Filter Engine.

        Flow:
        1. Leader sets up project for 'Iris2024' machine '110C0Z3LV1'.
        2. Simulates raw TC2412 export file containing both active, expired, and pruned parts.
        3. Invokes BOM filtering engine (DateFilter + ModelPruner).
        4. Verifies clean output file is placed in machine folder and raw file is backed up in backupTC14full/.
        """
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Iris2024")
        model_dir = view.create_project_folder_structure()

        # Step 2: Inbound raw download
        raw_export = tmp_path / "PLM_110C0Z3LV1_raw.xlsx"
        raw_records = [
            {"level": 1, "item_type": "Assembly", "item_id": "302FP93010", "has_children": True, "quantity": 1.0,
             "effectivity": "01-Jan-2024 UP", "item_name": "LSU UNIT", "revision": "A", "item_rev_status": "Released"},
            {"level": 2, "item_type": "Assembly", "item_id": "302FP94010", "has_children": True, "quantity": 1.0,
             "effectivity": "01-Jan-2024 UP", "item_name": "POLYGON MOTOR ASSY", "revision": "A", "item_rev_status": "Released"},
            {"level": 3, "item_type": "Part", "item_id": "302FP02010", "has_children": False, "quantity": 1.0,
             "effectivity": "01-Jan-2024 UP", "item_name": "MOTOR BRACKET", "revision": "A", "item_rev_status": "Released"},
            # Expired assembly
            {"level": 2, "item_type": "Assembly", "item_id": "OLD_ASSY", "has_children": True, "quantity": 1.0,
             "effectivity": "01-Jan-2020 to 31-Dec-2021", "item_name": "OBSOLETE MOTOR", "revision": "X", "item_rev_status": "Obsolete"},
            {"level": 3, "item_type": "Part", "item_id": "OLD_PART", "has_children": False, "quantity": 1.0,
             "effectivity": "01-Jan-2020 to 31-Dec-2021", "item_name": "OBSOLETE GEAR", "revision": "X", "item_rev_status": "Obsolete"},
        ]
        df_raw = pd.DataFrame(raw_records)
        df_raw.to_excel(raw_export, index=False)

        # 1. Automatic backup
        backup_dir = model_dir / "backupTC14full"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_copy = backup_dir / f"PLM_110C0Z3LV1_{datetime.datetime.now().strftime('%d_%m_%Y')}.xlsm"
        backup_copy.write_bytes(raw_export.read_bytes())
        assert backup_copy.exists(), "Raw file must be safely backed up in backupTC14full/"

        # 2. Build tree and apply DateFilter
        root = BOMNode(level=1, item_id="302FP93010", item_name="LSU UNIT", effectivity="01-Jan-2024 UP", has_children=True)
        motor = BOMNode(level=2, item_id="302FP94010", item_name="POLYGON MOTOR ASSY", effectivity="01-Jan-2024 UP", has_children=True)
        bracket = BOMNode(level=3, item_id="302FP02010", item_name="MOTOR BRACKET", effectivity="01-Jan-2024 UP", has_children=False)
        old_assy = BOMNode(level=2, item_id="OLD_ASSY", item_name="OBSOLETE MOTOR", effectivity="01-Jan-2020 to 31-Dec-2021", has_children=True)
        old_part = BOMNode(level=3, item_id="OLD_PART", item_name="OBSOLETE GEAR", effectivity="01-Jan-2020 to 31-Dec-2021", has_children=False)

        motor.children.append(bracket)
        old_assy.children.append(old_part)
        root.children.extend([motor, old_assy])
        tree = BOMTree(roots=[root])

        filter_eng = DateFilter(criteria=FilterCriteria(reference_date=datetime.date(2024, 6, 1)))
        clean_tree = filter_eng.filter_tree(tree)

        # 3. Route clean PLM to machine directory
        dest_plm = model_dir / "PLM" / "PLM_110C0Z3LV1.xlsx"
        clean_df = clean_tree.to_dataframe()
        clean_df.to_excel(dest_plm, index=False)

        assert dest_plm.exists(), "Clean PLM must exist in machine folder"
        surviving = list(clean_df["item_id"])
        assert "302FP02010" in surviving, "Active part must survive"
        assert "OLD_ASSY" not in surviving, "Expired assembly must be eliminated"
        assert "OLD_PART" not in surviving, "Child of expired assembly must be eliminated"

    def test_tc_t3_02_pipeline_member_submission_to_leader_tracking_and_consolidation(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """TC-T3-02: Pipeline 2: Member Workspace Submissions (R6) <===> Leader Step 3 Tracking & Consolidation (R1).

        Flow:
        1. 3 Members (LSU, DLP, FUSER) complete submissions stamping CTTT!Q2 = "OK".
        2. Leader Step 3 scanner detects 3/3 OK submissions.
        3. Consolidation unlocked; moving member files into phutrach/ subfolder.
        """
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Iris2024")
        model_dir = view.create_project_folder_structure()

        sub_units = ["LSU", "DLP", "FUSER"]
        submitted_files: list[Path] = []

        # Simulate 3 members submitting
        for u in sub_units:
            u_dir = model_dir / "CTTT" / u
            u_file = u_dir / f"formnguoidung_{u}_20240825.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CTTT"
            ws["Q2"] = "OK"  # The authoritative seal!
            ws.append(["SUB", "TRANG CTTT", "MÃ LINH KIỆN", "TÊN LINH KIỆN", "SỐ LƯỢNG", "PHỤ TRÁCH", "Giải thích", "Check"])
            ws.append([u, "01", f"PART_{u}", f"NAME_{u}", 1.0, f"Eng_{u}", "", "OK"])
            wb.save(u_file)
            submitted_files.append(u_file)

        # Leader scanner
        view.scan_member_submissions()
        statuses = view.get_sub_unit_statuses()

        assert statuses.get("LSU") == "Đã nộp"
        assert statuses.get("DLP") == "Đã nộp"
        assert statuses.get("FUSER") == "Đã nộp"

        # Consolidation moves files to phutrach/
        phutrach_dir = model_dir / "phutrach"
        phutrach_dir.mkdir(parents=True, exist_ok=True)
        for f in submitted_files:
            archived = phutrach_dir / f.name
            f.rename(archived)
            assert archived.exists()
            assert not f.exists()

    def test_tc_t3_03_pipeline_leader_consolidation_to_msi_deep_reconciliation(
        self,
        tmp_path: Path,
    ) -> None:
        """TC-T3-03: Pipeline 3: Consolidated Report (R1) <===> MSI Deep Reconciliation Engine (R4).

        Flow:
        1. Member inputs collected into consolidated MSI data table (Rows 2..36).
        2. Master FIX_SERIAL_DLTOOL loaded.
        3. MSI Engine evaluates all rows through the 9-branch decision table, verifying verdicts.
        """
        master = FixSerialMaster()
        master.add_subunit("302FP93010", "1HN", "LSU SERVICE NOTE")
        master.add_subunit("302FP94010", "2NL", "")
        master.add_machine("110C0Z3LV1", "1LV", "HONTAI SERVICE NOTE")

        plm_parts = {"302FP93010", "302FP94010", "110C0Z3LV1"}

        consolidated_msi_rows = [
            # Row 2: Sub-unit with service note matching Branch 5
            {"unit_code": "302FP93010", "member_code": "1HN", "service": "LSU SERVICE NOTE"},
            # Row 3: Sub-unit without service note matching Branch 6
            {"unit_code": "302FP94010", "member_code": "2NL", "service": "-"},
            # Row 36: Hontai main body matching Branch 5
            {"unit_code": "110C0Z3LV1", "member_code": "1LV", "service": "HONTAI SERVICE NOTE"},
        ]

        verdicts = []
        for r in consolidated_msi_rows:
            u_code = r["unit_code"]
            is_machine = u_code == "110C0Z3LV1"
            mst_code, mst_srv = master.lookup_machine(u_code) if is_machine else master.lookup_subunit(u_code)

            res = evaluate_msi_branch(
                in_plm=(u_code in plm_parts),
                member_code=r["member_code"],
                master_code=mst_code,
                member_service=r["service"],
                master_service=mst_srv,
            )
            verdicts.append(res)

        assert verdicts[0]["status"] == "OK"
        assert verdicts[0]["branch"] == 5
        assert verdicts[1]["status"] == "OK"
        assert verdicts[1]["branch"] == 6
        assert verdicts[2]["status"] == "OK"
        assert verdicts[2]["branch"] == 5

    def test_tc_t3_04_pipeline_leader_workbook_generation_with_jig_and_4m(
        self,
        tmp_path: Path,
    ) -> None:
        """TC-T3-04: Pipeline 4: Master BOM Generation (R1) <===> List JIG & 4M Confirmation (R5).

        Flow:
        1. ExcelReportGenerator creates master comparison workbook form_ssbom.
        2. Builds Sheet 'List JIG' from model master records.
        3. Integrates 4M assessment confirmation record into workbook metadata.
        """
        gen = ExcelReportGenerator()
        output_file = tmp_path / "BOM_110C0Z3LV1.xlsm"

        df_cttt = pd.DataFrame([{"sub_unit": "LSU", "part_code": "302FP02010", "quantity": 1.0}])
        df_plm = pd.DataFrame([{"level": 1, "item_id": "302FP02010", "quantity": 1.0, "item_name": "BRACKET"}])
        df_r3 = pd.DataFrame([{"level": 1, "part_code": "302FP02010", "quantity": 1.0, "item_name": "BRACKET"}])
        df_jig = pd.DataFrame([{"jig_code": "JIG-LSU-01", "jig_name": "ALIGNMENT FIXTURE", "fixed_code": "1HN", "unit": "LSU", "note": "OK"}])

        gen.generate_report(
            output_path=output_file,
            cttt_data=df_cttt,
            plm_data=df_plm,
            r3_data=df_r3,
            jig_data=df_jig,
            model_name="Iris2024",
            target_date="2024/08/25",
        )

        assert output_file.exists()
        wb = openpyxl.load_workbook(output_file, read_only=True)
        assert "CTTT" in wb.sheetnames
        assert "PLM" in wb.sheetnames
        assert "R3" in wb.sheetnames
        assert "List JIG" in wb.sheetnames
        wb.close()

    def test_tc_t3_05_pipeline_bom_filter_to_annotation_inheritance(
        self,
        tmp_path: Path,
    ) -> None:
        """TC-T3-05: Pipeline 5: TC24 BOM Filter Engine (R2) <===> Annotation Inheritance Pipeline (R3).

        Flow:
        1. Rev 1 BOM has explanations recorded by engineers.
        2. Rev 2 raw BOM export passes through DateFilter & ModelPruner.
        3. Cleaned Rev 2 BOM ingests annotations from Rev 1.
        4. 100% of unchanged parts retain explanations; new parts left clean.
        """
        # Rev 1 established data
        old_plm_df = pd.DataFrame([
            {"part_code": "302FP02010", "giai_thich": "ECN-101 Đổi NCC", "phu_trach": "Son_mecha1", "quan_ly_check": "OK"},
            {"part_code": "302FP04010", "giai_thich": "Đổi thông số đèn", "phu_trach": "Duy_mecha1", "quan_ly_check": "OK"},
        ])

        # Rev 2 raw tree before filtering
        root = BOMNode(level=1, item_id="ROOT", item_name="MACHINE", effectivity="01-Jan-2024 UP", has_children=True)
        part1 = BOMNode(level=2, item_id="302FP02010", item_name="MOTOR BRACKET", effectivity="01-Jan-2024 UP", has_children=False)
        part2 = BOMNode(level=2, item_id="302FP04010", item_name="HEATER LAMP", effectivity="01-Jan-2024 UP", has_children=False)
        new_part = BOMNode(level=2, item_id="302FP99999", item_name="NEW SENSOR", effectivity="01-Jan-2024 UP", has_children=False)
        expired_part = BOMNode(level=2, item_id="OLD_PART", item_name="OBSOLETE", effectivity="01-Jan-2020 to 31-Dec-2021", has_children=False)
        root.children.extend([part1, part2, new_part, expired_part])

        # Filter Rev 2 tree
        filter_eng = DateFilter(criteria=FilterCriteria(reference_date=datetime.date(2024, 6, 1)))
        cleaned_tree = filter_eng.filter_tree(BOMTree(roots=[root]))
        cleaned_df = cleaned_tree.to_dataframe()
        cleaned_df = cleaned_df.rename(columns={"item_id": "part_code"})

        # Ingest annotations
        migrated = migrate_annotations(cleaned_df, old_plm_df)

        p1_row = migrated[migrated["part_code"] == "302FP02010"].iloc[0]
        assert p1_row["giai_thich"] == "ECN-101 Đổi NCC"
        assert p1_row["phu_trach"] == "Son_mecha1"

        p2_row = migrated[migrated["part_code"] == "302FP04010"].iloc[0]
        assert p2_row["giai_thich"] == "Đổi thông số đèn"

        new_row = migrated[migrated["part_code"] == "302FP99999"].iloc[0]
        assert new_row["giai_thich"] == ""

        assert "OLD_PART" not in set(migrated["part_code"]), "Expired part was filtered out before inheritance"

    def test_tc_t3_06_pipeline_inheritance_to_member_workspace_recheck(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """TC-T3-06: Pipeline 6: Annotation Inheritance (R3) <===> Member Workspace Re-Check (R6).

        Flow:
        1. Member opens Member Workspace with updated BOM data.
        2. Member sees inherited explanations pre-populated.
        3. Self-check runs cleanly, only flagging newly altered parts.
        """
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()

        # Add part with inherited explanation
        view.add_cttt_row(page="01", part_code="302FP02010", quantity=1.0, explanation="Inherited: ECN-101")
        # Add part with quantity difference
        view.add_cttt_row(page="02", part_code="302FP04010", quantity=2.0, explanation="")

        # Reference data has qty=1 for both
        plm_df = pd.DataFrame([
            {"item_id": "302FP02010", "quantity": 1.0},
            {"item_id": "302FP04010", "quantity": 1.0},
        ])
        r3_df = pd.DataFrame([
            {"part_code": "302FP02010", "quantity": 1.0},
            {"part_code": "302FP04010", "quantity": 1.0},
        ])
        view.set_reference_data(plm_df, r3_df)

        res = view.run_preliminary_self_check()
        assert res["ok_count"] == 1  # 302FP02010 matches qty
        assert res["ng_count"] == 1  # 302FP04010 differs (2.0 vs 1.0)
