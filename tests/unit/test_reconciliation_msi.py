"""Comprehensive Unit Tests for BOM Cross-Reconciliation & MSI Decision Engine.

Covers:
- Feature F6: Three-way cross-reconciliation (CTTT vs PLM vs R3).
- Feature F7: Missing parts detection (reverse lookup on PLM).
- Feature F8: Cross-station total aggregation (CTTT_Total).
- Feature F9: Annotation migration engine (ham_match_index_mix).
- Feature F10: MSI & Fix Serial 9-branch decision table, lookup master, and table evaluator.
- Full ReconciliationEngine integration and ReconciliationResult contract compliance.
"""

from pathlib import Path
import pandas as pd

from src.core.models import BOMNode, BOMTree
from src.core.msi_engine import (
    FixSerialMaster,
    MSIEngine,
    MSIEvaluationResult,
    evaluate_msi_branch,
)
from src.core.reconciliation import (
    ReconciliationEngine,
    ReconciliationResult,
    aggregate_cross_station,
    detect_missing_parts,
    migrate_annotations,
    reconcile_single_row,
    reconcile_three_way,
)


# ============================================================================
# F6: Three-Way Cross-Reconciliation Tests
# ============================================================================

class TestF06ReconciliationUnit:
    """Detailed unit tests for Feature F6."""

    def test_single_row_perfect_match(self):
        """When CTTT, PLM, and R3 match in quantities and revisions, result is OK."""
        res = reconcile_single_row(
            cttt_qty=5.0,
            plm_qty=5.0,
            r3_qty=5.0,
            plm_rev="A",
            r3_rev="A",
        )
        assert res["comp_plm_qty"] == "OK"
        assert res["comp_r3_qty"] == "OK"
        assert res["comp_rev"] == "OK"
        assert res["overall_check"] == "OK"
        assert res["status"] == "OK"

    def test_single_row_plm_quantity_mismatch(self):
        """When CTTT quantity differs from PLM quantity, compare_plm is NG and overall is NG."""
        res = reconcile_single_row(
            cttt_qty=10.0,
            plm_qty=5.0,
            r3_qty=10.0,
            plm_rev="01",
            r3_rev="01",
        )
        assert res["comp_plm_qty"] == "NG"
        assert res["comp_r3_qty"] == "OK"
        assert res["comp_rev"] == "OK"
        assert res["overall_check"] == "NG"

    def test_single_row_r3_quantity_mismatch(self):
        """When CTTT quantity differs from R3 quantity, compare_r3 is NG and overall is NG."""
        res = reconcile_single_row(
            cttt_qty=2.0,
            plm_qty=2.0,
            r3_qty=1.0,
            plm_rev="B",
            r3_rev="B",
        )
        assert res["comp_plm_qty"] == "OK"
        assert res["comp_r3_qty"] == "NG"
        assert res["comp_rev"] == "OK"
        assert res["overall_check"] == "NG"

    def test_single_row_revision_mismatch(self):
        """When PLM revision differs from R3 revision, compare_rev is NG and overall is NG."""
        res = reconcile_single_row(
            cttt_qty=1.0,
            plm_qty=1.0,
            r3_qty=1.0,
            plm_rev="A",
            r3_rev="B",
        )
        assert res["comp_plm_qty"] == "OK"
        assert res["comp_r3_qty"] == "OK"
        assert res["comp_rev"] == "NG"
        assert res["overall_check"] == "NG"

    def test_single_row_zero_quantity_in_plm_or_r3(self):
        """When part exists in CTTT but has 0 qty in PLM or R3, check is NG."""
        res_zero_plm = reconcile_single_row(cttt_qty=1.0, plm_qty=0.0, r3_qty=1.0, plm_rev="A", r3_rev="A")
        assert res_zero_plm["overall_check"] == "NG"

        res_zero_r3 = reconcile_single_row(cttt_qty=1.0, plm_qty=1.0, r3_qty=0.0, plm_rev="A", r3_rev="A")
        assert res_zero_r3["overall_check"] == "NG"

    def test_single_row_float_tolerance(self):
        """Tolerates minor floating-point rounding within tolerance."""
        res = reconcile_single_row(
            cttt_qty=0.33333333,
            plm_qty=0.33333334,
            r3_qty=0.33333333,
            plm_rev="-",
            r3_rev="-",
            tolerance=1e-5,
        )
        assert res["comp_plm_qty"] == "OK"
        assert res["overall_check"] == "OK"

    def test_reconcile_three_way_table_generation(self):
        """Test full three-way table generation from CTTT, PLM BOMTree, and R3 DataFrame."""
        root = BOMNode(level=0, item_id="MACHINE_BODY", item_name="BODY")
        u1 = root.add_child(BOMNode(level=1, item_id="UNIT1", item_name="LSU"))
        u1.add_child(BOMNode(level=2, item_id="PART1", item_name="GEAR", quantity=2.0, revision="A"))
        u1.add_child(BOMNode(level=2, item_id="PART2", item_name="BRACKET", quantity=1.0, revision="01"))
        tree = BOMTree(roots=[root])

        cttt_df = pd.DataFrame([
            {
                "SUB": "LSU",
                "TRANG CTTT": "P10",
                "MÃ LINH KIỆN": "PART1",
                "TÊN LINH KIỆN": "GEAR",
                "SỐ LƯỢNG": 2.0,
                "PHỤ TRÁCH": "An",
            },
            {
                "SUB": "LSU",
                "TRANG CTTT": "P11",
                "MÃ LINH KIỆN": "PART2",
                "TÊN LINH KIỆN": "BRACKET",
                "SỐ LƯỢNG": 1.0,
                "PHỤ TRÁCH": "Binh",
            },
            {
                "SUB": "LSU",
                "TRANG CTTT": "P12",
                "MÃ LINH KIỆN": "PART_EXTRA",
                "TÊN LINH KIỆN": "NOT IN PLM",
                "SỐ LƯỢNG": 1.0,
                "PHỤ TRÁCH": "Chi",
            },
        ])

        r3_df = pd.DataFrame([
            {"part_code": "PART1", "r3_total_qty": 2.0, "revision": "A"},
            {"part_code": "PART2", "r3_total_qty": 2.0, "revision": "01"},  # Mismatch with CTTT (2.0 vs 1.0)
            {"part_code": "PART_EXTRA", "r3_total_qty": 1.0, "revision": "01"},
        ])

        result_df = reconcile_three_way(cttt_df, tree, r3_df)

        assert len(result_df) == 3
        assert "Check" in result_df.columns
        assert "Q.ty (PLM)" in result_df.columns
        assert "Qty (R3)" in result_df.columns

        # PART1: All match -> OK
        row0 = result_df.iloc[0]
        assert row0["MÃ LINH KIỆN"] == "PART1"
        assert row0["Check"] == "OK"
        assert row0["Q.ty (PLM)"] == 2.0
        assert row0["Qty (R3)"] == 2.0

        # PART2: CTTT qty = 1, R3 qty = 2 -> NG
        row1 = result_df.iloc[1]
        assert row1["MÃ LINH KIỆN"] == "PART2"
        assert row1["Compare (R3 Qty)"] == "NG"
        assert row1["Check"] == "NG"

        # PART_EXTRA: Not in PLM (PLM qty = 0) -> NG
        row2 = result_df.iloc[2]
        assert row2["MÃ LINH KIỆN"] == "PART_EXTRA"
        assert row2["Q.ty (PLM)"] == 0.0
        assert row2["Check"] == "NG"

    def test_reconcile_three_way_empty_inputs(self):
        """Empty inputs produce empty DataFrame with valid standard columns."""
        df = reconcile_three_way([], pd.DataFrame(), pd.DataFrame())
        assert df.empty
        assert "Check" in df.columns
        assert "MÃ LINH KIỆN" in df.columns


# ============================================================================
# F7: Missing Parts Detection Tests
# ============================================================================

class TestF07MissingPartsUnit:
    """Detailed unit tests for Feature F7."""

    def test_detect_omitted_engineered_parts(self):
        """Identifies engineered parts present in PLM but missing from CTTT."""
        plm_df = pd.DataFrame([
            {"item_id": "PART_A", "item_name": "GEAR", "quantity": 1.0, "unit_name": "LSU"},
            {"item_id": "PART_B", "item_name": "SPRING", "quantity": 2.0, "unit_name": "LSU"},
            {"item_id": "PART_C", "item_name": "SCREW", "quantity": 4.0, "unit_name": "FUSER"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "PART_A", "SỐ LƯỢNG": 1.0},
        ])

        missing = detect_missing_parts(plm_df, cttt_df)
        assert len(missing) == 2
        missing_ids = set(missing["item_id"])
        assert missing_ids == {"PART_B", "PART_C"}

    def test_zero_missing_when_complete(self):
        """When all PLM parts exist in CTTT, missing DataFrame is empty."""
        plm_df = pd.DataFrame([
            {"item_id": "302FP02010"},
            {"item_id": "302FP93010"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "302FP02010"},
            {"MÃ LINH KIỆN": "302FP93010"},
        ])
        missing = detect_missing_parts(plm_df, cttt_df)
        assert len(missing) == 0

    def test_leaf_filtering_option(self):
        """When only_leaves is True, assembly parent nodes with children are excluded from missing check."""
        plm_df = pd.DataFrame([
            {"item_id": "ASSY_TOP", "has_children": True, "item_name": "SUB ASSY"},
            {"item_id": "LEAF_SCREW", "has_children": False, "item_name": "M3 SCREW"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "LEAF_SCREW"},
        ])

        # If only_leaves=True, ASSY_TOP is not considered a missing production part
        missing_leaves = detect_missing_parts(plm_df, cttt_df, only_leaves=True)
        assert len(missing_leaves) == 0

        # If only_leaves=False, ASSY_TOP is surfaced
        missing_all = detect_missing_parts(plm_df, cttt_df, only_leaves=False)
        assert len(missing_all) == 1
        assert missing_all.iloc[0]["item_id"] == "ASSY_TOP"

    def test_case_and_whitespace_robustness(self):
        """Part code matching is case-insensitive and trims whitespace."""
        plm_df = pd.DataFrame([{"item_id": "  302fp02010  "}])
        cttt_df = pd.DataFrame([{"MÃ LINH KIỆN": "302FP02010"}])

        missing = detect_missing_parts(plm_df, cttt_df)
        assert len(missing) == 0


# ============================================================================
# F8: Cross-Station Total Aggregation Tests
# ============================================================================

class TestF08CrossStationUnit:
    """Detailed unit tests for Feature F8."""

    def test_aggregate_shared_screws_across_stations(self):
        """Components used across multiple sub-units sum accurately and reconcile vs R3."""
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "SCREW_M3", "SUB": "LSU", "SỐ LƯỢNG": 4.0},
            {"MÃ LINH KIỆN": "SCREW_M3", "SUB": "FUSER", "SỐ LƯỢNG": 6.0},
            {"MÃ LINH KIỆN": "SCREW_M3", "SUB": "FRAME", "SỐ LƯỢNG": 10.0},
            {"MÃ LINH KIỆN": "UNIQUE_LENS", "SUB": "LSU", "SỐ LƯỢNG": 1.0},
        ])
        r3_df = pd.DataFrame([
            {"part_code": "SCREW_M3", "r3_total_qty": 20.0},
            {"part_code": "UNIQUE_LENS", "r3_total_qty": 1.0},
        ])

        totals = aggregate_cross_station(cttt_df, r3_df)
        assert len(totals) == 2

        screw_row = totals[totals["MÃ LINH KIỆN"] == "SCREW_M3"].iloc[0]
        assert screw_row["CTTT_SUM_QTY"] == 20.0
        assert screw_row["r3_total_qty"] == 20.0
        assert screw_row["STATUS"] == "OK"

        lens_row = totals[totals["MÃ LINH KIỆN"] == "UNIQUE_LENS"].iloc[0]
        assert lens_row["CTTT_SUM_QTY"] == 1.0
        assert lens_row["STATUS"] == "OK"

    def test_cross_station_mismatch_flags_ng(self):
        """When aggregated sum differs from R3 grand total, status is NG."""
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "PART_X", "SỐ LƯỢNG": 3.0},
            {"MÃ LINH KIỆN": "PART_X", "SỐ LƯỢNG": 3.0},
        ])
        r3_df = pd.DataFrame([
            {"part_code": "PART_X", "r3_total_qty": 5.0},
        ])

        totals = aggregate_cross_station(cttt_df, r3_df)
        assert totals.iloc[0]["CTTT_SUM_QTY"] == 6.0
        assert totals.iloc[0]["r3_total_qty"] == 5.0
        assert totals.iloc[0]["STATUS"] == "NG"

    def test_part_missing_in_r3_defaults_to_zero_ng(self):
        """If a part in CTTT is completely missing in R3, R3 total defaults to 0 and flags NG."""
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "ORPHAN_PART", "SỐ LƯỢNG": 2.0},
        ])
        r3_df = pd.DataFrame([])

        totals = aggregate_cross_station(cttt_df, r3_df)
        assert len(totals) == 1
        assert totals.iloc[0]["r3_total_qty"] == 0.0
        assert totals.iloc[0]["STATUS"] == "NG"


# ============================================================================
# F9: Annotation Migration Engine Tests
# ============================================================================

class TestF09AnnotationMigrationUnit:
    """Detailed unit tests for Feature F9."""

    def test_migrate_member_annotations(self):
        """Preserves and migrates explanation, person in charge, and manager check."""
        old_df = pd.DataFrame([
            {
                "part_code": "302FP02010",
                "giai_thich": "Thay the boi ECN 123",
                "phu_trach": "Nguyen Van A",
                "quan_ly_check": "CHECKED_OK",
            },
            {
                "part_code": "302FP93010",
                "giai_thich": "Dung rieng cho ban 220V",
                "phu_trach": "Tran Van B",
                "quan_ly_check": "APPROVED",
            },
        ])

        new_df = pd.DataFrame([
            {"part_code": "302FP02010", "item_name": "BRACKET", "quantity": 1.0},
            {"part_code": "302FP93010", "item_name": "LSU UNIT", "quantity": 1.0},
            {"part_code": "302FP99999", "item_name": "NEW SENSOR", "quantity": 2.0},  # Brand new part
        ])

        migrated = migrate_annotations(new_df, old_df)

        assert len(migrated) == 3

        row0 = migrated[migrated["part_code"] == "302FP02010"].iloc[0]
        assert row0["giai_thich"] == "Thay the boi ECN 123"
        assert row0["phu_trach"] == "Nguyen Van A"
        assert row0["quan_ly_check"] == "CHECKED_OK"

        row1 = migrated[migrated["part_code"] == "302FP93010"].iloc[0]
        assert row1["giai_thich"] == "Dung rieng cho ban 220V"
        assert row1["phu_trach"] == "Tran Van B"

        # Brand new part gets clean empty strings (never 'nan')
        row2 = migrated[migrated["part_code"] == "302FP99999"].iloc[0]
        assert row2["giai_thich"] == ""
        assert row2["phu_trach"] == ""
        assert row2["quan_ly_check"] == ""

    def test_migrate_empty_old_sheet(self):
        """When previous PLM sheet has no annotations, initializes clean empty columns."""
        new_df = pd.DataFrame([{"part_code": "P1"}])
        migrated = migrate_annotations(new_df, pd.DataFrame())
        assert migrated.iloc[0]["giai_thich"] == ""


# ============================================================================
# F10: MSI & Fix Serial Decision Engine Tests (All 9 Branches)
# ============================================================================

class TestF10MSIDecisionUnit:
    """Detailed unit tests for Feature F10 covering all 9 branches."""

    def test_branch_1_missing_in_plm(self):
        """Branch 1: Unit code does not exist in PLM BOM -> NG."""
        res = evaluate_msi_branch(
            in_plm=False,
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 1
        assert res["status"] == "NG"
        assert res["service_warning"] is False
        assert "not in PLM" in res["reason"]
        assert "B" in res["highlight_red"]
        assert "K" in res["highlight_red"]

    def test_branch_2_service_normalization(self):
        """Branch 2: Blank or whitespace service field is normalized to '-'."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="",  # Blank
            master_service="",
        )
        assert res["member_service_normalized"] == "-"
        assert res["status"] == "OK"

    def test_branch_3_missing_in_plm_and_code_mismatch(self):
        """Branch 3: Unit missing in PLM and code mismatch -> NG (Red on B, D, K)."""
        res = evaluate_msi_branch(
            in_plm=False,
            member_code="2NL",
            master_code="1HN",  # Mismatch
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 3
        assert res["status"] == "NG"
        assert set(res["highlight_red"]) == {"B", "D", "K"}

    def test_branch_4_missing_in_fix_serial_master(self):
        """Branch 4: In PLM, but missing from Fix Serial master tool -> NG (Red on B, K)."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="",  # Not found in master
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 4
        assert res["status"] == "NG"
        assert "Missing from Fix Serial" in res["reason"]
        assert set(res["highlight_red"]) == {"B", "K"}

    def test_branch_5_code_and_service_exact_match(self):
        """Branch 5: In PLM, Code matches, Service matches -> OK (Green on B, D, E, K)."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="APPLY GREASE",
            master_service="APPLY GREASE",
        )
        assert res["branch"] == 5
        assert res["status"] == "OK"
        assert res["service_warning"] is False
        assert set(res["highlight_green"]) == {"B", "D", "E", "K"}
        assert len(res["highlight_red"]) == 0

    def test_branch_6_code_match_neither_has_service(self):
        """Branch 6: In PLM, Code matches, member has '-', master has '' -> OK (Green on B, D, E, K)."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="1HN",
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 6
        assert res["status"] == "OK"
        assert res["service_warning"] is False
        assert set(res["highlight_green"]) == {"B", "D", "E", "K"}

    def test_branch_7_code_match_member_missed_service_note(self):
        """Branch 7: Code matches, member has '-', but master has note -> OK with service_warning (Col E Red)."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="REQUIRED LASER CALIBRATION",
        )
        assert res["branch"] == 7
        assert res["status"] == "OK"
        assert res["service_warning"] is True
        assert "Service note required" in res["reason"]
        assert set(res["highlight_green"]) == {"B", "D", "K"}
        assert res["highlight_red"] == ["E"]

    def test_branch_8_3char_code_mismatch(self):
        """Branch 8: In PLM, but 3-character code mismatch -> NG (Col D, K Red)."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="3NL",  # Code mismatch
            member_service="-",
            master_service="",
        )
        assert res["branch"] == 8
        assert res["status"] == "NG"
        assert "3-char code mismatch" in res["reason"]
        assert "D" in res["highlight_red"]
        assert "K" in res["highlight_red"]

    def test_branch_9_service_comment_mismatch(self):
        """Branch 9: In PLM, Code matches, but service comments differ -> NG (Col E, K Red)."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="COMMENT ALPHA",
            master_service="COMMENT BETA",
        )
        assert res["branch"] == 9
        assert res["status"] == "NG"
        assert "Service mismatch" in res["reason"]
        assert "E" in res["highlight_red"]
        assert "K" in res["highlight_red"]


# ============================================================================
# FixSerialMaster & MSIEngine Table Evaluation Tests
# ============================================================================

class TestFixSerialMasterAndEngine:
    """Unit tests for FixSerialMaster dictionary and MSIEngine batch evaluator."""

    def test_master_lookup_subunit_prefix_9(self):
        """Sub-unit lookup matches first 9 characters of unit code."""
        master = FixSerialMaster()
        master.add_subunit("302NL93010_LSU", "2NL", "SERVICE NOTE 1")
        master.add_subunit("302HN93020_FUSER", "1HN", "")

        # Exact prefix
        msi, srv = master.lookup_subunit("302NL93010")
        assert msi == "2NL"
        assert srv == "SERVICE NOTE 1"

        # Extended string with same 9-char prefix
        msi2, srv2 = master.lookup_subunit("302HN93020-01")
        assert msi2 == "1HN"
        assert srv2 == ""

        # Not found
        msi_none, _ = master.lookup_subunit("UNKNOWN_UNIT")
        assert msi_none == ""

    def test_master_lookup_machine_prefix_10(self):
        """Machine body (Hontai) lookup matches first 10 characters."""
        master = FixSerialMaster()
        master.add_machine("2NL_HONTAI_BODY", "2NL", "MACHINE SERVICE")

        msi, srv = master.lookup_machine("2NL_HONTAI_01")
        assert msi == "2NL"
        assert srv == "MACHINE SERVICE"

    def test_msi_engine_evaluate_table_comprehensive(self):
        """Evaluates a batch MSI DataFrame with sub-units and Hontai machine body."""
        master = FixSerialMaster()
        master.add_subunit("302NL93010", "2NL", "")
        master.add_subunit("302HN93020", "1HN", "MANDATORY CHECK")
        master.add_machine("2NL_HONTAI", "2NL", "")

        plm_parts = ["302NL93010", "302HN93020", "2NL_HONTAI"]

        msi_df = pd.DataFrame([
            # Subunit 1: Perfect match
            {"MÃ UNIT": "302NL93010", "3 KÝ TỰ MSI": "2NL", "SERVICE": "-", "TÊN UNIT": "LSU"},
            # Subunit 2: Branch 7 (service warning)
            {"MÃ UNIT": "302HN93020", "3 KÝ TỰ MSI": "1HN", "SERVICE": "-", "TÊN UNIT": "FUSER"},
            # Subunit 3: Branch 1 (not in PLM)
            {"MÃ UNIT": "NOT_IN_PLM", "3 KÝ TỰ MSI": "999", "SERVICE": "-", "TÊN UNIT": "DRUM"},
            # Row 4: Hontai with blank unit code -> defaults to hontai_default_code
            {"MÃ UNIT": "", "3 KÝ TỰ MSI": "2NL", "SERVICE": "-", "TÊN UNIT": "HONTAI BODY", "is_hontai": True},
        ])

        engine = MSIEngine(master=master)
        evaluated = engine.evaluate_table(msi_df, plm_parts, hontai_default_code="2NL_HONTAI")

        assert len(evaluated) == 4

        # Row 0: OK (Branch 6)
        assert evaluated.iloc[0]["KẾT QUẢ"] == "OK"
        assert evaluated.iloc[0]["BRANCH"] == 6

        # Row 1: OK with service_warning (Branch 7)
        assert evaluated.iloc[1]["KẾT QUẢ"] == "OK"
        assert evaluated.iloc[1]["BRANCH"] == 7
        assert bool(evaluated.iloc[1]["SERVICE_WARNING"]) is True

        # Row 2: NG (Branch 3: not in PLM and code mismatch)
        assert evaluated.iloc[2]["KẾT QUẢ"] == "NG"
        assert evaluated.iloc[2]["BRANCH"] == 3

        # Row 3: Hontai with defaulted machine code -> OK (Branch 6)
        assert evaluated.iloc[3]["MÃ UNIT"] == "2NL_HONTAI"
        assert evaluated.iloc[3]["KẾT QUẢ"] == "OK"
        assert evaluated.iloc[3]["BRANCH"] == 6


# ============================================================================
# Full Pipeline: ReconciliationEngine & ReconciliationResult
# ============================================================================

class TestReconciliationEngineFullPipeline:
    """Unit tests for the integrated ReconciliationEngine pipeline."""

    def test_run_full_reconciliation_all_pass(self):
        """When CTTT, PLM, R3, and MSI all match, overall status is 'OK'."""
        cttt = pd.DataFrame([
            {"SUB": "LSU", "TRANG CTTT": "1", "MÃ LINH KIỆN": "P1", "TÊN LINH KIỆN": "GEAR", "SỐ LƯỢNG": 1.0, "PHỤ TRÁCH": "A"},
        ])
        plm = pd.DataFrame([
            {"item_id": "P1", "quantity": 1.0, "revision": "A", "has_children": False},
        ])
        r3 = pd.DataFrame([
            {"part_code": "P1", "r3_total_qty": 1.0, "revision": "A"},
        ])

        engine = ReconciliationEngine()
        result = engine.run_full_reconciliation(cttt, plm, r3)

        assert isinstance(result, ReconciliationResult)
        assert result.overall_status == "OK"
        assert len(result.cttt_rows) == 1
        assert len(result.plm_missing_rows) == 0
        assert len(result.cttt_totals) == 1
        assert result.cttt_totals.iloc[0]["STATUS"] == "OK"

    def test_run_full_reconciliation_fails_on_missing_parts(self):
        """When PLM has parts not included in CTTT, overall status is 'NG'."""
        cttt = pd.DataFrame([
            {"SUB": "LSU", "TRANG CTTT": "1", "MÃ LINH KIỆN": "P1", "SỐ LƯỢNG": 1.0},
        ])
        plm = pd.DataFrame([
            {"item_id": "P1", "quantity": 1.0, "revision": "A"},
            {"item_id": "P2_OMITTED", "quantity": 2.0, "revision": "01"},
        ])
        r3 = pd.DataFrame([
            {"part_code": "P1", "r3_total_qty": 1.0, "revision": "A"},
        ])

        engine = ReconciliationEngine()
        result = engine.run_full_reconciliation(cttt, plm, r3)

        assert result.overall_status == "NG"
        assert len(result.plm_missing_rows) == 1
        assert result.plm_missing_rows.iloc[0]["item_id"] == "P2_OMITTED"

    def test_run_full_reconciliation_with_msi_ng(self):
        """When MSI results contain 'NG', overall status flags 'NG'."""
        cttt = pd.DataFrame([
            {"SUB": "LSU", "TRANG CTTT": "1", "MÃ LINH KIỆN": "P1", "SỐ LƯỢNG": 1.0},
        ])
        plm = pd.DataFrame([
            {"item_id": "P1", "quantity": 1.0, "revision": "A"},
        ])
        r3 = pd.DataFrame([
            {"part_code": "P1", "r3_total_qty": 1.0, "revision": "A"},
        ])
        msi_results = pd.DataFrame([
            {"MÃ UNIT": "U1", "KẾT QUẢ": "NG", "GHI CHÚ": "3-char code mismatch"},
        ])

        engine = ReconciliationEngine()
        result = engine.run_full_reconciliation(cttt, plm, r3, msi_results=msi_results)
        assert result.overall_status == "NG"

    def test_run_full_reconciliation_with_msi_ok(self):
        """When MSI results contain only 'OK', overall status is 'OK'."""
        cttt = pd.DataFrame([
            {"SUB": "LSU", "TRANG CTTT": "1", "MÃ LINH KIỆN": "P1", "SỐ LƯỢNG": 1.0},
        ])
        plm = pd.DataFrame([
            {"item_id": "P1", "quantity": 1.0, "revision": "A"},
        ])
        r3 = pd.DataFrame([
            {"part_code": "P1", "r3_total_qty": 1.0, "revision": "A"},
        ])
        msi_results = pd.DataFrame([
            {"MÃ UNIT": "U1", "KẾT QUẢ": "OK", "GHI CHÚ": "Match"},
        ])

        engine = ReconciliationEngine()
        result = engine.run_full_reconciliation(cttt, plm, r3, msi_results=msi_results)
        assert result.overall_status == "OK"

    def test_engine_delegated_methods(self):
        """Verify ReconciliationEngine instance delegation methods."""
        engine = ReconciliationEngine()
        rec = engine.reconcile_single_row(1.0, 1.0, 1.0, "A", "A")
        assert rec["overall_check"] == "OK"

        cttt = pd.DataFrame([{"MÃ LINH KIỆN": "P1", "SỐ LƯỢNG": 1.0}])
        plm = pd.DataFrame([{"item_id": "P1", "quantity": 1.0, "revision": "A"}])
        r3 = pd.DataFrame([{"part_code": "P1", "r3_total_qty": 1.0, "revision": "A"}])

        df_rec = engine.reconcile_three_way(cttt, plm, r3)
        assert len(df_rec) == 1

        missing = engine.detect_missing_parts(plm, cttt)
        assert len(missing) == 0

        totals = engine.aggregate_cross_station(cttt, r3)
        assert len(totals) == 1

        old_plm = pd.DataFrame([{"part_code": "P1", "giai_thich": "NOTE", "phu_trach": "USER", "quan_ly_check": "OK"}])
        migrated = engine.migrate_annotations(plm, old_plm)
        assert migrated.iloc[0]["giai_thich"] == "NOTE"


# ============================================================================
# Extended Edge Case & File Loading Tests
# ============================================================================

class TestExtendedEdgeCasesAndFileLoading:
    """Tests covering file loading, empty inputs, fuzzy matching, and helpers."""

    def test_fix_serial_master_load_from_excel(self, tmp_path: Path):
        """Tests loading FixSerialMaster tables from real Excel workbook (.xlsx)."""
        xlsx_path = tmp_path / "TEST_FIX_SERIAL.xlsx"

        # Sheet UNIT: Col 0: Unit Code, Col 5: MSI Code, Col 7: Service
        unit_data = [
            ["302NL93010", "", "", "", "", "2NL", "", "OIL APPLY"],
            ["302HN93020", "", "", "", "", "1HN", "", ""],
        ]
        # Sheet MACHINE: Col 0: Machine Code, Col 5: MSI Code, Col 7: Service
        machine_data = [
            ["2NL_MACHINE_BODY", "", "", "", "", "2NL", "", "MACHINE TEST"],
        ]

        df_unit = pd.DataFrame(unit_data)
        df_machine = pd.DataFrame(machine_data)

        with pd.ExcelWriter(xlsx_path) as writer:
            df_unit.to_excel(writer, sheet_name="UNIT", index=False, header=False)
            df_machine.to_excel(writer, sheet_name="MACHINE", index=False, header=False)

        master = FixSerialMaster()
        loaded = master.load_from_file(xlsx_path)
        assert loaded is True

        msi_u, srv_u = master.lookup_subunit("302NL93010")
        assert msi_u == "2NL"
        assert srv_u == "OIL APPLY"

        msi_m, srv_m = master.lookup_machine("2NL_MACHINE_BODY")
        assert msi_m == "2NL"
        assert srv_m == "MACHINE TEST"

        # Non-existent file returns False
        assert master.load_from_file(tmp_path / "non_existent.xlsx") is False

    def test_fix_serial_master_lookups_empty_and_fuzzy(self):
        """Test lookup_subunit and lookup_machine with empty strings and fuzzy substring matching."""
        master = FixSerialMaster()
        master.add_subunit("302FP9301", "2FP", "SUBUNIT_SRV")
        master.add_machine("MACHINE_2FP", "2FP", "MACHINE_SRV")

        assert master.lookup_subunit("") == ("", "")
        assert master.lookup_subunit(None) == ("", "")
        assert master.lookup_machine("") == ("", "")
        assert master.lookup_machine(None) == ("", "")

        # Fuzzy substring match
        msi, srv = master.lookup_subunit("PREFIX_302FP9301_SUFFIX")
        assert msi == "2FP"

        msi_m, srv_m = master.lookup_machine("PREFIX_MACHINE_2FP_SUFFIX")
        assert msi_m == "2FP"

    def test_msi_evaluation_result_to_dict(self):
        """Test MSIEvaluationResult.to_dict structure."""
        res = MSIEvaluationResult(
            branch=5,
            status="OK",
            service_warning=False,
            reason="Match",
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="-",
            in_plm=True,
            highlight_green=["B", "D", "E", "K"],
            highlight_red=[],
        )
        d = res.to_dict()
        assert d["branch"] == 5
        assert d["status"] == "OK"
        assert d["service_warning"] is False
        assert d["highlight_green"] == ["B", "D", "E", "K"]

    def test_msi_engine_evaluate_branch_direct(self):
        """Test MSIEngine.evaluate_branch direct method."""
        engine = MSIEngine()
        branch_dict = engine.evaluate_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="",
        )
        assert branch_dict["status"] == "OK"
        assert branch_dict["branch"] == 6

    def test_msi_engine_evaluate_table_empty_and_hontai_by_name(self):
        """Test evaluate_table with empty DataFrame and with Hontai recognized by item name."""
        master = FixSerialMaster()
        master.add_machine("HONTAI_BODY", "2NL", "")
        engine = MSIEngine(master=master)

        # Empty
        df_empty = engine.evaluate_table([], [])
        assert df_empty.empty
        assert "KẾT QUẢ" in df_empty.columns

        # Hontai recognized by row name "HONTAI" at last index
        df_hontai = pd.DataFrame([
            {"MÃ UNIT": "", "3 KÝ TỰ MSI": "2NL", "SERVICE": "-", "TÊN UNIT": "HONTAI MAIN BODY"},
        ])
        evaluated = engine.evaluate_table(df_hontai, ["HONTAI_BODY"], hontai_default_code="HONTAI_BODY")
        assert len(evaluated) == 1
        assert evaluated.iloc[0]["KẾT QUẢ"] == "OK"

    def test_reconciliation_summary_extractions_edge_cases(self):
        """Test summary extractors with lists of dicts, empty sets, and missing columns."""
        from src.core.reconciliation import (
            _clean_part_code,
            _extract_plm_summary,
            _extract_r3_summary,
            _normalize_rev,
            _to_float,
        )

        # Helper conversions
        assert _to_float(None) == 0.0
        assert _to_float("invalid") == 0.0
        assert _to_float(42.5) == 42.5
        assert _clean_part_code(None) == ""
        assert _clean_part_code("  part_xyz  ") == "PART_XYZ"
        assert _normalize_rev(None) == ""
        assert _normalize_rev("nan") == ""
        assert _normalize_rev("  A  ") == "A"

        # PLM extractor with list of dicts
        plm_list = [{"item_id": "P1", "quantity": 2.0, "revision": "A"}]
        summary_plm = _extract_plm_summary(plm_list)
        assert len(summary_plm) == 1
        assert summary_plm.iloc[0]["plm_qty"] == 2.0

        # PLM extractor with empty / invalid
        assert _extract_plm_summary(None).empty
        assert _extract_plm_summary([]).empty

        # R3 extractor with list of dicts
        r3_list = [{"part_code": "P1", "quantity": 3.0, "revision": "A"}]
        summary_r3 = _extract_r3_summary(r3_list)
        assert len(summary_r3) == 1
        assert summary_r3.iloc[0]["r3_qty"] == 3.0

        # R3 extractor with empty / invalid
        assert _extract_r3_summary(None).empty
        assert _extract_r3_summary([]).empty
