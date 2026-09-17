"""Comprehensive Unit Tests for Core BOM Tree & Rapid Unit Resolver Engine.

Covers:
- Models: BOMNode, BOMTree, FilterCriteria, ModelRule, PruneAction
- Tree Parser: 14-col and 13-col PLM parsing, stack hierarchy construction, edge cases
- Date Filter: Dual-pass effectivity filtering, 'UP' retention, elapsed year/month
- Model Pruner: The 4 action rules on matched nodes across the 6 machine models
- Unit Resolver: O(N) depth-first stack unit resolution replacing Hamtimlinhkienthuoc
- Integrated Pipeline: End-to-end parse -> date filter -> model prune -> unit resolve
"""

from __future__ import annotations

import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from src.core.date_filter import (
    DateFilter,
    extract_expiry_date,
    filter_by_date,
    is_effectivity_expired,
)
from src.core.model_pruner import (
    ModelPruner,
    prune_by_model,
)
from src.core.models import (
    BOMNode,
    BOMTree,
    ModelRule,
    PruneAction,
)
from src.core.tree_parser import (
    PLMTreeParser,
    _parse_bool,
    _parse_float_qty,
    _parse_int_level,
    parse_plm_excel,
)
from src.core.unit_resolver import (
    UnitResolver,
    resolve_units,
)

# ============================================================================
# 1. Model Tests
# ============================================================================

class TestBOMModels:
    """Tests for BOMNode, BOMTree, and ModelRule data models."""

    def test_bom_node_creation_and_defaults(self) -> None:
        node = BOMNode(level=1, item_id="302FP20250", item_name="PWB MAIN ASSY")
        assert node.level == 1
        assert node.item_id == "302FP20250"
        assert node.item_name == "PWB MAIN ASSY"
        assert node.quantity == 1.0
        assert node.has_children is False
        assert node.children == []
        assert node.is_leaf() is True

    def test_bom_node_hierarchy_and_flatten(self) -> None:
        root = BOMNode(level=1, item_id="U001", item_name="FRAME ASSY", has_children=True)
        child1 = BOMNode(level=2, item_id="P001", item_name="BRACKET")
        child2 = BOMNode(level=2, item_id="P002", item_name="MOTOR ASSY", has_children=True)
        subchild = BOMNode(level=3, item_id="P003", item_name="SCREW")

        child2.add_child(subchild)
        root.add_child(child1)
        root.add_child(child2)

        assert root.is_leaf() is False
        assert len(root.children) == 2
        assert len(child2.children) == 1
        assert subchild.is_leaf() is True

        # Flatten pre-order traversal order: root -> child1 -> child2 -> subchild
        flat = root.flatten()
        assert len(flat) == 4
        assert [n.item_id for n in flat] == ["U001", "P001", "P002", "P003"]

    def test_bom_tree_methods(self) -> None:
        root1 = BOMNode(level=1, item_id="U001", item_name="OPTICAL UNIT")
        child1 = BOMNode(level=2, item_id="P001", item_name="MIRROR")
        root1.add_child(child1)

        root2 = BOMNode(level=1, item_id="U002", item_name="FUSER UNIT")

        tree = BOMTree(roots=[root1, root2])
        assert tree.size() == 3
        assert tree.max_depth() == 2

        # Search by item ID
        found_id = tree.find_by_id("p001")
        assert len(found_id) == 1
        assert found_id[0].item_name == "MIRROR"

        # Search by item name (exact and partial)
        found_exact = tree.find_by_name("FUSER UNIT", exact=True)
        assert len(found_exact) == 1
        found_partial = tree.find_by_name("unit", exact=False)
        assert len(found_partial) == 2

        # Conversion to DataFrame
        df = tree.to_dataframe()
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 3
        assert list(df["item_id"]) == ["U001", "P001", "U002"]

    def test_model_rule_matching(self) -> None:
        # 1. Exact Part Code Match
        rule_code = ModelRule(part_code="302FP20250")
        node_match = BOMNode(level=2, item_id="302FP20250", item_name="SOMETHING ELSE")
        node_no_match = BOMNode(level=2, item_id="1102FP0000", item_name="SOMETHING ELSE")
        assert rule_code.matches(node_match) is True
        assert rule_code.matches(node_no_match) is False

        # 2. Full Name Exact Match
        rule_full = ModelRule(item_name="BOTTLE WASTE", match_mode="Full_name")
        node_full_match = BOMNode(level=1, item_id="B01", item_name="BOTTLE WASTE")
        node_partial = BOMNode(level=1, item_id="B02", item_name="BOTTLE WASTE SUB")
        assert rule_full.matches(node_full_match) is True
        assert rule_full.matches(node_partial) is False

        # 3. Part Name Substring Match
        rule_part = ModelRule(item_name="PWB MAIN", match_mode="Part_name")
        node_substring = BOMNode(level=2, item_id="E01", item_name="PWB MAIN ASSY WITH SOFTWARE")
        assert rule_part.matches(node_substring) is True

    def test_model_rule_determine_action(self) -> None:
        rule = ModelRule(item_name="TEST UNIT")

        # Rule 1: has_children=True and level=6 -> DELETE_NODE
        node_r1 = BOMNode(level=6, item_id="N1", has_children=True)
        assert rule.determine_action(node_r1) == PruneAction.DELETE_NODE

        # Rule 2: has_children=True and has children -> DELETE_CHILDREN
        node_r2 = BOMNode(level=2, item_id="N2", has_children=True)
        node_r2.add_child(BOMNode(level=3, item_id="N3"))
        assert rule.determine_action(node_r2) == PruneAction.DELETE_CHILDREN

        # Rule 3: has_children=True and no children, level < 6 -> KEEP
        node_r3 = BOMNode(level=2, item_id="N4", has_children=True, children=[])
        assert rule.determine_action(node_r3) == PruneAction.KEEP

        # Rule 4: has_children=False -> DELETE_NODE
        node_r4 = BOMNode(level=3, item_id="N5", has_children=False)
        assert rule.determine_action(node_r4) == PruneAction.DELETE_NODE


# ============================================================================
# 2. Tree Parser Tests
# ============================================================================

class TestPLMTreeParser:
    """Tests for 14-column and 13-column PLM Excel BOM parsing."""

    def test_type_converters(self) -> None:
        assert _parse_bool("True") is True
        assert _parse_bool("TRUE") is True
        assert _parse_bool(1) is True
        assert _parse_bool("False") is False
        assert _parse_bool("") is False
        assert _parse_bool(None) is False

        assert _parse_int_level("1") == 1
        assert _parse_int_level("..3") == 3
        assert _parse_int_level(2.0) == 2
        assert _parse_int_level(None) is None
        assert _parse_int_level("invalid") is None

        assert _parse_float_qty("2.5") == 2.5
        assert _parse_float_qty("3,5") == 3.5
        assert _parse_float_qty("", default=1.0) == 1.0

    def test_parse_from_dataframe_14_column(self) -> None:
        df_data = {
            "Object": [1, 2, 3, 4, 5],
            "Level": [1, 2, 3, 2, 1],
            "Item Type": ["Assembly", "Assembly", "Part", "Part", "Assembly"],
            "Item Id": ["U100", "S200", "P300", "P201", "U200"],
            "Has Children": ["True", "True", "False", "False", "False"],
            "Quantity": [1.0, 1.0, 4.0, 2.0, 1.0],
            "1st Parts": ["", "", "", "", ""],
            "2nd BOM Flag": ["", "", "", "", ""],
            "Occurrence Effectivities": ["01-Jan-2024 UP", "01-Jan-2024 UP", "01-Jan-2024 UP", "", "to 31-Dec-2023"],
            "Item Revision Projects List": ["", "", "", "", ""],
            "Item Name": ["UNIT A", "SUB-ASSY B", "SCREW M3", "SENSOR", "UNIT C"],
            "Notice No": ["", "", "", "", ""],
            "Revision": ["A", "01", "-", "B", "02"],
            "Item Rev Status": ["Released", "Released", "Released", "Draft", "Released"],
        }
        df = pd.DataFrame(df_data)
        parser = PLMTreeParser()
        tree = parser.parse_dataframe(df)

        assert tree.size() == 5
        assert len(tree.roots) == 2  # U100 and U200
        root1, root2 = tree.roots

        assert root1.item_id == "U100"
        assert len(root1.children) == 2  # S200 and P201
        sub_b = root1.children[0]
        assert sub_b.item_id == "S200"
        assert len(sub_b.children) == 1
        assert sub_b.children[0].item_id == "P300"

        assert root2.item_id == "U200"
        assert len(root2.children) == 0

    def test_parse_excel_file_end_to_end(self, tmp_path: Path) -> None:
        excel_path = tmp_path / "PLM_test_export.xlsx"

        # Create realistic Excel file with TC14 headers
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        headers = [
            "Object", "Level", "Item Type", "Item Id", "Has Children", "Quantity",
            "1st Parts", "2nd BOM Flag", "Occurrence Effectivities",
            "Item Revision Projects List", "Item Name", "Notice No", "Revision", "Item Rev Status"
        ]
        ws.append(headers)

        rows = [
            ["1", "1", "Assembly", "1102FP0000", "True", 1, "", "", "01-May-2024 UP", "", "PAPER FEED UNIT", "", "A", "Released"],
            ["2", "2", "Part", "302FP00010", "False", 2, "", "", "01-May-2024 UP", "", "ROLLER FEED", "", "01", "Released"],
            ["3", "2", "Part", "302FP00020", "False", 1, "", "", "01-May-2024 UP", "", "PULLEY PICKUP ASSY", "", "01", "Released"],
            ["4", "1", "Assembly", "1102FP1000", "True", 1, "", "", "01-May-2024 UP", "", "FUSER UNIT", "", "B", "Released"],
            ["5", "2", "Part", "302FP10010", "False", 1, "", "", "to 31-Dec-2022", "", "HEATER LAMP", "", "-", "Released"],
        ]
        for r in rows:
            ws.append(r)
        wb.save(str(excel_path))
        wb.close()

        tree = parse_plm_excel(excel_path)
        assert tree.size() == 5
        assert len(tree.roots) == 2
        assert tree.roots[0].item_name == "PAPER FEED UNIT"
        assert len(tree.roots[0].children) == 2
        assert tree.roots[1].item_name == "FUSER UNIT"
        assert len(tree.roots[1].children) == 1


# ============================================================================
# 3. Date Filter Tests
# ============================================================================

class TestDateFilter:
    """Tests for dual-pass date validity filtering with 'UP' retention."""

    def test_extract_expiry_date(self) -> None:
        assert extract_expiry_date("01-Jan-2023 to 31-Dec-2023") == datetime.date(2023, 12, 31)
        assert extract_expiry_date("to 30/06/2024") == datetime.date(2024, 6, 30)
        assert extract_expiry_date("to 2024-05-15") == datetime.date(2024, 5, 15)
        assert extract_expiry_date("01-May-2024 UP") is None
        assert extract_expiry_date("") is None
        assert extract_expiry_date("Active") is None

    def test_is_effectivity_expired_legacy_rules(self) -> None:
        ref_date = datetime.date(2024, 6, 1)

        # 1. 'UP' retention unconditionally kept
        assert is_effectivity_expired("01-Jan-2020 to 31-Dec-2020 UP", ref_date) is False
        assert is_effectivity_expired("01-May-2024 UP", ref_date) is False

        # 2. Year expired (diff_year > 0) -> True (expired)
        assert is_effectivity_expired("01-Jan-2023 to 31-Dec-2023", ref_date) is True

        # 3. Same year, diff_month > 1 (e.g. June vs April: 6 - 4 = 2 > 1) -> True (expired)
        assert is_effectivity_expired("01-Jan-2024 to 30-Apr-2024", ref_date) is True

        # 4. Same year, diff_month <= 1 (e.g. June vs May: 6 - 5 = 1 <= 1) -> False (active)
        assert is_effectivity_expired("01-Jan-2024 to 31-May-2024", ref_date) is False

        # 5. Same year, same month (diff_month = 0) -> False (active)
        assert is_effectivity_expired("01-Jan-2024 to 30-Jun-2024", ref_date) is False

        # 6. Future date in same year or next year -> False (active)
        assert is_effectivity_expired("to 31-Dec-2024", ref_date) is False
        assert is_effectivity_expired("to 31-Dec-2025", ref_date) is False

        # 7. Empty effectivity (Pass 2)
        assert is_effectivity_expired("", ref_date, prune_empty=True) is True
        assert is_effectivity_expired("", ref_date, prune_empty=False) is False

    def test_filter_tree_recursive_prune(self) -> None:
        ref_date = datetime.date(2024, 6, 1)

        # Hierarchy:
        # Root 1 (Active)
        #   - Child 1A (Active)
        #   - Child 1B (Expired year: 2023)
        #       - Subchild 1B-1 (Active string, but parent is expired)
        # Root 2 (Expired subtree: 2022)
        #   - Child 2A (Active string, but root is expired)
        # Root 3 (Empty effectivity -> pruned in Pass 2)
        root1 = BOMNode(level=1, item_id="U1", item_name="UNIT 1", effectivity="01-Jan-2024 UP")
        child_1a = BOMNode(level=2, item_id="P1A", item_name="PART 1A", effectivity="to 31-Dec-2024")
        child_1b = BOMNode(level=2, item_id="P1B", item_name="SUB 1B", effectivity="to 31-Dec-2023")
        sub_1b1 = BOMNode(level=3, item_id="P1B1", item_name="PART 1B1", effectivity="01-Jan-2024 UP")
        child_1b.add_child(sub_1b1)
        root1.add_child(child_1a)
        root1.add_child(child_1b)

        root2 = BOMNode(level=1, item_id="U2", item_name="UNIT 2", effectivity="to 31-Dec-2022")
        child_2a = BOMNode(level=2, item_id="P2A", item_name="PART 2A", effectivity="01-Jan-2024 UP")
        root2.add_child(child_2a)

        root3 = BOMNode(level=1, item_id="U3", item_name="UNIT 3", effectivity="")

        tree = BOMTree(roots=[root1, root2, root3])
        filtered = filter_by_date(tree, reference_date=ref_date)

        # Root 2 and Root 3 must be completely pruned
        assert len(filtered.roots) == 1
        surviving_root = filtered.roots[0]
        assert surviving_root.item_id == "U1"

        # Under Root 1: Child 1A survives, Child 1B and its subchild 1B-1 are pruned!
        assert len(surviving_root.children) == 1
        assert surviving_root.children[0].item_id == "P1A"


# ============================================================================
# 4. Model Pruner Tests
# ============================================================================

class TestModelPruner:
    """Tests for model decomposition pruner supporting the 4 action rules across 6 models."""

    def test_default_rules_loaded_for_all_6_models(self) -> None:
        pruner = ModelPruner()
        expected_models = ["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"]
        for m in expected_models:
            rules = pruner.get_rules_for_model(m)
            assert len(rules) > 0, f"Model {m} should have default rules"

    def test_rule_1_prune_level_6_with_children(self) -> None:
        pruner = ModelPruner()
        custom_rule = ModelRule(item_name="PHANTOM LEVEL 6", match_mode="Full_name")

        root = BOMNode(level=1, item_id="U1", item_name="UNIT")
        node_l6 = BOMNode(level=6, item_id="L6", item_name="PHANTOM LEVEL 6", has_children=True)
        root.add_child(node_l6)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, rules=[custom_rule])

        # Rule 1 prunes the node itself
        assert len(pruned.roots[0].children) == 0

    def test_rule_2_prune_children_keep_unit_node(self) -> None:
        """Rule 2: Unit assembly matched in BolocBom keeps the unit row, prunes internal components."""
        pruner = ModelPruner()
        rule = ModelRule(item_name="FRAME CONVEYING", match_mode="Full_name")

        # Root is Machine
        #   Level 1: FRAME CONVEYING (has children)
        #     Level 2: BRACKET
        #     Level 2: SCREW
        unit_node = BOMNode(level=1, item_id="FC01", item_name="FRAME CONVEYING", has_children=True)
        child1 = BOMNode(level=2, item_id="B01", item_name="BRACKET")
        child2 = BOMNode(level=2, item_id="S01", item_name="SCREW")
        unit_node.add_child(child1)
        unit_node.add_child(child2)

        tree = BOMTree(roots=[unit_node])
        assert tree.size() == 3

        pruned = pruner.prune_tree(tree, rules=[rule])

        # Unit node itself is KEPT, but its children are CLEARED!
        assert pruned.size() == 1
        assert len(pruned.roots) == 1
        assert pruned.roots[0].item_name == "FRAME CONVEYING"
        assert len(pruned.roots[0].children) == 0

    def test_rule_4_prune_leaf_node(self) -> None:
        """Rule 4: Matched node with has_children == False is pruned."""
        pruner = ModelPruner()
        rule = ModelRule(item_name="BOTTLE WASTE", match_mode="Full_name")

        root = BOMNode(level=1, item_id="U1", item_name="MAIN ASSY", has_children=True)
        leaf = BOMNode(level=2, item_id="BW01", item_name="BOTTLE WASTE", has_children=False)
        sibling = BOMNode(level=2, item_id="P02", item_name="KEEP ME", has_children=False)
        root.add_child(leaf)
        root.add_child(sibling)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, rules=[rule])

        assert len(pruned.roots[0].children) == 1
        assert pruned.roots[0].children[0].item_name == "KEEP ME"

    def test_pruning_with_real_bolocbom_model_virgo(self) -> None:
        # In BolocBom for Virgo:
        # 1. BOTTLE WASTE (Full_name)
        # 2. FILL UP CONTAINER ASSY (Part_name)
        # 3. PWB PANEL MAIN ASSY WITH SOFTWARE (Full_name)
        pruner = ModelPruner()

        root = BOMNode(level=1, item_id="U1", item_name="TOP UNIT", has_children=True)
        # Should be pruned by BOTTLE WASTE
        node1 = BOMNode(level=2, item_id="BW01", item_name="BOTTLE WASTE", has_children=False)
        # Should have children cleared by PWB PANEL MAIN ASSY WITH SOFTWARE
        node2 = BOMNode(level=2, item_id="PWB01", item_name="PWB PANEL MAIN ASSY WITH SOFTWARE", has_children=True)
        node2.add_child(BOMNode(level=3, item_id="RES01", item_name="RESISTOR"))
        # Unrelated node should remain intact with its child
        node3 = BOMNode(level=2, item_id="GEAR01", item_name="GEAR ASSY", has_children=True)
        node3.add_child(BOMNode(level=3, item_id="PIN01", item_name="PIN"))

        root.add_child(node1)
        root.add_child(node2)
        root.add_child(node3)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, model_name="Virgo")

        surviving_children = pruned.roots[0].children
        assert len(surviving_children) == 2

        # node1 pruned
        assert [c.item_name for c in surviving_children] == [
            "PWB PANEL MAIN ASSY WITH SOFTWARE",
            "GEAR ASSY",
        ]
        # node2 has children cleared
        assert len(surviving_children[0].children) == 0
        # node3 has child preserved
        assert len(surviving_children[1].children) == 1


# ============================================================================
# 5. Unit Resolver Tests
# ============================================================================

class TestUnitResolver:
    """Tests for fast O(N) depth-first Unit Resolver replacing Hamtimlinhkienthuoc."""

    def test_unit_resolution_multi_level(self) -> None:
        # Level 1: LSU
        #   Level 2: LASER DIODE
        #   Level 2: POLYGON MOTOR
        #     Level 3: BEARING
        # Level 1: FUSER
        #   Level 2: HEAT ROLLER
        #     Level 3: BUSHING
        u_lsu = BOMNode(level=1, item_id="U1", item_name="LSU", has_children=True)
        u_lsu.add_child(BOMNode(level=2, item_id="P1", item_name="LASER DIODE"))
        poly = BOMNode(level=2, item_id="P2", item_name="POLYGON MOTOR", has_children=True)
        poly.add_child(BOMNode(level=3, item_id="P3", item_name="BEARING"))
        u_lsu.add_child(poly)

        u_fuser = BOMNode(level=1, item_id="U2", item_name="FUSER", has_children=True)
        roller = BOMNode(level=2, item_id="P4", item_name="HEAT ROLLER", has_children=True)
        roller.add_child(BOMNode(level=3, item_id="P5", item_name="BUSHING"))
        u_fuser.add_child(roller)

        tree = BOMTree(roots=[u_lsu, u_fuser])
        resolver = UnitResolver()
        resolver.resolve_tree(tree)

        flat = tree.flatten()
        assert len(flat) == 7

        # Check LSU subtree (4 nodes)
        assert flat[0].unit_name == "LSU"
        assert flat[1].unit_name == "LSU"
        assert flat[2].unit_name == "LSU"
        assert flat[3].unit_name == "LSU"

        # Check FUSER subtree (3 nodes)
        assert flat[4].unit_name == "FUSER"
        assert flat[5].unit_name == "FUSER"
        assert flat[6].unit_name == "FUSER"

    def test_unit_resolution_dataframe(self) -> None:
        df = pd.DataFrame({
            "level": [1, 2, 3, 1, 2],
            "item_name": ["DEVELOPER", "MAG ROLLER", "GEAR", "DRUM", "CLEANING BLADE"],
        })
        resolver = UnitResolver()
        res_df = resolver.resolve_dataframe(df)

        assert list(res_df["unit_name"]) == [
            "DEVELOPER", "DEVELOPER", "DEVELOPER", "DRUM", "DRUM"
        ]

    def test_unit_resolver_performance_10k_nodes(self) -> None:
        """Performance benchmark: 10,000 nodes resolved in under 100ms."""
        import time

        roots: list[BOMNode] = []
        for u in range(100):
            unit_node = BOMNode(level=1, item_id=f"UNIT_{u}", item_name=f"UNIT_ASSY_{u}", has_children=True)
            for c in range(99):
                unit_node.add_child(BOMNode(level=2, item_id=f"PART_{u}_{c}", item_name=f"SUB_PART_{u}_{c}"))
            roots.append(unit_node)

        tree = BOMTree(roots=roots)
        assert tree.size() == 10000

        resolver = UnitResolver()
        start = time.perf_counter()
        resolver.resolve_tree(tree)
        duration = time.perf_counter() - start

        assert duration < 0.2, f"Expected < 200ms for 10,000 nodes, took {duration*1000:.2f}ms"
        assert tree.roots[99].children[98].unit_name == "UNIT_ASSY_99"


# ============================================================================
# 6. Integrated Pipeline Test
# ============================================================================

class TestEndToEndPipeline:
    """Integrated test verifying the complete chain:
    Parse -> Date Filter -> Model Prune -> Unit Resolve -> Flatten.
    """

    def test_full_pipeline_workflow(self, tmp_path: Path) -> None:
        excel_file = tmp_path / "PLM_integrated.xlsx"
        ref_date = datetime.date(2024, 6, 1)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([
            "Object", "Level", "Item Type", "Item Id", "Has Children", "Quantity",
            "1st Parts", "2nd BOM Flag", "Occurrence Effectivities",
            "Item Revision Projects List", "Item Name", "Notice No", "Revision", "Item Rev Status"
        ])

        # Data for machine model 'Virgo':
        # 1. Level 1: DRIVE UNIT (Active UP)
        #    2. Level 2: MOTOR (Active UP)
        #    3. Level 2: OLD BRACKET (Expired to 31-Dec-2022) -> Should be pruned by DateFilter
        # 2. Level 1: WASTE UNIT (Active)
        #    4. Level 2: BOTTLE WASTE (Active UP, has_children=False) -> Should be pruned by ModelPruner (Virgo Rule 4)
        #    5. Level 2: PWB PANEL MAIN ASSY WITH SOFTWARE (Active, has_children=True) -> Should have children pruned by ModelPruner (Virgo Rule 2)
        #       6. Level 3: INTERNAL CHIP (Active)
        # 3. Level 1: EXPIRED SUB (Expired to 31-Jan-2023) -> Entire unit pruned by DateFilter
        #    7. Level 2: SUB COMPONENT
        data_rows = [
            ["1", "1", "Assembly", "U10", "True", 1, "", "", "01-Jan-2024 UP", "", "DRIVE UNIT", "", "A", "Released"],
            ["2", "2", "Part", "P11", "False", 1, "", "", "01-Jan-2024 UP", "", "MOTOR", "", "01", "Released"],
            ["3", "2", "Part", "P12", "False", 1, "", "", "to 31-Dec-2022", "", "OLD BRACKET", "", "01", "Released"],
            ["4", "1", "Assembly", "U20", "True", 1, "", "", "01-Jan-2024 UP", "", "WASTE UNIT", "", "B", "Released"],
            ["5", "2", "Part", "BW01", "False", 1, "", "", "01-Jan-2024 UP", "", "BOTTLE WASTE", "", "01", "Released"],
            ["6", "2", "Assembly", "PWB01", "True", 1, "", "", "01-Jan-2024 UP", "", "PWB PANEL MAIN ASSY WITH SOFTWARE", "", "01", "Released"],
            ["7", "3", "Part", "CHP01", "False", 1, "", "", "01-Jan-2024 UP", "", "INTERNAL CHIP", "", "01", "Released"],
            ["8", "1", "Assembly", "U30", "True", 1, "", "", "to 31-Jan-2023", "", "EXPIRED SUB", "", "01", "Released"],
            ["9", "2", "Part", "P31", "False", 1, "", "", "01-Jan-2024 UP", "", "SUB COMPONENT", "", "01", "Released"],
        ]
        for row in data_rows:
            ws.append(row)
        wb.save(str(excel_file))
        wb.close()

        # Step 1: Parse
        raw_tree = parse_plm_excel(excel_file)
        assert raw_tree.size() == 9

        # Step 2: Date Filter
        date_filtered_tree = filter_by_date(raw_tree, reference_date=ref_date)
        # U30 and P31 pruned; P12 pruned. Remaining: U10(P11), U20(BW01, PWB01(CHP01)) = 6 nodes
        assert date_filtered_tree.size() == 6
        assert len(date_filtered_tree.roots) == 2

        # Step 3: Model Pruner (Virgo)
        # Under Virgo rules:
        # BW01 (BOTTLE WASTE) is pruned (Rule 4)
        # PWB PANEL MAIN ASSY WITH SOFTWARE has internal child CHP01 pruned, keeps PWB node (Rule 2)
        model_pruned_tree = prune_by_model(date_filtered_tree, model_name="Virgo")
        # Remaining: U10(P11), U20(PWB01) = 4 nodes
        assert model_pruned_tree.size() == 4

        # Step 4: Unit Resolver
        resolved_tree = resolve_units(model_pruned_tree)

        # Step 5: Convert to DataFrame
        final_df = resolved_tree.to_dataframe()
        assert len(final_df) == 4

        # Verify exact unit names and item descriptions
        expected_items = ["DRIVE UNIT", "MOTOR", "WASTE UNIT", "PWB PANEL MAIN ASSY WITH SOFTWARE"]
        assert list(final_df["item_name"]) == expected_items

        expected_units = ["DRIVE UNIT", "DRIVE UNIT", "WASTE UNIT", "WASTE UNIT"]
        assert list(final_df["unit_name"]) == expected_units


# ============================================================================
# 7. Additional Branch Coverage & Edge Case Tests
# ============================================================================

class TestBranchCoverageAndEdgeCases:
    """Extra tests targeting boundary branches and coverage."""

    def test_date_filter_flat_nodes(self) -> None:
        ref_date = datetime.date(2024, 6, 1)
        # Parent (Expired to 2022) with 2 children (Active UP)
        p1 = BOMNode(level=1, item_id="P1", effectivity="to 31-Dec-2022")
        c1 = BOMNode(level=2, item_id="C1", effectivity="01-Jan-2024 UP")
        c2 = BOMNode(level=3, item_id="C2", effectivity="01-Jan-2024 UP")
        # Sibling parent (Active UP) with 1 child (Active UP)
        p2 = BOMNode(level=1, item_id="P2", effectivity="01-Jan-2024 UP")
        c3 = BOMNode(level=2, item_id="C3", effectivity="01-Jan-2024 UP")

        nodes = [p1, c1, c2, p2, c3]
        filtrator = DateFilter()
        filtered = filtrator.filter_flat_nodes(nodes, reference_date=ref_date)

        assert len(filtered) == 2
        assert [n.item_id for n in filtered] == ["P2", "C3"]

    def test_model_pruner_flat_nodes_all_rules(self) -> None:
        pruner = ModelPruner()
        # Rule 1: Level 6 with children -> pop
        r1_node = BOMNode(level=6, item_id="N1", item_name="DELETE ME 1", has_children=True)
        # Rule 2: Level 2 with children -> delete children, keep node
        r2_unit = BOMNode(level=2, item_id="U2", item_name="PWB PANEL MAIN ASSY WITH SOFTWARE", has_children=True)
        r2_child1 = BOMNode(level=3, item_id="C2_1", item_name="INTERNAL 1")
        r2_child2 = BOMNode(level=4, item_id="C2_2", item_name="INTERNAL 2")
        # Rule 3: Level 2 with has_children=True but no children -> keep
        r3_unit = BOMNode(level=2, item_id="U3", item_name="KEEP ME 3", has_children=True)
        # Rule 4: Level 2 with has_children=False -> pop
        r4_leaf = BOMNode(level=2, item_id="N4", item_name="BOTTLE WASTE", has_children=False)
        # Unmatched node
        keep_node = BOMNode(level=2, item_id="OK", item_name="UNTOUCHED", has_children=False)

        flat_nodes = [r1_node, r2_unit, r2_child1, r2_child2, r3_unit, r4_leaf, keep_node]

        custom_rules = [
            ModelRule(item_name="DELETE ME 1"),
            ModelRule(item_name="PWB PANEL MAIN ASSY WITH SOFTWARE"),
            ModelRule(item_name="KEEP ME 3"),
            ModelRule(item_name="BOTTLE WASTE"),
        ]

        pruned = pruner.prune_flat_nodes(flat_nodes, rules=custom_rules)
        pruned_ids = [n.item_id for n in pruned]
        # r1_node popped; r2_unit kept, but c2_1 and c2_2 deleted; r3_unit kept; r4_leaf popped; keep_node kept
        assert pruned_ids == ["U2", "U3", "OK"]

    def test_model_pruner_strict_unknown_model_raises(self) -> None:
        pruner = ModelPruner(strict=True)
        with pytest.raises(ValueError, match="not found in BolocBom"):
            pruner.get_rules_for_model("NON_EXISTENT_MODEL")

    def test_model_pruner_custom_registry(self) -> None:
        custom_rule = ModelRule(item_name="SPECIAL")
        pruner = ModelPruner(custom_rules={"CustomMachine": [custom_rule]})
        rules = pruner.get_rules_for_model("CustomMachine")
        assert len(rules) == 1
        assert rules[0].item_name == "SPECIAL"

    def test_unit_resolver_level_0_machine_root(self) -> None:
        # Level 0: HONTAI MACHINE BODY
        #   Level 1: LSU
        #     Level 2: DIODE
        #   Level 1: DRUM
        #     Level 2: BLADE
        hontai = BOMNode(level=0, item_id="1102FP0000", item_name="HONTAI", has_children=True)
        lsu = BOMNode(level=1, item_id="LSU01", item_name="LSU", has_children=True)
        diode = BOMNode(level=2, item_id="D01", item_name="DIODE")
        drum = BOMNode(level=1, item_id="DRUM01", item_name="DRUM", has_children=True)
        blade = BOMNode(level=2, item_id="B01", item_name="BLADE")

        lsu.add_child(diode)
        drum.add_child(blade)
        hontai.add_child(lsu)
        hontai.add_child(drum)

        tree = BOMTree(roots=[hontai])
        resolver = UnitResolver()
        resolver.resolve_tree(tree)

        assert hontai.unit_name == "HONTAI"
        assert lsu.unit_name == "LSU"
        assert diode.unit_name == "LSU"
        assert drum.unit_name == "DRUM"
        assert blade.unit_name == "DRUM"

    def test_unit_resolver_flat_nodes(self) -> None:
        nodes = [
            BOMNode(level=1, item_id="U1", item_name="FRAME"),
            BOMNode(level=2, item_id="P1", item_name="SCREW"),
            BOMNode(level=1, item_id="U2", item_name="OPTICS"),
            BOMNode(level=2, item_id="P2", item_name="LENS"),
        ]
        resolver = UnitResolver()
        resolved = resolver.resolve_flat_nodes(nodes)
        assert [n.unit_name for n in resolved] == ["FRAME", "FRAME", "OPTICS", "OPTICS"]

    def test_unit_resolver_unsupported_type(self) -> None:
        with pytest.raises(TypeError, match="Unsupported target type"):
            resolve_units(12345)  # type: ignore

    def test_tree_parser_records_and_errors(self) -> None:
        parser = PLMTreeParser()

        # Empty records returns empty tree
        empty_tree = parser.parse_records([])
        assert empty_tree.size() == 0

        # Non-existent file raises FileNotFoundError
        with pytest.raises(FileNotFoundError):
            parser.parse_excel("non_existent_file.xlsx")

        # 13-column positional fallback test
        matrix_13_cols = [
            [1, "Assy", "P13_01", "True", 1, "", "", "01-Jan-2024 UP", "", "ASSY 13", "", "A", "Rel"],
            [2, "Part", "P13_02", "False", 2, "", "", "01-Jan-2024 UP", "", "LEAF 13", "", "01", "Rel"],
        ]
        tree_13 = parser._build_tree_from_matrix(matrix_13_cols, mapping=parser._get_positional_fallback(13))
        assert tree_13.size() == 2
        assert tree_13.roots[0].item_id == "P13_01"
        assert tree_13.roots[0].children[0].item_id == "P13_02"

