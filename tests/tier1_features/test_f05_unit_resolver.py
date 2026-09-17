"""Feature F5: O(N) Unit Resolver Isolation Tests.

Verifies:
1. Recursive unit name propagation from Level 1 nodes down to deepest leaves.
2. Single-pass forward fill resolution on flat node sequences.
3. DataFrame column unit resolution.
4. Level 0 machine body handling where Level 1 children become units.
5. Linear O(N) performance on large BOM trees (replacing 5,619-row spreadsheet).
"""

import time
import pandas as pd
import pytest

from src.core.models import BOMNode, BOMTree
from src.core.unit_resolver import UnitResolver


class TestF05UnitResolver:
    """Test suite for Feature F5: O(N) Unit Resolver."""

    def test_f05_hierarchical_unit_propagation(self):
        """Test 1: Propagate governing unit name from Level 1 down the hierarchy."""
        resolver = UnitResolver()

        # Unit 1: LSU
        lsu = BOMNode(level=1, item_id="302FP93010", item_name="LSU UNIT")
        motor = BOMNode(level=2, item_id="302FP94010", item_name="POLYGON MOTOR")
        screw = BOMNode(level=3, item_id="B1303060", item_name="SCREW M3X6")
        motor.add_child(screw)
        lsu.add_child(motor)

        # Unit 2: FUSER
        fuser = BOMNode(level=1, item_id="302FP93020", item_name="FUSER UNIT")
        lamp = BOMNode(level=2, item_id="302FP04010", item_name="HEATER LAMP")
        fuser.add_child(lamp)

        tree = BOMTree(roots=[lsu, fuser])
        resolved_tree = resolver.resolve_tree(tree)

        # Check LSU subtree
        assert resolved_tree.roots[0].unit_name == "LSU UNIT"
        assert resolved_tree.roots[0].children[0].unit_name == "LSU UNIT"
        assert resolved_tree.roots[0].children[0].children[0].unit_name == "LSU UNIT"

        # Check FUSER subtree
        assert resolved_tree.roots[1].unit_name == "FUSER UNIT"
        assert resolved_tree.roots[1].children[0].unit_name == "FUSER UNIT"

    def test_f05_flat_node_unit_resolution(self):
        """Test 2: Single-pass forward fill across flat pre-order node lists."""
        resolver = UnitResolver()

        nodes = [
            BOMNode(level=1, item_id="U1", item_name="UNIT DRUM"),
            BOMNode(level=2, item_id="P1", item_name="DRUM BLADE"),
            BOMNode(level=3, item_id="P2", item_name="BLADE SPRING"),
            BOMNode(level=1, item_id="U2", item_name="UNIT DEVELOPER"),
            BOMNode(level=2, item_id="P3", item_name="DEV ROLLER"),
        ]

        resolved = resolver.resolve_flat_nodes(nodes)

        assert [n.unit_name for n in resolved] == [
            "UNIT DRUM", "UNIT DRUM", "UNIT DRUM",
            "UNIT DEVELOPER", "UNIT DEVELOPER"
        ]

    def test_f05_dataframe_unit_resolution(self):
        """Test 3: Assign unit_name column in a pandas DataFrame."""
        resolver = UnitResolver()

        df = pd.DataFrame([
            {"level": 1, "item_name": "FRAME UNIT", "item_id": "U100"},
            {"level": 2, "item_name": "SUB BRACKET", "item_id": "P101"},
            {"level": 1, "item_name": "CASSETTE UNIT", "item_id": "U200"},
            {"level": 2, "item_name": "PICKUP ROLLER", "item_id": "P201"},
        ])

        res_df = resolver.resolve_dataframe(df)
        assert list(res_df["unit_name"]) == [
            "FRAME UNIT", "FRAME UNIT",
            "CASSETTE UNIT", "CASSETTE UNIT"
        ]

    def test_f05_level0_machine_body_handling(self):
        """Test 4: Handle Level 0 root where Level 1 immediate children are units."""
        resolver = UnitResolver()

        root = BOMNode(level=0, item_id="110C103NL0", item_name="MA4500ifx HONTAI")
        u1 = BOMNode(level=1, item_id="U1", item_name="SCANNER UNIT")
        u1.add_child(BOMNode(level=2, item_id="P1", item_name="CCD SENSOR"))
        u2 = BOMNode(level=1, item_id="U2", item_name="FEED UNIT")
        u2.add_child(BOMNode(level=2, item_id="P2", item_name="FEED MOTOR"))

        root.add_child(u1)
        root.add_child(u2)

        tree = BOMTree(roots=[root])
        resolved = resolver.resolve_tree(tree)

        # Level 0 root gets machine body name, children get their own unit names
        assert resolved.roots[0].unit_name == "MA4500ifx HONTAI"
        assert resolved.roots[0].children[0].unit_name == "SCANNER UNIT"
        assert resolved.roots[0].children[0].children[0].unit_name == "SCANNER UNIT"
        assert resolved.roots[0].children[1].unit_name == "FEED UNIT"
        assert resolved.roots[0].children[1].children[0].unit_name == "FEED UNIT"

    def test_f05_performance_linear_scale(self):
        """Test 5: Demonstrate O(N) linear time scaling on 10,000 nodes in under 200ms."""
        resolver = UnitResolver()
        nodes = []

        # Generate 100 units each with 100 components = 10,000 nodes
        for u in range(100):
            nodes.append(BOMNode(level=1, item_id=f"UNIT_{u}", item_name=f"UNIT_{u}"))
            for c in range(99):
                nodes.append(BOMNode(level=2, item_id=f"PART_{u}_{c}", item_name=f"PART_{u}_{c}"))

        t0 = time.perf_counter()
        resolved = resolver.resolve_flat_nodes(nodes)
        elapsed = time.perf_counter() - t0

        assert len(resolved) == 10000
        assert resolved[9999].unit_name == "UNIT_99"
        # O(N) traversal must complete within 0.2 seconds
        assert elapsed < 0.2
