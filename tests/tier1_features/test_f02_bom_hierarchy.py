"""Feature F2: 6-Level BOM Tree Hierarchy Isolation Tests.

Verifies:
1. Building 6-level hierarchy with parent-child stack indexing.
2. Pre-order depth-first flattening preserves sequence.
3. Maximum depth calculation and size tracking.
4. Fast item ID and name querying across tree nodes.
5. Deep copying / cloning immutability.
"""

import pytest
from src.core.models import BOMNode, BOMTree


class TestF02BOMHierarchy:
    """Test suite for Feature F2: 6-Level BOM Tree Hierarchy."""

    def test_f02_parent_child_stack_indexing(self):
        """Test 1: Construct an explicit 6-level tree hierarchy."""
        root = BOMNode(level=1, item_id="L1_ROOT", item_name="LEVEL 1")
        curr = root
        for lv in range(2, 7):
            child = BOMNode(level=lv, item_id=f"L{lv}_PART", item_name=f"LEVEL {lv}")
            curr.add_child(child)
            curr = child

        tree = BOMTree(roots=[root])
        assert tree.size() == 6
        assert tree.max_depth() == 6
        assert len(root.children) == 1
        assert root.children[0].children[0].children[0].children[0].children[0].item_id == "L6_PART"

    def test_f02_tree_flatten_roundtrip(self):
        """Test 2: Verify pre-order DFS flattening preserves traversal order."""
        root = BOMNode(level=1, item_id="ROOT", item_name="ROOT UNIT")
        c1 = BOMNode(level=2, item_id="C1", item_name="CHILD 1")
        c2 = BOMNode(level=2, item_id="C2", item_name="CHILD 2")
        gc1 = BOMNode(level=3, item_id="GC1", item_name="GRANDCHILD 1")
        
        c1.add_child(gc1)
        root.add_child(c1)
        root.add_child(c2)

        tree = BOMTree(roots=[root])
        flat_nodes = tree.flatten()
        
        assert len(flat_nodes) == 4
        assert [n.item_id for n in flat_nodes] == ["ROOT", "C1", "GC1", "C2"]

    def test_f02_max_depth_calculation(self):
        """Test 3: Accurately compute tree max depth across multiple roots."""
        r1 = BOMNode(level=1, item_id="R1", item_name="ROOT 1")
        r1.add_child(BOMNode(level=2, item_id="R1_C1", item_name="R1 CHILD"))

        r2 = BOMNode(level=1, item_id="R2", item_name="ROOT 2")
        c2 = BOMNode(level=2, item_id="R2_C1", item_name="R2 CHILD")
        gc2 = BOMNode(level=3, item_id="R2_GC1", item_name="R2 GRANDCHILD")
        c2.add_child(gc2)
        r2.add_child(c2)

        tree = BOMTree(roots=[r1, r2])
        assert tree.max_depth() == 3
        assert tree.size() == 5

    def test_f02_find_by_id_and_name(self):
        """Test 4: Fast case-insensitive search by item_id and item_name."""
        root = BOMNode(level=1, item_id="302FP93010", item_name="LSU UNIT")
        child = BOMNode(level=2, item_id="B1303060", item_name="SCREW M3X6")
        root.add_child(child)
        tree = BOMTree(roots=[root])

        # Search by id
        res1 = tree.find_by_id("b1303060")
        assert len(res1) == 1
        assert res1[0].item_name == "SCREW M3X6"

        # Search by name exact
        res2 = tree.find_by_name("LSU UNIT")
        assert len(res2) == 1

        # Search by name substring
        res3 = tree.find_by_name("screw", exact=False)
        assert len(res3) == 1
        assert res3[0].item_id == "B1303060"

    def test_f02_tree_clone_deep_copy(self):
        """Test 5: Verify tree cloning creates an independent deep copy."""
        root = BOMNode(level=1, item_id="302FP93010", item_name="LSU UNIT")
        child = BOMNode(level=2, item_id="B1303060", item_name="SCREW M3X6", quantity=4.0)
        root.add_child(child)
        tree = BOMTree(roots=[root])

        cloned = tree.clone()
        cloned.roots[0].children[0].quantity = 10.0

        # Original must not mutate
        assert tree.roots[0].children[0].quantity == 4.0
        assert cloned.roots[0].children[0].quantity == 10.0
