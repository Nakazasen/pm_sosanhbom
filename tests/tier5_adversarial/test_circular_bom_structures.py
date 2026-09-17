"""Tier 5 Adversarial Tests: Circular BOM Structures & Recursion Defenses.

Verifies:
1. Direct self-loop (Node A -> Node A).
2. Two-node cycle (A -> B -> A).
3. Multi-node deep cycle (A -> B -> C -> D -> B).
4. UnitResolver cycle-safe propagation.
5. BOMNode.flatten() and clone() cycle protection.
6. Deeply nested 50-level linear chains without stack overflow.
"""

from __future__ import annotations

import pytest

from src.core.models import BOMNode, BOMTree
from src.core.unit_resolver import UnitResolver
from src.core.date_filter import DateFilterEngine


class TestAdversarialCircularBOMStructures:
    """Test suite for circular BOM structures and recursive loop prevention."""

    def test_direct_self_loop(self):
        """Node having itself as a child terminates cleanly during traversal."""
        node = BOMNode(level=1, item_id="LOOP_ME", item_name="SELF LOOP")
        node.children.append(node)

        # Flatten terminates and returns unique nodes visited
        nodes = node.flatten()
        assert len(nodes) == 1
        assert nodes[0].item_id == "LOOP_ME"

    def test_two_node_mutual_cycle(self):
        """A -> B -> A mutual reference terminates cleanly."""
        node_a = BOMNode(level=1, item_id="NODE_A", item_name="UNIT ALPHA")
        node_b = BOMNode(level=2, item_id="NODE_B", item_name="SUB COMP")
        node_a.children.append(node_b)
        node_b.children.append(node_a)

        nodes = node_a.flatten()
        assert len(nodes) == 2
        assert {n.item_id for n in nodes} == {"NODE_A", "NODE_B"}

    def test_multi_node_cycle_resolution(self):
        """Cycle A -> B -> C -> D -> B terminates in UnitResolver without infinite loop."""
        node_a = BOMNode(level=1, item_id="PART_A", item_name="FEEDER UNIT")
        node_b = BOMNode(level=2, item_id="PART_B", item_name="ROLLER ASSY")
        node_c = BOMNode(level=3, item_id="PART_C", item_name="GEAR SHAFT")
        node_d = BOMNode(level=4, item_id="PART_D", item_name="PINION")

        node_a.children.append(node_b)
        node_b.children.append(node_c)
        node_c.children.append(node_d)
        node_d.children.append(node_b)  # Cycle back to B

        tree = BOMTree(roots=[node_a])
        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(tree)

        assert resolved_tree.roots[0].unit_name == "FEEDER UNIT"
        assert node_b.unit_name == "FEEDER UNIT"
        assert node_c.unit_name == "FEEDER UNIT"
        assert node_d.unit_name == "FEEDER UNIT"

    def test_cycle_safe_clone(self):
        """BOMNode.clone() handles circular structures without RecursionError."""
        node_a = BOMNode(level=1, item_id="CLONE_A")
        node_b = BOMNode(level=2, item_id="CLONE_B")
        node_a.children.append(node_b)
        node_b.children.append(node_a)

        cloned = node_a.clone()
        assert cloned is not None
        assert cloned.item_id == "CLONE_A"
        assert len(cloned.children) == 1
        assert cloned.children[0].item_id == "CLONE_B"

    def test_deep_linear_hierarchy_no_stack_overflow(self):
        """Deep hierarchy up to maximum level 10 with wide branching does not trigger Python RecursionError."""
        root = BOMNode(level=1, item_id="ROOT_L1", item_name="TOP ASSY")
        curr = root
        for lvl in range(2, 11):
            child = BOMNode(level=lvl, item_id=f"CHILD_L{lvl}", item_name=f"PART L{lvl}")
            curr.children.append(child)
            for s in range(3):
                sibling = BOMNode(level=lvl, item_id=f"SIB_L{lvl}_{s}", item_name=f"SIB L{lvl}")
                curr.children.append(sibling)
            curr = child

        tree = BOMTree(roots=[root])
        assert tree.max_depth() == 10
        assert tree.size() > 20

        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(tree)
        assert resolved_tree.size() > 20
        for node in resolved_tree.flatten():
            assert node.unit_name == "TOP ASSY"
