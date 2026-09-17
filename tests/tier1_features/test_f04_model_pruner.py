"""Feature F4: Model Decomposition & Pruning Isolation Tests.

Verifies:
1. Rule 1: Level 6 node with has_children=True is deleted.
2. Rule 2: Matched assembly keeps the unit node itself but prunes internal children.
3. Rule 3: Node with has_children=True but no children expanded (level < 6) is kept.
4. Rule 4: Matched leaf component (has_children=False) is deleted.
5. Rule execution across machine models (Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris).
"""

import pytest

from src.core.model_pruner import ModelPruner
from src.core.models import BOMNode, BOMTree, ModelRule


class TestF04ModelPruner:
    """Test suite for Feature F4: Model Decomposition & Pruning."""

    def test_f04_rule1_level6_with_children_deleted(self):
        """Test 1: Rule 1 - Prune phantom branch where Level == 6 and has_children=True."""
        rule = ModelRule(item_name="PWB MAIN", match_mode="Part_name")
        pruner = ModelPruner()

        root = BOMNode(level=1, item_id="R1", item_name="ROOT")
        # Build chain down to level 6
        l5 = BOMNode(level=5, item_id="L5", item_name="LV5")
        l6 = BOMNode(level=6, item_id="L6", item_name="PWB MAIN ASSY", has_children=True)
        l6.add_child(BOMNode(level=7, item_id="L7", item_name="PHANTOM"))
        l5.add_child(l6)
        root.add_child(l5)

        tree = BOMTree(roots=[root])
        pruned_tree = pruner.prune_tree(tree, rules=[rule])

        # Level 6 node and its phantom child must be deleted
        assert len(pruned_tree.roots[0].children[0].children) == 0

    def test_f04_rule2_prune_internal_children_keep_unit(self):
        """Test 2: Rule 2 - Matched assembly keeps the assembly row for inventory, prunes sub-components."""
        rule = ModelRule(item_name="BOTTLE WASTE", match_mode="Full_name")
        pruner = ModelPruner()

        root = BOMNode(level=1, item_id="BW01", item_name="BOTTLE WASTE", has_children=True)
        sub1 = BOMNode(level=2, item_id="P1", item_name="INTERNAL BRACKET")
        sub2 = BOMNode(level=2, item_id="P2", item_name="INTERNAL SPONGE")
        root.add_child(sub1)
        root.add_child(sub2)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, rules=[rule])

        # BOTTLE WASTE remains, but internal children are pruned
        assert pruned.size() == 1
        assert pruned.roots[0].item_name == "BOTTLE WASTE"
        assert len(pruned.roots[0].children) == 0

    def test_f04_rule3_leaf_with_children_flag_kept(self):
        """Test 3: Rule 3 - When matched node has has_children=True but no children in tree, it is kept."""
        rule = ModelRule(item_name="UNEXPANDED ASSY", match_mode="Full_name")
        pruner = ModelPruner()

        # Node level < 6 with has_children=True and len(children) == 0
        node = BOMNode(level=2, item_id="UA01", item_name="UNEXPANDED ASSY", has_children=True)
        root = BOMNode(level=1, item_id="R1", item_name="ROOT")
        root.add_child(node)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, rules=[rule])

        assert pruned.size() == 2
        assert pruned.roots[0].children[0].item_id == "UA01"

    def test_f04_rule4_matched_leaf_deleted(self):
        """Test 4: Rule 4 - Matched leaf part (has_children=False) is deleted completely."""
        rule = ModelRule(item_name="OPTIONAL SCREW", match_mode="Full_name")
        pruner = ModelPruner()

        root = BOMNode(level=1, item_id="R1", item_name="ROOT")
        leaf1 = BOMNode(level=2, item_id="S01", item_name="OPTIONAL SCREW", has_children=False)
        leaf2 = BOMNode(level=2, item_id="S02", item_name="MANDATORY SCREW", has_children=False)
        root.add_child(leaf1)
        root.add_child(leaf2)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, rules=[rule])

        assert pruned.size() == 2
        assert len(pruned.roots[0].children) == 1
        assert pruned.roots[0].children[0].item_id == "S02"

    def test_f04_six_models_rule_execution(self):
        """Test 5: Verify model pruner initializes and executes rules across 6 machine models."""
        pruner = ModelPruner()

        for model in ["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"]:
            rules = pruner.get_rules_for_model(model)
            assert len(rules) > 0

            # Test specific part code rule in Mebius (302FP20250)
            if model == "Mebius":
                root = BOMNode(level=1, item_id="M_ROOT", item_name="MEBIUS ROOT")
                target_part = BOMNode(level=2, item_id="302FP20250", item_name="SPECIAL PART", has_children=False)
                other_part = BOMNode(level=2, item_id="OTHER123", item_name="OTHER PART", has_children=False)
                root.add_child(target_part)
                root.add_child(other_part)

                pruned = pruner.prune_tree(BOMTree(roots=[root]), model_name="Mebius")
                assert len(pruned.roots[0].children) == 1
                assert pruned.roots[0].children[0].item_id == "OTHER123"
