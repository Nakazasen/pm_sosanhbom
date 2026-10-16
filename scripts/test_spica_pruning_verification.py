import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.core.bom_filter_manager import BOMFilterManager
from src.core.model_pruner import ModelPruner
from src.core.tree_parser import BOMTreeParser

def verify_spica_pruning():
    print("Testing Spica pruning across multiple PLM files...")
    mgr = BOMFilterManager()
    rules = mgr.get_model_rules_for_pruner("Spica")
    print(f"Loaded {len(rules)} Spica ModelRules from DB.")
    
    # Check 3V2YJ01090 rule
    pwb_mr = next((r for r in rules if r.part_code == "3V2YJ01090"), None)
    assert pwb_mr is not None, "Rule for 3V2YJ01090 not found!"
    assert pwb_mr.action == "prune_node", f"Expected prune_node, got {pwb_mr.action}"
    assert pwb_mr.parent_part_code is None, f"Expected None branch, got {pwb_mr.parent_part_code}"

    parser = BOMTreeParser()
    pruner = ModelPruner(filter_manager=mgr)

    test_files = [
        "20261002/PLM_T10C1T3NL0.xlsx", # EU 35ppm (original branch source)
        "20261002/PLM_T10C0L2US0.xlsx", # US variant
        "20261002/PLM_T10C393NL0.xlsx", # NL variant
        "20261002/PLM_T10C1Y9JP0.xlsx", # JP variant
    ]

    for file_rel in test_files:
        p = Path(file_rel)
        if not p.exists():
            print(f"Skipping {file_rel} (not found)")
            continue

        print(f"\n--- Checking file: {p.name} ---")
        raw_tree = parser.parse_file(p)
        
        # Flatten all nodes before pruning
        raw_nodes = []
        def collect(n):
            raw_nodes.append(n)
            for c in n.children:
                collect(c)
        for r in raw_tree.roots:
            collect(r)

        raw_ids = [n.item_id for n in raw_nodes]
        print(f"Raw tree total nodes: {len(raw_nodes)}")
        assert "3V2YJ01090" in raw_ids, f"3V2YJ01090 not present in raw tree {p.name}!"

        # Prune tree
        pruned_tree = pruner.prune_tree(raw_tree, model_name="Spica")
        
        pruned_nodes = []
        def collect_pruned(n):
            pruned_nodes.append(n)
            for c in n.children:
                collect_pruned(c)
        for r in pruned_tree.roots:
            collect_pruned(r)

        pruned_ids = [n.item_id for n in pruned_nodes]
        print(f"Pruned tree total nodes: {len(pruned_nodes)} (reduction: {len(raw_nodes) - len(pruned_nodes)} nodes)")

        # Verify 3V2YJ01090 is completely gone
        assert "3V2YJ01090" not in pruned_ids, f"FAIL: 3V2YJ01090 still present in pruned tree of {p.name}!"
        print(f"PASS: 3V2YJ01090 was successfully pruned (prune_node) from {p.name}!")

        # Verify children of 3V2YJ01090 are also gone
        # E.g., 3V2YJPX590, 3V2YJ47090
        for child_code in ["3V2YJPX590", "3V2YJ47090", "7SKQ1106002+H01", "7YCG02M1096+H01", "7PA1239ASW+AH01"]:
            if child_code in raw_ids:
                assert child_code not in pruned_ids, f"FAIL: Child {child_code} of 3V2YJ01090 still in {p.name}!"
        print(f"PASS: All children of 3V2YJ01090 were successfully eliminated from {p.name}!")

        # Check Level 2 rules (e.g. 31C1T04120 COVER LEFT HIGH M)
        # It had prune_children
        if "31C1T04120" in raw_ids:
            # The node itself should survive, but its children should be cleared
            matching_pruned = [n for n in pruned_nodes if n.item_id == "31C1T04120"]
            assert len(matching_pruned) > 0, f"31C1T04120 was deleted in {p.name}, should be kept!"
            for m in matching_pruned:
                assert len(m.children) == 0, f"31C1T04120 still has children in {p.name}!"
            print(f"PASS: Level 2 rule 31C1T04120 (prune_children) correctly applied in {p.name}!")

    print("\n[ALL PLM PRUNING VERIFICATIONS PASSED SUCCESSFULLY!]")

if __name__ == "__main__":
    verify_spica_pruning()
