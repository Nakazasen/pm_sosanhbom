import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.core.bom_filter_manager import BOMFilterManager
from src.core.model_pruner import ModelPruner
from src.core.tree_parser import BOMTreeParser

def verify_all_29():
    print("Running verification across all 29 PLM files in 20261002/...")
    mgr = BOMFilterManager()
    parser = BOMTreeParser()
    pruner = ModelPruner(filter_manager=mgr)

    files = sorted(Path("20261002").glob("PLM_*.xlsx"))
    print(f"Found {len(files)} PLM files.")
    assert len(files) == 29, f"Expected 29 files, found {len(files)}"

    results = []
    pwb_found_count = 0
    pwb_pruned_count = 0

    for idx, p in enumerate(files, start=1):
        raw_tree = parser.parse_file(p)
        
        raw_ids = set()
        def collect_raw(n):
            if n.item_id:
                raw_ids.add(n.item_id.strip().upper())
            for c in n.children:
                collect_raw(c)
        for r in raw_tree.roots:
            collect_raw(r)

        had_pwb = "3V2YJ01090" in raw_ids
        if had_pwb:
            pwb_found_count += 1

        pruned_tree = pruner.prune_tree(raw_tree, model_name="Spica")
        
        pruned_nodes = []
        pruned_ids = set()
        def collect_pruned(n):
            pruned_nodes.append(n)
            if n.item_id:
                pruned_ids.add(n.item_id.strip().upper())
            for c in n.children:
                collect_pruned(c)
        for r in pruned_tree.roots:
            collect_pruned(r)

        if had_pwb:
            assert "3V2YJ01090" not in pruned_ids, f"FAIL: 3V2YJ01090 still in {p.name}!"
            pwb_pruned_count += 1

        # Check that children of 3V2YJ01090 are gone
        for child_code in ["3V2YJPX590", "3V2YJ47090", "7SKQ1106002+H01", "7YCG02M1096+H01", "7PA1239ASW+AH01"]:
            assert child_code not in pruned_ids, f"FAIL: Child {child_code} still in {p.name}!"

        raw_total = sum(1 for _ in raw_ids)
        pruned_total = len(pruned_nodes)
        reduction = raw_tree.size() - pruned_total
        results.append((p.name, raw_tree.size(), pruned_total, reduction, had_pwb))

    print(f"\nVerification summary:")
    print(f"Files inspected: {len(results)}")
    print(f"Files containing 3V2YJ01090: {pwb_found_count}")
    print(f"Files where 3V2YJ01090 was successfully pruned: {pwb_pruned_count}")
    assert pwb_found_count == pwb_pruned_count, "Not all 3V2YJ01090 occurrences were pruned!"

    print("\nDetailed file statistics (First 10):")
    for fname, r_cnt, p_cnt, red, had_pwb in results[:10]:
        print(f"  {fname}: Raw={r_cnt}, Pruned={p_cnt}, Reduction={red}, Had 3V2YJ01090={had_pwb}")

    print("\nDetailed file statistics (Remaining 19):")
    for fname, r_cnt, p_cnt, red, had_pwb in results[10:]:
        print(f"  {fname}: Raw={r_cnt}, Pruned={p_cnt}, Reduction={red}, Had 3V2YJ01090={had_pwb}")

    print("\n[SUCCESS] All 29 PLM files verified with 100% success!")

if __name__ == "__main__":
    verify_all_29()
