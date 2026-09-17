"""Feature F3: Dual-Pass Date Validity Filter Isolation Tests.

Verifies:
1. Pruning branches where effectivity year is expired (chenhlechnam > 0).
2. Pruning branches where year is current but month expired > 1 (chenhlechthang > 1).
3. Preserving branches where month expired <= 1 in the same year.
4. Unconditionally preserving effectivities containing 'UP'.
5. Pruning empty effectivity branches (Pass 2A subtrees and Pass 2B leaves).
"""

import datetime
import pytest

from src.core.date_filter import DateFilter
from src.core.models import BOMNode, BOMTree, FilterCriteria


class TestF03DateFilter:
    """Test suite for Feature F3: Dual-Pass Date Validity Filter."""

    def test_f03_prune_expired_year(self):
        """Test 1: Prune branch with effectivity expired in an earlier year."""
        ref_date = datetime.date(2026, 9, 17)
        filter_engine = DateFilter(FilterCriteria(reference_date=ref_date))

        # Root valid, child expired in 2024
        root = BOMNode(level=1, item_id="UNIT1", item_name="UNIT 1", effectivity="01-Jan-2026 UP")
        child = BOMNode(level=2, item_id="OLD_PART", item_name="OLD PART", effectivity="01-Jan-2023 to 31-Dec-2024")
        root.add_child(child)

        tree = BOMTree(roots=[root])
        filtered_tree = filter_engine.filter_tree(tree)

        assert filtered_tree.size() == 1
        assert filtered_tree.roots[0].item_id == "UNIT1"
        assert len(filtered_tree.roots[0].children) == 0

    def test_f03_prune_same_year_month_diff_gt_1(self):
        """Test 2: Prune branch in same year when expired by more than 1 month (chenhlechthang > 1)."""
        ref_date = datetime.date(2026, 9, 17)  # Month 9
        filter_engine = DateFilter(FilterCriteria(reference_date=ref_date))

        # Expired in month 6 (9 - 6 = 3 > 1) -> must prune
        root = BOMNode(level=1, item_id="UNIT1", item_name="UNIT 1", effectivity="01-Jan-2026 UP")
        child = BOMNode(level=2, item_id="PART_JUNE", item_name="JUNE PART", effectivity="01-Jan-2026 to 30-Jun-2026")
        root.add_child(child)

        tree = BOMTree(roots=[root])
        filtered_tree = filter_engine.filter_tree(tree)

        assert len(filtered_tree.roots[0].children) == 0

    def test_f03_preserve_same_year_month_diff_le_1(self):
        """Test 3: Keep branch in same year when month diff <= 1 (e.g. August 2026 vs September 2026)."""
        ref_date = datetime.date(2026, 9, 17)  # Month 9
        filter_engine = DateFilter(FilterCriteria(reference_date=ref_date))

        # Expired in month 8 (9 - 8 = 1 <= 1) -> must retain according to legacy formula
        root = BOMNode(level=1, item_id="UNIT1", item_name="UNIT 1", effectivity="01-Jan-2026 UP")
        child = BOMNode(level=2, item_id="PART_AUG", item_name="AUG PART", effectivity="01-Jan-2026 to 31-Aug-2026")
        root.add_child(child)

        tree = BOMTree(roots=[root])
        filtered_tree = filter_engine.filter_tree(tree)

        assert len(filtered_tree.roots[0].children) == 1
        assert filtered_tree.roots[0].children[0].item_id == "PART_AUG"

    def test_f03_preserve_up_effectivity(self):
        """Test 4: Unconditionally preserve validity strings containing 'UP'."""
        ref_date = datetime.date(2026, 9, 17)
        filter_engine = DateFilter(FilterCriteria(reference_date=ref_date))

        root = BOMNode(level=1, item_id="ROOT", item_name="ACTIVE UNIT", effectivity="01-May-2020 UP")
        child = BOMNode(level=2, item_id="UP_PART", item_name="UP COMPONENT", effectivity="01-May-2024 UP")
        root.add_child(child)

        tree = BOMTree(roots=[root])
        filtered_tree = filter_engine.filter_tree(tree)

        assert filtered_tree.size() == 2
        assert filtered_tree.roots[0].item_id == "ROOT"
        assert filtered_tree.roots[0].children[0].item_id == "UP_PART"

    def test_f03_prune_empty_effectivity_pass2(self):
        """Test 5: Prune empty effectivities according to Pass 2A (subtree) and Pass 2B (leaf)."""
        ref_date = datetime.date(2026, 9, 17)
        filter_engine = DateFilter(FilterCriteria(reference_date=ref_date, prune_empty_effectivity=True))

        # Root has empty effectivity and has_children=False -> Pruned
        empty_leaf = BOMNode(level=1, item_id="EMPTY_LEAF", item_name="INACTIVE LEAF", has_children=False, effectivity="")
        # Root has valid effectivity, child has empty effectivity and has_children=True -> Pruned
        valid_root = BOMNode(level=1, item_id="VALID_ROOT", item_name="VALID ROOT", effectivity="01-Jan-2026 UP")
        empty_sub = BOMNode(level=2, item_id="EMPTY_SUB", item_name="EMPTY SUBTREE", has_children=True, effectivity="")
        empty_sub.add_child(BOMNode(level=3, item_id="SUB_CHILD", item_name="SUB CHILD", effectivity=""))
        valid_root.add_child(empty_sub)

        tree = BOMTree(roots=[empty_leaf, valid_root])
        filtered_tree = filter_engine.filter_tree(tree)

        assert filtered_tree.size() == 1
        assert filtered_tree.roots[0].item_id == "VALID_ROOT"
        assert len(filtered_tree.roots[0].children) == 0
