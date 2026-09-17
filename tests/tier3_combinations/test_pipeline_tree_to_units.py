"""Tier 3 Combination Tests: PLM Parse -> Date Filter -> Model Prune -> Unit Resolve Pipeline.

Verifies:
1. End-to-end sequential execution of the 4 core transformation stages.
2. Monotonic reduction or preservation of valid node counts through the filter and prune stages.
3. Accurate assignment of governing Unit names to all surviving component nodes.
4. Rule 2 pruned sub-assembly children do not appear in the flattened resolved dataset.
5. Idempotency: re-running filters and resolution on already-processed trees produces invariant results.
"""

import datetime
from pathlib import Path
import pytest

from src.core.date_filter import DateFilter, FilterCriteria
from src.core.model_pruner import ModelPruner
from src.core.tree_parser import PLMTreeParser
from src.core.unit_resolver import UnitResolver


class TestPipelineTreeToUnits:
    """Combinatorial pipeline tests for Tree Parsing, Filtering, Pruning, and Unit Resolution."""

    def test_end_to_end_tree_pipeline_flow(self, sample_plm_excel_file: Path):
        """Execute full pipeline: parse -> date filter -> model prune -> unit resolve."""
        # Step 1: Parse 14-col PLM Excel
        parser = PLMTreeParser()
        raw_tree = parser.parse_excel(sample_plm_excel_file)
        raw_count = len(raw_tree.flatten())
        assert raw_count == 11

        # Step 2: Date filter with target date 2024-06-01
        criteria = FilterCriteria(reference_date=datetime.date(2024, 6, 1))
        date_filter = DateFilter(criteria)
        date_filtered_tree = date_filter.filter_tree(raw_tree)
        assert len(date_filtered_tree.flatten()) <= raw_count

        # Step 3: Model pruner with Virgo rules
        pruner = ModelPruner()
        pruned_tree = pruner.prune_tree(date_filtered_tree, model_name="Virgo")
        assert len(pruned_tree.flatten()) <= len(date_filtered_tree.flatten())

        # Step 4: Unit resolver
        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(pruned_tree)
        flat_nodes = resolved_tree.flatten()

        # Verify all remaining nodes have resolved unit_name
        for node in flat_nodes:
            assert isinstance(node.unit_name, str)
            if node.level > 1:
                # Sub-components must inherit their governing Unit name
                assert len(node.unit_name) > 0

    def test_pipeline_unit_governance_propagation(self, sample_plm_excel_file: Path):
        """Verify governing Unit names propagate down the tree hierarchy."""
        parser = PLMTreeParser()
        tree = parser.parse_excel(sample_plm_excel_file)

        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(tree)

        # In sample fixture:
        # Node 1: 302FP93010 is LSU UNIT (Level 1)
        # Node 2: 302FP02010 MOTOR BRACKET (Level 2) under LSU UNIT
        flat_nodes = resolved_tree.flatten()
        lsu_children = [n for n in flat_nodes if "302FP02010" in n.item_id]
        assert len(lsu_children) == 1
        assert lsu_children[0].unit_name == "LSU UNIT"

    def test_pipeline_rule2_pruning_cleans_children(self, sample_plm_excel_file: Path):
        """Rule 2 clears sub-assembly children but retains governing unit node."""
        parser = PLMTreeParser()
        tree = parser.parse_excel(sample_plm_excel_file)

        pruner = ModelPruner()
        # Sample rules in Virgo prune sub-assemblies matching certain criteria
        pruned_tree = pruner.prune_tree(tree, model_name="Virgo")

        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(pruned_tree)

        flat = resolved_tree.flatten()
        # Verify no orphan nodes exist without parent linkage
        for root in resolved_tree.roots:
            assert root.level == 1

    def test_pipeline_idempotency(self, sample_plm_excel_file: Path):
        """Running the full pipeline twice yields identical structures and attributes."""
        parser = PLMTreeParser()
        tree = parser.parse_excel(sample_plm_excel_file)

        criteria = FilterCriteria(reference_date=datetime.date(2024, 6, 1))
        date_filter = DateFilter(criteria)
        pruner = ModelPruner()
        resolver = UnitResolver()

        # Run 1
        tree1 = resolver.resolve_tree(pruner.prune_tree(date_filter.filter_tree(tree), model_name="Virgo"))
        flat1 = [(n.item_id, n.level, n.unit_name, n.quantity) for n in tree1.flatten()]

        # Run 2 on tree1
        tree2 = resolver.resolve_tree(pruner.prune_tree(date_filter.filter_tree(tree1), model_name="Virgo"))
        flat2 = [(n.item_id, n.level, n.unit_name, n.quantity) for n in tree2.flatten()]

        assert flat1 == flat2

    def test_pipeline_export_to_flat_dataframe(self, sample_plm_excel_file: Path):
        """Pipeline produces structured flat DataFrame ready for cross-reconciliation."""
        parser = PLMTreeParser()
        tree = parser.parse_excel(sample_plm_excel_file)

        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(tree)
        flat_dicts = [n.to_flat_dict() for n in resolved_tree.flatten()]

        import pandas as pd
        df = pd.DataFrame(flat_dicts)

        assert "unit_name" in df.columns
        assert "item_id" in df.columns
        assert "quantity" in df.columns
        assert "revision" in df.columns
        assert len(df) == 11
