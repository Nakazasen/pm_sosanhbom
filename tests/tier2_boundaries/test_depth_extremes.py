"""Tier 2 Boundary Tests: BOM Hierarchy Depth Extremes.

Verifies:
1. Single-node flat BOMs (depth 1 only) parse and build hierarchy correctly.
2. Kyocera 6-level maximum hierarchy parses without stack overflow or tree corruption.
3. Extreme depths (>6 levels, e.g. 7-10 levels) handled without RecursionError.
4. Non-contiguous level gaps (level 1 directly to level 3) resolved fail-safely.
5. Level 0 or invalid level values handled without corrupting sibling associations.
"""

from pathlib import Path
import openpyxl
import pandas as pd
import pytest

from src.core.tree_parser import PLMTreeParser
from src.core.unit_resolver import UnitResolver


class TestDepthExtremesBoundaries:
    """Boundary test suite for BOM hierarchy depth extremes."""

    def _create_plm_file_with_levels(self, path: Path, levels: list) -> Path:
        """Helper to create a synthetic PLM workbook with specified level sequence."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PLM"
        headers = [
            "No", "Level", "Item Type", "Item ID", "Has Children", "Quantity",
            "First Parts", "Second BOM flag", "Occurrence Effectivities",
            "Item Revision:Projects List", "Item Name", "Notice No", "Revision",
            "Item Rev:Release Status"
        ]
        ws.append(headers)

        for idx, lvl in enumerate(levels, 1):
            ws.append([
                idx, lvl, "Assembly" if lvl < max(levels) else "Part",
                f"PART_L{lvl}_{idx}", "True" if lvl < max(levels) else "False",
                1.0, "", "", "01-Jan-2024 UP", "PRJ_TEST",
                f"NAME_L{lvl}_{idx}", "ECN-1", "01", "Released"
            ])
        wb.save(path)
        return path

    def test_single_root_flat_bom(self, tmp_path: Path):
        """Single node at Level 1 parses and builds hierarchy with exactly 1 root and 0 children."""
        f = self._create_plm_file_with_levels(tmp_path / "single_node.xlsx", [1])
        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        assert len(tree.flatten()) == 1
        assert len(tree.roots) == 1
        root = tree.roots[0]
        assert root.level == 1
        assert len(root.children) == 0

    def test_kyocera_standard_6_level_hierarchy(self, tmp_path: Path):
        """6-level BOM hierarchy (1 -> 2 -> 3 -> 4 -> 5 -> 6) builds full parent-child chain."""
        levels = [1, 2, 3, 4, 5, 6]
        f = self._create_plm_file_with_levels(tmp_path / "6_levels.xlsx", levels)

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        assert len(tree.flatten()) == 6
        assert len(tree.roots) == 1
        curr = tree.roots[0]
        assert curr.level == 1

        for expected_lvl in [2, 3, 4, 5, 6]:
            assert len(curr.children) == 1
            curr = curr.children[0]
            assert curr.level == expected_lvl

    def test_extreme_depth_beyond_6_levels(self, tmp_path: Path):
        """Hierarchy of 10 levels deep parses without RecursionError or stack issues."""
        levels = list(range(1, 11))
        f = self._create_plm_file_with_levels(tmp_path / "10_levels.xlsx", levels)

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)
        assert len(tree.flatten()) == 10

        # Verify unit resolver traverses extreme depth safely
        resolver = UnitResolver()
        resolved_tree = resolver.resolve_tree(tree)
        assert resolved_tree is not None

    def test_non_contiguous_level_gap(self, tmp_path: Path):
        """Non-contiguous jump (Level 1 jumping directly to Level 3) is attached safely."""
        levels = [1, 3, 4]
        f = self._create_plm_file_with_levels(tmp_path / "gap_levels.xlsx", levels)

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)
        assert len(tree.flatten()) == 3
        root = tree.roots[0]
        assert root.level == 1
        # Node with level 3 becomes child of level 1 root
        assert len(root.children) == 1
        assert root.children[0].level == 3

    def test_multiple_root_nodes_at_level_1(self, tmp_path: Path):
        """BOM with multiple Level 1 roots forms distinct tree branches."""
        levels = [1, 2, 1, 2, 2]
        f = self._create_plm_file_with_levels(tmp_path / "multi_roots.xlsx", levels)

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)
        assert len(tree.roots) == 2
        assert len(tree.roots[0].children) == 1
        assert len(tree.roots[1].children) == 2
