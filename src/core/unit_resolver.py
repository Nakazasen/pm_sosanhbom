"""Rapid O(N) Unit Resolver Engine.

Replaces the 5,619-row lookup spreadsheet Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx
with a high-performance, single-pass depth-first stack traversal.

Determines the governing Unit assembly name (e.g. LSU, DEVELOPER, DRUM, FUSER,
CASSETTE, MAIN FRAME) for every sub-component and leaf part in the BOM tree.
"""

from __future__ import annotations

import logging
from typing import overload

import pandas as pd

from src.core.models import BOMNode, BOMTree

logger = logging.getLogger(__name__)


class UnitResolver:
    """Fast single-pass O(N) depth-first Unit Resolver."""

    def __init__(self, top_unit_level: int = 1) -> None:
        """Args:
            top_unit_level: The hierarchy level that defines a Unit (default 1).
                           If level 0 is present, level 1 nodes become the units.
        """
        self.top_unit_level = top_unit_level

    def resolve_tree(self, tree: BOMTree, in_place: bool = True) -> BOMTree:
        """Assign unit_name to every node in the tree hierarchy in O(N) time."""
        target_tree = tree if in_place else tree.clone()

        for root in target_tree.roots:
            if root.level == 0:
                # Level 0 is Machine body; each Level 1 child is a Unit
                root.unit_name = root.item_name
                for child in root.children:
                    unit_label = child.item_name if child.item_name else child.item_id
                    self._propagate_unit(child, unit_label)
            elif root.level <= self.top_unit_level:
                unit_label = root.item_name if root.item_name else root.item_id
                self._propagate_unit(root, unit_label)
            else:
                # Root is below standard unit level; use root's own item_name as unit
                unit_label = root.item_name if root.item_name else root.item_id
                self._propagate_unit(root, unit_label)

        return target_tree

    def _propagate_unit(self, node: BOMNode, current_unit: str, visited: set[int] | None = None) -> None:
        """Recursively propagate the governing unit name to the node and all its descendants."""
        if visited is None:
            visited = set()
        if id(node) in visited:
            return
        visited.add(id(node))
        node.unit_name = current_unit
        for child in node.children:
            self._propagate_unit(child, current_unit, visited)

    def resolve_flat_nodes(self, nodes: list[BOMNode], in_place: bool = True) -> list[BOMNode]:
        """Assign unit_name across a flat pre-order list of nodes in a single forward pass.
        
        Replicates the forward-fill behavior of legacy formula:
        Col AI = IF(Level==1, Item_Name, Col_AI_previous)
        """
        target_nodes = nodes if in_place else [n.clone() for n in nodes]
        current_unit = ""

        for node in target_nodes:
            # If at or above the top unit level (e.g. Level 1), update current_unit
            if node.level == self.top_unit_level or (not current_unit and node.level > 0):
                current_unit = node.item_name if node.item_name else node.item_id

            node.unit_name = current_unit

        return target_nodes

    def resolve_dataframe(
        self,
        df: pd.DataFrame,
        level_col: str = "level",
        name_col: str = "item_name",
        unit_col: str = "unit_name",
        in_place: bool = False,
    ) -> pd.DataFrame:
        """Resolve unit names for a pandas DataFrame representation of the BOM."""
        target_df = df if in_place else df.copy()

        if level_col not in target_df.columns or name_col not in target_df.columns:
            logger.warning("DataFrame missing required columns '%s' or '%s'", level_col, name_col)
            target_df[unit_col] = ""
            return target_df

        unit_series: list[str] = []
        current_unit = ""

        for _, row in target_df.iterrows():
            lvl = row[level_col]
            try:
                lvl_int = int(lvl)
            except (ValueError, TypeError):
                lvl_int = 1

            if lvl_int == self.top_unit_level or not current_unit:
                val = row[name_col]
                current_unit = str(val).strip() if pd.notna(val) else ""

            unit_series.append(current_unit)

        target_df[unit_col] = unit_series
        return target_df


@overload
def resolve_units(target: BOMTree, top_unit_level: int = 1, in_place: bool = True) -> BOMTree: ...


@overload
def resolve_units(target: list[BOMNode], top_unit_level: int = 1, in_place: bool = True) -> list[BOMNode]: ...


@overload
def resolve_units(target: pd.DataFrame, top_unit_level: int = 1, in_place: bool = True) -> pd.DataFrame: ...


def resolve_units(
    target: BOMTree | list[BOMNode] | pd.DataFrame,
    top_unit_level: int = 1,
    in_place: bool = True,
) -> BOMTree | list[BOMNode] | pd.DataFrame:
    """Convenience functional API to resolve unit names for trees, node lists, or DataFrames."""
    resolver = UnitResolver(top_unit_level=top_unit_level)
    if isinstance(target, BOMTree):
        return resolver.resolve_tree(target, in_place=in_place)
    elif isinstance(target, list):
        return resolver.resolve_flat_nodes(target, in_place=in_place)
    elif isinstance(target, pd.DataFrame):
        return resolver.resolve_dataframe(target, in_place=in_place)
    raise TypeError(f"Unsupported target type for resolve_units: {type(target)}")
