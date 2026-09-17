"""Core Data Models for BOM Tree and Filtering Engine.

Defines Pydantic V2 models for:
- BOMNode: Hierarchical node in the 6-level BOM tree.
- BOMTree: Complete in-memory BOM tree structure.
- FilterCriteria: Options for effectivity date filtering.
- ModelRule & MatchMode: Machine model decomposition filter rules.
"""

from __future__ import annotations

import datetime
from enum import Enum
from typing import Any

import pandas as pd
from pydantic import BaseModel, ConfigDict, Field


class MatchMode(str, Enum):
    """Matching strategy for model decomposition rules."""
    FULL_NAME = "Full_name"
    PART_NAME = "Part_name"


class PruneAction(str, Enum):
    """Action to execute when a node matches a model decomposition rule."""
    DELETE_NODE = "DELETE_NODE"          # Rule 1 (Lv 6 with children) or Rule 4 (no children)
    DELETE_CHILDREN = "DELETE_CHILDREN"  # Rule 2 (has children, Lv < next -> prune sub-assembly)
    KEEP = "KEEP"                        # Rule 3 (has children, leaf/already unexpanded, Lv < 6)


class BOMNode(BaseModel):
    """A node in the hierarchical 6-level BOM tree."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    level: int = Field(ge=0, le=20, description="Hierarchy level (1..6 typically; up to 20 for deep nested hierarchies; 0 for root machine body)")
    item_id: str = Field(default="", description="Kyocera Part Code or Unit Code (e.g. 302FP20250)")
    item_name: str = Field(default="", description="Item description (e.g. PWB MAIN ASSY)")
    has_children: bool = Field(default=False, description="Whether node has child components")
    quantity: float = Field(default=1.0, description="Component quantity")
    effectivity: str = Field(default="", description="Occurrence effectivity validity string")
    revision: str = Field(default="", description="Engineering revision (e.g. 'A', '01', '-')")
    unit_name: str = Field(default="", description="Resolved governing Unit name")
    children: list[BOMNode] = Field(default_factory=list, description="Child nodes in the hierarchy")

    # Additional metadata fields from PLM exports
    item_type: str = Field(default="", description="PLM item type (e.g. Part, Assembly)")
    notice_no: str = Field(default="", description="ECN design change notice number")
    first_parts: str = Field(default="", description="1st parts indicator")
    second_bom_flag: str = Field(default="", description="2nd BOM flag")
    item_rev_status: str = Field(default="", description="Item revision status")
    row_index: int | None = Field(default=None, description="Original Excel row index")

    def is_leaf(self) -> bool:
        """Check if node is a leaf (no children attached)."""
        return len(self.children) == 0

    def add_child(self, child: BOMNode) -> BOMNode:
        """Attach a child node to this node and ensure has_children is True."""
        self.children.append(child)
        self.has_children = True
        return child

    def flatten(self, visited: set[int] | None = None) -> list[BOMNode]:
        """Flatten this node and all descendants in pre-order depth-first traversal."""
        if visited is None:
            visited = set()
        if id(self) in visited:
            return []
        visited.add(id(self))
        result: list[BOMNode] = [self]
        for child in self.children:
            result.extend(child.flatten(visited=visited))
        return result

    def to_flat_dict(self) -> dict[str, Any]:
        """Convert this node to a flat dictionary row (excluding recursive children)."""
        return {
            "level": self.level,
            "item_type": self.item_type,
            "item_id": self.item_id,
            "has_children": "True" if self.has_children else "False",
            "quantity": self.quantity,
            "first_parts": self.first_parts,
            "second_bom_flag": self.second_bom_flag,
            "effectivity": self.effectivity,
            "item_name": self.item_name,
            "notice_no": self.notice_no,
            "revision": self.revision,
            "item_rev_status": self.item_rev_status,
            "unit_name": self.unit_name,
            "row_index": self.row_index,
        }

    def clone(self, visited: dict[int, BOMNode] | None = None) -> BOMNode:
        """Deep copy this node and all children with cycle detection."""
        if visited is None:
            visited = {}
        if id(self) in visited:
            return visited[id(self)]

        new_node = BOMNode(
            level=self.level,
            item_id=self.item_id,
            item_name=self.item_name,
            has_children=self.has_children,
            quantity=self.quantity,
            effectivity=self.effectivity,
            revision=self.revision,
            unit_name=self.unit_name,
            children=[],
            item_type=self.item_type,
            notice_no=self.notice_no,
            first_parts=self.first_parts,
            second_bom_flag=self.second_bom_flag,
            item_rev_status=self.item_rev_status,
            row_index=self.row_index,
        )
        visited[id(self)] = new_node
        new_node.children = [child.clone(visited=visited) for child in self.children]
        return new_node


class BOMTree(BaseModel):
    """Container representing a complete multi-level BOM hierarchy."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    roots: list[BOMNode] = Field(default_factory=list, description="Top-level nodes (Level 0 or 1)")
    source_file: str | None = Field(default=None, description="Path to source Excel file")
    model_name: str | None = Field(default=None, description="Machine model name if filtered")
    metadata: dict[str, Any] = Field(default_factory=dict, description="Arbitrary metadata")

    def __init__(self, **data: Any) -> None:
        if "root" in data and "roots" not in data:
            root_val = data.pop("root")
            data["roots"] = [root_val] if root_val is not None else []
        super().__init__(**data)

    @property
    def root(self) -> BOMNode | None:
        """Convenience accessor for primary root node."""
        return self.roots[0] if self.roots else None

    def flatten(self) -> list[BOMNode]:
        """Flatten the entire tree into a pre-order depth-first list of nodes."""
        result: list[BOMNode] = []
        for root in self.roots:
            result.extend(root.flatten())
        return result

    def to_records(self) -> list[dict[str, Any]]:
        """Convert all nodes in the tree to a list of flat record dictionaries."""
        return [node.to_flat_dict() for node in self.flatten()]

    def to_dataframe(self) -> pd.DataFrame:
        """Convert tree nodes to a pandas DataFrame."""
        records = self.to_records()
        if not records:
            return pd.DataFrame(columns=[
                "level", "item_type", "item_id", "has_children", "quantity",
                "first_parts", "second_bom_flag", "effectivity", "item_name",
                "notice_no", "revision", "item_rev_status", "unit_name", "row_index"
            ])
        return pd.DataFrame(records)

    def size(self) -> int:
        """Return total number of nodes in the tree."""
        return len(self.flatten())

    def max_depth(self) -> int:
        """Return the maximum level present in the tree."""
        nodes = self.flatten()
        return max((n.level for n in nodes), default=0)

    def find_by_id(self, item_id: str) -> list[BOMNode]:
        """Find all nodes matching given item_id (case-insensitive, stripped)."""
        target = item_id.strip().lower()
        return [n for n in self.flatten() if n.item_id.strip().lower() == target]

    def find_by_name(self, item_name: str, exact: bool = True) -> list[BOMNode]:
        """Find all nodes matching given item_name."""
        target = item_name.strip().lower()
        if exact:
            return [n for n in self.flatten() if n.item_name.strip().lower() == target]
        return [n for n in self.flatten() if target in n.item_name.strip().lower()]

    def clone(self) -> BOMTree:
        """Deep copy this tree."""
        return BOMTree(
            roots=[root.clone() for root in self.roots],
            source_file=self.source_file,
            model_name=self.model_name,
            metadata=dict(self.metadata),
        )


class FilterCriteria(BaseModel):
    """Configuration for date validity and model decomposition filtering."""

    reference_date: datetime.date | None = Field(
        default=None,
        description="Target reference date for validity comparison (defaults to today)"
    )
    month_tolerance: int = Field(
        default=1,
        description="Expired month threshold when year is the same (legacy rule: chenhlechthang > 1)"
    )
    keep_up: bool = Field(
        default=True,
        description="Retain effectivity strings containing 'UP' unconditionally"
    )
    prune_empty_effectivity: bool = Field(
        default=True,
        description="Prune nodes with empty effectivity strings (legacy Pass 2A/2B)"
    )
    model_name: str | None = Field(
        default=None,
        description="Machine model name to apply decomposition pruning (e.g. 'Virgo', 'Iris2024')"
    )


class ModelRule(BaseModel):
    """Rule from BolocBom sheet defining decomposition action for a machine model."""

    item_name: str | None = Field(default=None, description="Assembly or component name")
    match_mode: str = Field(default="Full_name", description="'Full_name' or 'Part_name'")
    part_code: str | None = Field(default=None, description="Exact Kyocera part code if specified")

    def matches(self, node: BOMNode) -> bool:
        """Evaluate if the BOMNode matches this rule."""
        # 1. Exact match on part code takes precedence
        if self.part_code is not None and self.part_code.strip():
            return node.item_id.strip().lower() == self.part_code.strip().lower()

        # 2. Match by item_name
        if self.item_name is not None and self.item_name.strip():
            target_name = self.item_name.strip().lower()
            node_name = node.item_name.strip().lower()

            mode = self.match_mode.strip() if self.match_mode else "Full_name"
            if mode.lower() == "part_name":
                return target_name in node_name
            # Default to Full_name (exact match)
            return target_name == node_name

        return False

    def determine_action(self, node: BOMNode) -> PruneAction:
        """Determine which of the 4 legacy action rules applies to a matched node:
        
        Rule 1: has_children == True AND level == 6 -> DELETE_NODE
        Rule 2: has_children == True AND len(children) > 0 -> DELETE_CHILDREN (keep unit, prune internals)
        Rule 3: has_children == True AND len(children) == 0 AND level < 6 -> KEEP
        Rule 4: has_children == False -> DELETE_NODE
        """
        if node.has_children:
            if node.level == 6:
                return PruneAction.DELETE_NODE
            if len(node.children) > 0:
                return PruneAction.DELETE_CHILDREN
            return PruneAction.KEEP
        return PruneAction.DELETE_NODE
