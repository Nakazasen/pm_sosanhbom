"""Dual-Pass Effectivity Date Filtering Engine.

Faithfully implements the two-pass validity filtering logic from legacy locbomfull.bas:
- Pass 1: Identifies expired effectivity strings containing "to <date>" (without "UP"),
  calculating year and month differences against a reference date (default today).
  Expired if: (year_diff > 0) OR (year_diff == 0 AND month_diff > 1).
- Pass 2: Prunes nodes with empty effectivity strings ("") according to has_children rules:
  Pass 2A: Empty effectivity on subtrees (has_children = True) prunes the subtree.
  Pass 2B: Empty effectivity on leaves (has_children = False) prunes the leaf.
- "UP" retention: Retains strings containing "UP" unconditionally.
"""

from __future__ import annotations

import datetime
import logging
import re

from src.core.models import BOMNode, BOMTree, FilterCriteria

logger = logging.getLogger(__name__)

# Common date formats in Teamcenter Active Workspace and Excel exports
DATE_FORMATS = [
    "%d-%b-%Y",  # 31-Dec-2023
    "%d-%b-%y",  # 31-Dec-23
    "%d/%m/%Y",  # 31/12/2023
    "%d/%m/%y",  # 31/12/23
    "%d-%m-%Y",  # 31-12-2023
    "%d-%m-%y",  # 31-12-23
    "%Y-%m-%d",  # 2023-12-31
    "%Y/%m/%d",  # 2023/12/31
    "%d.%m.%Y",  # 31.12.2023
]


def extract_expiry_date(effectivity: str) -> datetime.date | None:
    """Extract expiry date following 'to' in an effectivity string.
    
    Examples:
        '01-Jan-2023 to 31-Dec-2023' -> 2023-12-31
        'to 30/06/2024' -> 2024-06-30
        '01-May-2024 UP' -> None (no 'to')
    """
    if not effectivity or not isinstance(effectivity, str):
        return None

    # Search for date pattern immediately following 'to'
    match = re.search(r"\bto\s+([0-9]{1,2}[-/][A-Za-z0-9]+[-/][0-9]{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})", effectivity, re.IGNORECASE)
    if not match:
        return None

    date_str = match.group(1).strip()
    for fmt in DATE_FORMATS:
        try:
            dt = datetime.datetime.strptime(date_str, fmt).replace(tzinfo=datetime.timezone.utc)
            return dt.date()
        except ValueError:
            continue

    logger.debug("Failed to parse extracted date string '%s' from effectivity '%s'", date_str, effectivity)
    return None


def _get_current_date() -> datetime.date:
    """Return current UTC date."""
    return datetime.datetime.now(datetime.timezone.utc).date()


def is_effectivity_expired(
    effectivity: str,
    reference_date: datetime.date,
    month_tolerance: int = 1,
    keep_up: bool = True,
    prune_empty: bool = True,
) -> bool:
    """Evaluate whether an effectivity string is expired based on legacy locbomfull rules.
    
    Returns:
        True if the effectivity is expired (should be pruned).
        False if the effectivity is active/valid (should be retained).
    """
    eff = effectivity.strip() if effectivity else ""

    # Pass 2: Empty effectivity check
    if not eff:
        return prune_empty

    # Guard: Fail-closed on strings lacking digits and not 'UP' (e.g. 'GARBAGE_DATE_STRING')
    has_digits = any(c.isdigit() for c in eff)
    if not has_digits and eff.upper() != "UP":
        return True

    # If 'to' is present, validate and evaluate the expiry date
    if re.search(r"\bto\s+", eff, re.IGNORECASE):
        expiry_date = extract_expiry_date(eff)
        if expiry_date is None:
            # Malformed/unparseable date after 'to' (e.g. 'to 9999-99-99 UP')
            return True

        # Unconditionally preserve valid effectivity strings containing 'UP' per legacy rule
        if keep_up and ("up" in eff.lower()):
            return False

        diff_year = reference_date.year - expiry_date.year
        diff_month = reference_date.month - expiry_date.month
        return (diff_year > 0) or (diff_year == 0 and diff_month > month_tolerance)

    # Pass 1 Check: Retention of "UP" without 'to'
    if keep_up and ("up" in eff.lower()):
        return False

    return False


class DateFilter:
    """Filters BOM tree hierarchies based on effectivity date validity."""

    def __init__(
        self,
        criteria: FilterCriteria | None = None,
        reference_date_str: str | None = None,
    ) -> None:
        if criteria is not None:
            self.criteria = criteria
        elif reference_date_str is not None:
            try:
                ref_date = datetime.date.fromisoformat(reference_date_str)
            except Exception:
                ref_date = None
            self.criteria = FilterCriteria(reference_date=ref_date)
        else:
            self.criteria = FilterCriteria()

    def is_valid(self, effectivity: str, reference_date: datetime.date | None = None) -> bool:
        """Check if an effectivity string is valid (not expired)."""
        ref_date = reference_date or self.criteria.reference_date or _get_current_date()
        expired = is_effectivity_expired(
            effectivity=effectivity,
            reference_date=ref_date,
            month_tolerance=self.criteria.month_tolerance,
            keep_up=self.criteria.keep_up,
            prune_empty=self.criteria.prune_empty_effectivity,
        )
        return not expired

    def is_effective(self, effectivity: str, reference_date: datetime.date | None = None) -> bool:
        """Alias for is_valid."""
        return self.is_valid(effectivity, reference_date)

    def filter_tree(
        self,
        tree: BOMTree,
        reference_date: datetime.date | None = None,
        in_place: bool = False,
    ) -> BOMTree:
        """Filter the entire BOMTree in-memory, pruning expired nodes and subtrees."""
        ref_date = reference_date or self.criteria.reference_date or _get_current_date()
        target_tree = tree if in_place else tree.clone()

        filtered_roots: list[BOMNode] = []
        visited: set[int] = set()
        for root in target_tree.roots:
            filtered_node = self._filter_node_recursive(root, ref_date, visited)
            if filtered_node is not None:
                filtered_roots.append(filtered_node)

        target_tree.roots = filtered_roots
        return target_tree

    def _filter_node_recursive(
        self,
        node: BOMNode,
        reference_date: datetime.date,
        visited: set[int] | None = None,
    ) -> BOMNode | None:
        """Recursively filter a node and its children with cyclic graph protection.
        
        If node itself is expired, returns None (pruning node and all descendants).
        Otherwise, filters children recursively and returns node.
        """
        if visited is None:
            visited = set()
        if id(node) in visited:
            return node
        visited.add(id(node))

        if not self.is_valid(node.effectivity, reference_date):
            return None

        # Filter children recursively
        surviving_children: list[BOMNode] = []
        for child in node.children:
            filtered_child = self._filter_node_recursive(child, reference_date, visited=visited)
            if filtered_child is not None:
                surviving_children.append(filtered_child)

        node.children = surviving_children
        return node

    def filter_flat_nodes(
        self,
        nodes: list[BOMNode],
        reference_date: datetime.date | None = None,
    ) -> list[BOMNode]:
        """Filter a flat list of nodes, faithfully replicating sequential row deletion from Excel VBA.
        
        When an expired parent node is encountered, prunes that node and all consecutive
        descendant nodes with level > parent.level.
        """
        ref_date = reference_date or self.criteria.reference_date or _get_current_date()
        result: list[BOMNode] = []
        idx = 0
        n = len(nodes)

        while idx < n:
            node = nodes[idx]
            if not self.is_valid(node.effectivity, ref_date):
                # Expired! Prune this node and skip all descendants
                current_level = node.level
                idx += 1
                while idx < n and nodes[idx].level > current_level:
                    idx += 1
                continue
            else:
                result.append(node)
                idx += 1

        return result


def filter_by_date(
    tree: BOMTree,
    reference_date: datetime.date | None = None,
    month_tolerance: int = 1,
    keep_up: bool = True,
    prune_empty_effectivity: bool = True,
) -> BOMTree:
    """Convenience functional API to filter a BOM tree by effectivity date."""
    criteria = FilterCriteria(
        reference_date=reference_date,
        month_tolerance=month_tolerance,
        keep_up=keep_up,
        prune_empty_effectivity=prune_empty_effectivity,
    )
    filtrator = DateFilter(criteria)
    return filtrator.filter_tree(tree)


# Alias for compatibility with test frameworks
DateFilterEngine = DateFilter
