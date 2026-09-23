"""Machine Model Decomposition and Sub-Assembly Pruning Engine.

Implements the 4 action rules from legacy locbomfull.bas (xulyriengchocacloaimay)
and the BolocBom sheet across the 6 Kyocera machine models:
- Virgo
- Libra2
- Iris2024
- Sirius2
- Mebius
- Polaris

The 4 Action Rules on Matched Nodes:
- Rule 1: has_children == True AND level == 6 -> Prune node itself (phantom branch).
- Rule 2: has_children == True AND level < next_level (has children) ->
          Clear node's children, KEEP the node itself (retains unit assembly for line tracking,
          prunes sub-components that members do not assemble individually).
- Rule 3: has_children == True AND has no children AND level < 6 -> Keep node (already unexpanded).
- Rule 4: has_children == False -> Prune node itself (part is omitted for this model).
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from src.core.default_rules import DEFAULT_MODEL_RULES
from src.core.models import BOMNode, BOMTree, ModelRule, PruneAction

logger = logging.getLogger(__name__)


def normalize_model_name(name: str) -> str | None:
    """Normalize model name for case-insensitive lookup."""
    if not name or not isinstance(name, str):
        return None
    cleaned = name.strip().lower()
    for known in DEFAULT_MODEL_RULES:
        k_low = known.lower()
        if k_low == cleaned or k_low.startswith(cleaned) or cleaned.startswith(k_low):
            return known
        k_alpha = re.sub(r"[^a-zA-Z0-9]", "", k_low)
        c_alpha = re.sub(r"[^a-zA-Z0-9]", "", cleaned)
        if k_alpha and (k_alpha == c_alpha or c_alpha.startswith(k_alpha) or k_alpha.startswith(c_alpha)):
            return known
    return None


class ModelPruner:
    """Decomposition pruner supporting the 4 action rules across machine models."""

    def __init__(
        self,
        custom_rules: dict[str, list[ModelRule]] | None = None,
        strict: bool = False,
        filter_manager: Any | None = None,
        use_db: bool = True,
    ) -> None:
        self.rules_registry: dict[str, list[ModelRule]] = {}
        self.strict = strict
        self._has_custom_rules = bool(custom_rules)
        self.filter_manager = filter_manager
        self.use_db = use_db

        # Populate with default rules from BolocBom
        for model, rule_dicts in DEFAULT_MODEL_RULES.items():
            self.rules_registry[model] = [
                ModelRule(
                    item_name=r.get("item_name"),
                    match_mode=r.get("match_mode", "Full_name"),
                    part_code=r.get("part_code"),
                )
                for r in rule_dicts
            ]

        # Merge any custom rules
        if custom_rules:
            for model, r_list in custom_rules.items():
                self.rules_registry[model] = r_list

    def get_rules_for_model(self, model_name: str) -> list[ModelRule]:
        """Retrieve the list of decomposition rules for a given machine model."""
        if not model_name:
            return []

        # If custom rules were explicitly passed to __init__, respect them first
        if self._has_custom_rules and model_name in self.rules_registry:
            return self.rules_registry[model_name]

        # 1. Query dynamic rules from shared SQLite DB
        if self.use_db:
            try:
                if self.filter_manager is None:
                    from src.core.bom_filter_manager import BOMFilterManager
                    self.filter_manager = BOMFilterManager()
                db_rules = self.filter_manager.get_model_rules_for_pruner(model_name)
                if db_rules:
                    return db_rules
            except Exception as ex:
                logger.warning("Could not load rules from DB for model %s: %s", model_name, ex)

        # 2. Check registry / fallback default rules
        if model_name in self.rules_registry:
            return self.rules_registry[model_name]

        # Case-insensitive match on registry keys
        clean = model_name.strip().lower()
        for key in self.rules_registry:
            if key.lower() == clean:
                return self.rules_registry[key]

        normalized = normalize_model_name(model_name)
        if normalized and normalized in self.rules_registry:
            return self.rules_registry[normalized]
        if self.strict:
            raise ValueError(
                f"Model '{model_name}' not found in BolocBom rules registry. "
                f"Available models: {list(self.rules_registry.keys())}"
            )
        logger.warning("No decomposition rules found for model '%s'; skipping model pruning.", model_name)
        return []

    def prune_tree(
        self,
        tree: BOMTree,
        model_name: str | None = None,
        rules: list[ModelRule] | None = None,
        in_place: bool = False,
    ) -> BOMTree:
        """Prune BOM tree according to model decomposition rules."""
        target_tree = tree if in_place else tree.clone()

        active_rules: list[ModelRule] = []
        if rules is not None:
            active_rules = rules
        elif model_name:
            active_rules = self.get_rules_for_model(model_name)
            target_tree.model_name = model_name

        if not active_rules:
            return target_tree

        # In legacy VBA, rules are evaluated sequentially over the BOM
        for rule in active_rules:
            visited: set[int] = set()
            target_tree.roots = self._apply_rule_to_nodes(target_tree.roots, rule, visited)

        return target_tree

    def _apply_rule_to_nodes(
        self,
        nodes: list[BOMNode],
        rule: ModelRule,
        visited: set[int] | None = None,
        parent: BOMNode | None = None,
    ) -> list[BOMNode]:
        """Apply a single ModelRule to a list of sibling nodes and their subtrees with cycle protection."""
        if visited is None:
            visited = set()

        surviving_nodes: list[BOMNode] = []

        for node in nodes:
            if id(node) in visited:
                surviving_nodes.append(node)
                continue
            visited.add(id(node))

            if rule.matches(node, parent=parent):
                action = rule.determine_action(node)

                if action == PruneAction.DELETE_NODE:
                    # Rule 1 or Rule 4: Delete the node itself
                    logger.debug("Pruning node '%s' (%s) under action %s", node.item_name, node.item_id, action)
                    continue
                elif action == PruneAction.DELETE_CHILDREN:
                    # Rule 2: Keep the unit assembly node, but prune all its child sub-components
                    logger.debug("Pruning children of unit '%s' (%s) under Rule 2", node.item_name, node.item_id)
                    node.children = []
                    surviving_nodes.append(node)
                elif action == PruneAction.KEEP:
                    # Rule 3: Keep node
                    surviving_nodes.append(node)
            else:
                # Node didn't match rule; apply rule recursively to its children
                if node.children:
                    node.children = self._apply_rule_to_nodes(node.children, rule, visited=visited, parent=node)
                surviving_nodes.append(node)

        return surviving_nodes

    def prune_flat_nodes(
        self,
        nodes: list[BOMNode],
        model_name: str | None = None,
        rules: list[ModelRule] | None = None,
    ) -> list[BOMNode]:
        """Replicate legacy VBA row deletion sequentially on a flat list of nodes."""
        active_rules: list[ModelRule] = []
        if rules is not None:
            active_rules = rules
        elif model_name:
            active_rules = self.get_rules_for_model(model_name)

        if not active_rules:
            return list(nodes)

        current_nodes = list(nodes)

        for rule in active_rules:
            idx = 0
            parent_stack: list[BOMNode] = []
            while idx < len(current_nodes):
                node = current_nodes[idx]
                while parent_stack and parent_stack[-1].level >= node.level:
                    parent_stack.pop()
                parent = parent_stack[-1] if parent_stack else None

                if not rule.matches(node, parent=parent):
                    parent_stack.append(node)
                    idx += 1
                    continue

                # Node matches rule!
                current_level = node.level

                if rule.action:
                    act = rule.action.strip().lower()
                    if act in ("prune_node", "delete_node"):
                        current_nodes.pop(idx)
                        continue
                    elif act in ("prune_children", "delete_children"):
                        del_start = idx + 1
                        del_end = del_start
                        while del_end < len(current_nodes) and current_nodes[del_end].level > current_level:
                            del_end += 1
                        del current_nodes[del_start:del_end]
                        parent_stack.append(node)
                        idx += 1
                        continue
                    elif act in ("keep",):
                        parent_stack.append(node)
                        idx += 1
                        continue

                # Legacy flat-node determination based on subsequent row levels
                has_children = node.has_children
                next_level = current_nodes[idx + 1].level if idx + 1 < len(current_nodes) else 0

                # Rule 1: has_children == True AND level == 6
                if has_children and current_level == 6:
                    current_nodes.pop(idx)
                    # Don't increment idx; next item shifted into idx

                # Rule 2: has_children == True AND level < next_level
                elif has_children and current_level < next_level:
                    # Keep current_nodes[idx], delete subsequent rows where level > current_level
                    del_start = idx + 1
                    del_end = del_start
                    while del_end < len(current_nodes) and current_nodes[del_end].level > current_level:
                        del_end += 1
                    del current_nodes[del_start:del_end]
                    parent_stack.append(node)
                    idx += 1

                # Rule 3: has_children == True AND level >= next_level AND level < 6
                elif has_children and current_level >= next_level and current_level < 6:
                    # Do nothing
                    parent_stack.append(node)
                    idx += 1

                # Rule 4: has_children == False
                elif not has_children:
                    current_nodes.pop(idx)
                    # Don't increment idx

                else:
                    parent_stack.append(node)
                    idx += 1

        return current_nodes

    def prune_excel_file(
        self,
        input_path: str | Path,
        output_path: str | Path,
        model_name: str | None = None,
        rules: list[ModelRule] | None = None,
        sheet_name: str | int | None = None,
    ) -> Path:
        """Prune an Excel PLM file while preserving 100% of formatting, headers, and column structure.

        Faithfully mirrors legacy VBA behavior: opens the workbook, identifies rows to delete,
        and removes them directly from the worksheet so all cell formatting, fonts, colors,
        column dimensions, and header names are completely retained.
        """
        import openpyxl
        from src.core.tree_parser import PLMTreeParser

        in_p = Path(input_path)
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)

        parser = PLMTreeParser()
        raw_tree = parser.parse_file(in_p, sheet_name=sheet_name)
        pruned_tree = self.prune_tree(raw_tree, model_name=model_name, rules=rules)

        retained_nodes = pruned_tree.flatten()
        retained_row_indices = {n.row_index for n in retained_nodes if n.row_index is not None}

        wb = openpyxl.load_workbook(str(in_p))
        try:
            if sheet_name is not None:
                ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[sheet_name]
            else:
                ws = wb.active

            max_r = ws.max_row
            # Collect rows to delete in descending order
            rows_to_delete = sorted(
                [r for r in range(2, max_r + 1) if r not in retained_row_indices],
                reverse=True,
            )

            # Group into contiguous blocks to maximize delete performance
            blocks: list[tuple[int, int]] = []
            if rows_to_delete:
                curr_end = rows_to_delete[0]
                curr_start = rows_to_delete[0]
                for r in rows_to_delete[1:]:
                    if r == curr_start - 1:
                        curr_start = r
                    else:
                        blocks.append((curr_start, curr_end - curr_start + 1))
                        curr_start = r
                        curr_end = r
                blocks.append((curr_start, curr_end - curr_start + 1))

            for start_row, count in blocks:
                ws.delete_rows(start_row, count)

            # Clean up openpyxl dangling/hyperlink cells beyond surviving rows
            expected_max_row = max_r - len(rows_to_delete)
            excess_coords = [coord for coord in ws._cells if coord[0] > expected_max_row]
            for coord in excess_coords:
                del ws._cells[coord]

            wb.save(str(out_p))
        finally:
            wb.close()

        return out_p


def prune_by_model(
    tree: BOMTree,
    model_name: str,
    rules: list[ModelRule] | None = None,
) -> BOMTree:
    """Convenience functional API to prune a BOM tree for a specific machine model."""
    pruner = ModelPruner()
    return pruner.prune_tree(tree, model_name=model_name, rules=rules)
