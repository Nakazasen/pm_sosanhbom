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

from src.core.default_rules import DEFAULT_MODEL_RULES
from src.core.models import BOMNode, BOMTree, ModelRule, PruneAction

logger = logging.getLogger(__name__)


def normalize_model_name(name: str) -> str | None:
    """Normalize model name for case-insensitive lookup."""
    if not name or not isinstance(name, str):
        return None
    cleaned = name.strip().lower()
    for known in DEFAULT_MODEL_RULES:
        if known.lower() == cleaned:
            return known
    # Check if prefix matches (e.g. 'iris' -> 'Iris2024')
    for known in DEFAULT_MODEL_RULES:
        if known.lower().startswith(cleaned):
            return known
    return None


class ModelPruner:
    """Decomposition pruner supporting the 4 action rules across machine models."""

    def __init__(
        self,
        custom_rules: dict[str, list[ModelRule]] | None = None,
        strict: bool = False,
    ) -> None:
        self.rules_registry: dict[str, list[ModelRule]] = {}
        self.strict = strict

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

            if rule.matches(node):
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
                    node.children = self._apply_rule_to_nodes(node.children, rule, visited=visited)
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
            while idx < len(current_nodes):
                node = current_nodes[idx]
                if not rule.matches(node):
                    idx += 1
                    continue

                # Node matches rule!
                has_children = node.has_children
                current_level = node.level
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
                    idx += 1

                # Rule 3: has_children == True AND level >= next_level AND level < 6
                elif has_children and current_level >= next_level and current_level < 6:
                    # Do nothing
                    idx += 1

                # Rule 4: has_children == False
                elif not has_children:
                    current_nodes.pop(idx)
                    # Don't increment idx

                else:
                    idx += 1

        return current_nodes


def prune_by_model(
    tree: BOMTree,
    model_name: str,
    rules: list[ModelRule] | None = None,
) -> BOMTree:
    """Convenience functional API to prune a BOM tree for a specific machine model."""
    pruner = ModelPruner()
    return pruner.prune_tree(tree, model_name=model_name, rules=rules)
