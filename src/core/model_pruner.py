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

import datetime
import logging
import re
import shutil
from pathlib import Path
from typing import Any, Callable, Sequence

from src.core.default_rules import DEFAULT_MODEL_RULES
from src.core.models import BOMNode, BOMTree, ModelRule, PruneAction

logger = logging.getLogger(__name__)


COMMON_MODELS: list[str] = ["Virgo", "Libra2", "Spica", "Iris2024", "Sirius2", "Mebius", "Polaris"]


def normalize_model_name(name: str) -> str | None:
    """Normalize model name for case-insensitive lookup."""
    if not name or not isinstance(name, str):
        return None
    cleaned = name.strip().lower()
    for known in COMMON_MODELS:
        k_low = known.lower()
        if k_low == cleaned or k_low.startswith(cleaned) or cleaned.startswith(k_low):
            return known
        k_alpha = re.sub(r"[^a-zA-Z0-9]", "", k_low)
        c_alpha = re.sub(r"[^a-zA-Z0-9]", "", cleaned)
        if k_alpha and (k_alpha == c_alpha or c_alpha.startswith(k_alpha) or k_alpha.startswith(c_alpha)):
            return known
    return None


def auto_detect_model_from_path(file_path: str | Path) -> str | None:
    """Intelligently detect Kyocera machine model name from filename or directory hierarchy."""
    if not file_path:
        return None
    p = Path(file_path)
    stem = p.stem.strip()
    full_str = str(p).replace("\\", "/")

    models_to_check = list(COMMON_MODELS)
    try:
        from src.core.bom_filter_manager import get_bom_filter_manager
        db_models = get_bom_filter_manager().get_all_models()
        for m in db_models:
            if m not in models_to_check:
                models_to_check.append(m)
    except Exception:
        pass

    # 1. Direct match with registered models in stem or path
    for known in models_to_check:
        pattern = re.compile(rf"(?<![a-zA-Z0-9]){re.escape(known)}(?![a-zA-Z0-9])", re.IGNORECASE)
        if pattern.search(stem) or pattern.search(full_str):
            return known
        k_clean = re.sub(r"[^a-zA-Z0-9]", "", known.lower())
        s_clean = re.sub(r"[^a-zA-Z0-9]", "", stem.lower())
        if k_clean and len(k_clean) >= 4 and k_clean in s_clean:
            return known

    # 2. Extract material / machine code and query MachineDictService
    try:
        from src.services.machine_dict_service import get_machine_dict_service
        dict_svc = get_machine_dict_service()
        # Check 10-char codes (e.g. 110C0V4PH0, T10C0L2US0)
        codes_10 = re.findall(r"[0-9A-Za-z]{10}", stem)
        for code in codes_10:
            resolved = dict_svc.get_model_for_machine_code(code)
            if resolved:
                norm = normalize_model_name(resolved)
                if norm:
                    return norm

        # Check discrete 4-character tokens (e.g. C0V4, 2YJ0, 02Y4)
        tokens = re.split(r"[-_\s.]+", stem)
        for t in tokens:
            if len(t) == 4 and any(c.isdigit() for c in t) and t.isalnum():
                resolved = dict_svc.get_model_for_machine_code(t)
                if resolved:
                    norm = normalize_model_name(resolved)
                    if norm:
                        return norm
    except Exception:
        pass

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
                    from src.core.bom_filter_manager import get_bom_filter_manager
                    self.filter_manager = get_bom_filter_manager()
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

    def auto_detect_model(self, file_path: str | Path) -> str | None:
        """Intelligently detect machine model name from filename or folder path."""
        return auto_detect_model_from_path(file_path)

    def prune_file(
        self,
        input_path: str | Path,
        output_path: str | Path | None = None,
        model_name: str | None = None,
        rules: list[ModelRule] | None = None,
        sheet_name: str | int | None = None,
        create_backup: bool = True,
        backup_dir: str | Path | None = None,
    ) -> dict[str, Any]:
        """Prune an existing PLM Excel file on disk according to model decomposition rules.

        Supports:
        - In-place pruning (default when output_path is None or identical to input_path).
        - Writing to an alternative destination file.
        - Automatic safety backup (default: True) into backup_before_filter/ or custom directory.
        - Full format preservation and .xlsm VBA macro preservation (keep_vba=True).
        - Safe handling of already-pruned files (xóa thêm/lọc sót logic).
        - Comprehensive statistics dictionary return.
        """
        import openpyxl
        from src.core.tree_parser import PLMTreeParser

        in_p = Path(input_path).resolve()
        if not in_p.is_file():
            raise FileNotFoundError(f"Tệp PLM không tồn tại: {in_p}")

        if in_p.suffix.lower() == ".xls":
            raise ValueError(
                f"Định dạng tệp Excel cũ (.xls) không được hỗ trợ để cắt tỉa trực tiếp: {in_p.name}. "
                "Vui lòng mở tệp trong Microsoft Excel và lưu dưới định dạng .xlsx hoặc .xlsm trước khi lọc."
            )

        effective_model = model_name or self.auto_detect_model(in_p)
        effective_rules = rules if rules is not None else (self.get_rules_for_model(effective_model) if effective_model else [])

        out_p = Path(output_path).resolve() if output_path else in_p
        out_p.parent.mkdir(parents=True, exist_ok=True)

        backup_file_path: Path | None = None
        if create_backup:
            b_dir = Path(backup_dir).resolve() if backup_dir else in_p.parent / "backup_before_filter"
            b_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file_path = b_dir / f"{in_p.stem}_{ts}{in_p.suffix}"
            try:
                shutil.copy2(str(in_p), str(backup_file_path))
                logger.info("Created pre-pruning backup: %s", backup_file_path)
            except PermissionError as p_err:
                raise PermissionError(
                    f"Không thể tạo bản sao lưu vì tệp đang bị khóa: {in_p}. Vui lòng đóng tệp trong Excel trước khi lọc."
                ) from p_err
            except Exception as b_err:
                logger.warning("Could not create backup for %s: %s", in_p, b_err)

        parser = PLMTreeParser()
        raw_tree = parser.parse_file(in_p, sheet_name=sheet_name)
        if not effective_model and getattr(raw_tree, "model_name", None):
            effective_model = raw_tree.model_name
            if not effective_rules:
                effective_rules = self.get_rules_for_model(effective_model)

        pruned_tree = self.prune_tree(raw_tree, model_name=effective_model, rules=effective_rules)
        retained_nodes = pruned_tree.flatten()
        retained_row_indices = {n.row_index for n in retained_nodes if n.row_index is not None}

        is_vba = in_p.suffix.lower() in (".xlsm", ".xltm")
        try:
            wb = openpyxl.load_workbook(str(in_p), keep_vba=is_vba)
        except PermissionError as p_err:
            raise PermissionError(
                f"Tệp đang được mở bởi ứng dụng khác (Microsoft Excel): {in_p}. Vui lòng đóng tệp trước khi lọc."
            ) from p_err

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

            if len(rows_to_delete) > 0 or out_p != in_p:
                try:
                    wb.save(str(out_p))
                except PermissionError as p_err:
                    raise PermissionError(
                        f"Không thể ghi đè tệp vì tệp đang mở trong Microsoft Excel: {out_p}. "
                        "Vui lòng đóng tệp trong Excel và thử lại."
                    ) from p_err
        finally:
            wb.close()

        final_rows = max_r - len(rows_to_delete)
        return {
            "input_path": str(in_p),
            "output_path": str(out_p),
            "backup_path": str(backup_file_path) if backup_file_path else None,
            "model_name": effective_model or "Unknown",
            "initial_rows": max_r,
            "final_rows": final_rows,
            "deleted_rows": len(rows_to_delete),
            "rules_count": len(effective_rules),
            "success": True,
            "error": None,
        }

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
        res = self.prune_file(
            input_path=input_path,
            output_path=output_path,
            model_name=model_name,
            rules=rules,
            sheet_name=sheet_name,
            create_backup=False,
        )
        return Path(res["output_path"])

    def prune_directory(
        self,
        dir_path: str | Path,
        model_name: str | None = None,
        rules: list[ModelRule] | None = None,
        recursive: bool = True,
        file_pattern: str = "PLM_*.xls*",
        create_backup: bool = True,
        backup_dir: str | Path | None = None,
        output_dir: str | Path | None = None,
        callback: Callable[[int, int, dict[str, Any]], None] | None = None,
    ) -> list[dict[str, Any]]:
        """Batch prune multiple PLM Excel files located in a directory.

        Args:
            dir_path: Directory path to scan.
            model_name: Explicit machine model name or None for auto-detection.
            rules: Optional explicit list of ModelRule records to apply.
            recursive: Whether to scan subdirectories recursively.
            file_pattern: Filename match pattern (default 'PLM_*.xls*').
            create_backup: Whether to create backups before pruning.
            backup_dir: Target backup folder or None for local backup_before_filter.
            output_dir: Target output directory (if None, overwrites in-place).
            callback: Optional progress callback receiving (current_idx, total_count, file_result).

        Returns:
            List of result dictionaries.
        """
        d_p = Path(dir_path).resolve()
        if not d_p.is_dir():
            raise NotADirectoryError(f"Thư mục không tồn tại: {d_p}")

        iterator = d_p.rglob(file_pattern) if recursive else d_p.glob(file_pattern)
        found_files: list[Path] = []
        for f in iterator:
            if not f.is_file():
                continue
            if any(b in f.parts for b in ("backup_before_filter", "backupTC14full", ".bak", "old")):
                continue
            if f.name.startswith("~$") or f.stem.endswith("_filtered"):
                continue
            if f.suffix.lower() in (".xlsx", ".xlsm", ".xltm"):
                found_files.append(f)

        found_files.sort(key=lambda p: str(p).lower())
        results: list[dict[str, Any]] = []
        total = len(found_files)

        for idx, file_p in enumerate(found_files, start=1):
            out_target = None
            if output_dir:
                out_d = Path(output_dir).resolve()
                out_d.mkdir(parents=True, exist_ok=True)
                out_target = out_d / file_p.name

            try:
                res = self.prune_file(
                    input_path=file_p,
                    output_path=out_target,
                    model_name=model_name,
                    rules=rules,
                    create_backup=create_backup,
                    backup_dir=backup_dir,
                )
            except Exception as exc:
                logger.error("Error pruning file %s: %s", file_p, exc)
                res = {
                    "input_path": str(file_p),
                    "output_path": str(out_target or file_p),
                    "backup_path": None,
                    "model_name": model_name or "Unknown",
                    "initial_rows": 0,
                    "final_rows": 0,
                    "deleted_rows": 0,
                    "rules_count": 0,
                    "success": False,
                    "error": str(exc),
                }

            results.append(res)
            if callback:
                try:
                    callback(idx, total, res)
                except Exception:
                    pass

        return results



def prune_by_model(
    tree: BOMTree,
    model_name: str,
    rules: list[ModelRule] | None = None,
) -> BOMTree:
    """Convenience functional API to prune a BOM tree for a specific machine model."""
    pruner = ModelPruner()
    return pruner.prune_tree(tree, model_name=model_name, rules=rules)
