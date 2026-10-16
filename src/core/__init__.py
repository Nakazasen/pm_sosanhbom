"""Core BOM Tree, Filtering & Unit Resolution Engine.

Modern Python modernization of Kyocera BOM comparison core.
Exports are lazily loaded on attribute access to prevent importing
heavy dependencies (pandas, openpyxl, selenium) during basic application startup.
"""

from __future__ import annotations

import importlib
from typing import Any

_ENTRY_MAP: dict[str, str] = {
    "Assessment4M": "src.core.jig_manager",
    "Decision4M": "src.core.jig_manager",
    "DEFAULT_MODEL_RULES": "src.core.default_rules",
    "BOMNode": "src.core.models",
    "BOMTree": "src.core.models",
    "DateFilter": "src.core.date_filter",
    "ERPProvider": "src.core.adapters",
    "ExcelPLMAdapter": "src.core.adapters",
    "ExcelR3Adapter": "src.core.adapters",
    "ExplanationRecord": "src.core.inheritance_engine",
    "FilterCriteria": "src.core.models",
    "FixSerialMaster": "src.core.msi_engine",
    "InheritanceEngine": "src.core.inheritance_engine",
    "InheritanceSummary": "src.core.inheritance_engine",
    "JIGManager": "src.core.jig_manager",
    "MatchMode": "src.core.models",
    "ModelPruner": "src.core.model_pruner",
    "ModelRule": "src.core.models",
    "MSIEngine": "src.core.msi_engine",
    "MSIEvaluationResult": "src.core.msi_engine",
    "PLMProvider": "src.core.adapters",
    "PLMTreeParser": "src.core.tree_parser",
    "Presence4M": "src.core.jig_manager",
    "PruneAction": "src.core.models",
    "ReconciliationEngine": "src.core.reconciliation",
    "ReconciliationResult": "src.core.reconciliation",
    "SAPR3COMAdapter": "src.core.adapters",
    "STANDARD_JIG_SERIES": "src.core.jig_manager",
    "TeamcenterSeleniumAdapter": "src.core.adapters",
    "UnitResolver": "src.core.unit_resolver",
    "aggregate_cross_station": "src.core.reconciliation",
    "detect_missing_parts": "src.core.reconciliation",
    "evaluate_msi_branch": "src.core.msi_engine",
    "extract_expiry_date": "src.core.date_filter",
    "filter_by_date": "src.core.date_filter",
    "get_app_base_dir": "src.core.path_utils",
    "get_asset_path": "src.core.path_utils",
    "is_effectivity_expired": "src.core.date_filter",
    "is_unc_path_accessible": "src.core.path_utils",
    "match_index_mix": "src.core.inheritance_engine",
    "migrate_annotations": "src.core.reconciliation",
    "normalize_model_name": "src.core.model_pruner",
    "parse_plm_excel": "src.core.tree_parser",
    "prune_by_model": "src.core.model_pruner",
    "reconcile_single_row": "src.core.reconciliation",
    "reconcile_three_way": "src.core.reconciliation",
    "refresh_workbook_pivots": "src.core.inheritance_engine",
    "resolve_units": "src.core.unit_resolver",
    "safe_exists": "src.core.path_utils",
    "safe_is_dir": "src.core.path_utils",
    "safe_is_file": "src.core.path_utils",
    "update_workbook_with_inheritance": "src.core.inheritance_engine",
}


def __getattr__(name: str) -> Any:
    if name in _ENTRY_MAP:
        mod = importlib.import_module(_ENTRY_MAP[name])
        val = getattr(mod, name)
        globals()[name] = val
        return val
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = list(_ENTRY_MAP.keys())
