"""Core BOM Tree, Filtering & Unit Resolution Engine.

Modern Python modernization of Kyocera BOM comparison core:
- In-memory BOM tree models with multi-level stack indexing.
- 14-column and 13-column PLM Excel parser.
- Dual-pass effectivity date filtering with 'UP' retention.
- Machine model decomposition pruner across 6 Kyocera models.
- Fast O(N) depth-first Unit Resolver replacing 5,619-row lookup sheet.
"""

from src.core.adapters import (
    ERPProvider,
    ExcelPLMAdapter,
    ExcelR3Adapter,
    PLMProvider,
    SAPR3COMAdapter,
    TeamcenterSeleniumAdapter,
)
from src.core.date_filter import (
    DateFilter,
    extract_expiry_date,
    filter_by_date,
    is_effectivity_expired,
)
from src.core.default_rules import DEFAULT_MODEL_RULES
from src.core.model_pruner import (
    ModelPruner,
    normalize_model_name,
    prune_by_model,
)
from src.core.models import (
    BOMNode,
    BOMTree,
    FilterCriteria,
    MatchMode,
    ModelRule,
    PruneAction,
)
from src.core.msi_engine import (
    FixSerialMaster,
    MSIEngine,
    MSIEvaluationResult,
    evaluate_msi_branch,
)
from src.core.reconciliation import (
    ReconciliationEngine,
    ReconciliationResult,
    aggregate_cross_station,
    detect_missing_parts,
    migrate_annotations,
    reconcile_single_row,
    reconcile_three_way,
)
from src.core.tree_parser import (
    PLMTreeParser,
    parse_plm_excel,
)
from src.core.unit_resolver import (
    UnitResolver,
    resolve_units,
)

__all__ = [
    "DEFAULT_MODEL_RULES",
    "BOMNode",
    "BOMTree",
    "DateFilter",
    "ERPProvider",
    "ExcelPLMAdapter",
    "ExcelR3Adapter",
    "FilterCriteria",
    "FixSerialMaster",
    "MatchMode",
    "ModelPruner",
    "ModelRule",
    "MSIEngine",
    "MSIEvaluationResult",
    "PLMProvider",
    "PLMTreeParser",
    "PruneAction",
    "ReconciliationEngine",
    "ReconciliationResult",
    "SAPR3COMAdapter",
    "TeamcenterSeleniumAdapter",
    "UnitResolver",
    "aggregate_cross_station",
    "detect_missing_parts",
    "evaluate_msi_branch",
    "extract_expiry_date",
    "filter_by_date",
    "is_effectivity_expired",
    "migrate_annotations",
    "normalize_model_name",
    "parse_plm_excel",
    "prune_by_model",
    "reconcile_single_row",
    "reconcile_three_way",
    "resolve_units",
]
