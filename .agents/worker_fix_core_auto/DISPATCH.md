## 2026-09-17T06:38:07Z

# Task Assignment: Worker Fix Core & Automation

## Identity
- Archetype: teamwork_preview_worker
- Role: Core Engine & Automation Remediation Worker
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\worker_fix_core_auto
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

MANDATORY INTEGRITY WARNING:
DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A teamwork_preview_auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.

## Context & Inputs
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` first.
- Read `D:\Sandbox\pm_sosanhbom\PROJECT.md`.
- Read `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md`.
- Read Explorer Fix 2 report at `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_2\handoff.md`.
- Read Explorer Fix 3 report at `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_3\handoff.md`.
- Read Auditor 1 report at `D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md`.

## Exclusive File Ownership
You exclusively own and must edit/create:
1. `src/core/models.py`: Relax BOMNode depth `level: int = Field(ge=0, le=20)`.
2. `src/core/date_filter.py`: Add `visited: set[int] = None` tracking `id(node)` to `_filter_node_recursive()` to prevent infinite recursion on cyclic graphs.
3. `src/core/model_pruner.py`: Add `visited: set[int] = None` tracking `id(node)` to `_apply_rule_to_nodes()` to prevent infinite recursion on cyclic graphs.
4. `src/core/reconciliation.py`:
   - Enhance `_normalize_rev()` to handle leading zeros (`"01"` vs `"1"`), case insensitivity (`"a"` vs `"A"`), float formatting (`1.0` vs `"1"`), and prefix stripping (`"Rev.01"`, `"/A;"`).
   - Enhance `_to_float()` to handle European/Vietnamese comma decimals (`"1,5"` -> `1.5`) and thousands separators.
5. `src/core/adapters.py`:
   - Implement R5 Provider Adapter pattern: `PLMProvider(ABC)` with `TeamcenterSeleniumAdapter` and `ExcelPLMAdapter`; `ERPProvider(ABC)` with `SAPR3COMAdapter` and `ExcelR3Adapter`.
6. `src/core/__init__.py`: Export provider adapters.
7. `src/automation/tc14/client.py` & `apps/1.0.0/src/automation/tc14/client.py`:
   - Fix `login()`: add `timeout: Optional[float] = None` parameter and initialize `wait = WebDriverWait(driver, op_timeout)`.
8. `src/automation/tc14/session.py`: Enhance `is_session_alive()` and add `ensure_driver_alive()` self-healing reconnection.
9. `src/automation/sap/connection.py`: Enhance `is_connected()` and add self-healing reconnect logic to `get_session()`.
10. `src/automation/sap/parser.py`: Export backward-compatible `SAPBOMParser` class aliasing `ResilientR3Parser` and `parse_r3_cs12_file`.
11. `tests/unit/test_adapters.py`: Unit tests verifying provider swapping and core engine independence.

Run tests: `pytest tests/unit/test_adapters.py tests/unit/test_tc14_automation.py tests/tier1_features/test_f12_tc14_authentication.py -v`.
Write report to `D:\Sandbox\pm_sosanhbom\.agents\worker_fix_core_auto\handoff.md` and message parent orchestrator.
