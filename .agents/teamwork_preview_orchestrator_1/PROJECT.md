# Project: BOM Comparison Automation Rewrite (Modernization)

## Architecture
Modern, high-performance, modular Python system replacing Excel VBA/VBScript:
- `src/core/`: In-memory BOM tree models, 6-level date effectivity pruning, machine model pruning (`BolocBom`), $O(N)$ Unit resolver, 3-way cross-reconciliation engine (CTTT vs PLM vs R3), and MSI 9-branch decision engine.
- `src/automation/tc14/`: Selenium WebDriver client for Siemens Teamcenter Active Workspace (TC14) at `http://tcmp3gwb:3000/`, headless CDP download, auto-relogin.
- `src/automation/sap/`: Python `win32com.client` COM engine for SAP Logon 770 (`P1J(ERP60-AWS)-VN`), CS12 transaction execution, fail-closed status bar guard, dynamic header parser.
- `src/gui/`: PyQt6 desktop user interface with dedicated Leader and Member workspaces, visual progress tracking, and configuration management.
- `src/reporting/`: Consolidated Excel report generator (`form_ssbom` format) and Outlook HTML notification generator.
- `tests/`: Multi-tier automated test suite (Tiers 1-5) ensuring 100% bit-accurate numerical and textual match against legacy VBA baselines.

```
                  ┌────────────────────────────────────────────────────────┐
                  │                 PyQt6 Desktop GUI                      │
                  │   [Leader Sub-module]          [Member Sub-module]     │
                  └──────────────┬─────────────────────────┬───────────────┘
                                 │                         │
                                 ▼                         ▼
                  ┌────────────────────────────────────────────────────────┐
                  │              Application Orchestration Layer           │
                  └───────┬────────────────────────┬───────────────┬───────┘
                          │                        │               │
                          ▼                        ▼               ▼
          ┌───────────────────────────┐  ┌──────────────────┐  ┌───────────────────────────┐
          │   Teamcenter TC14 Client  │  │  Core BOM Tree   │  │    SAP R3 CS12 Client     │
          │ (Selenium / Edge / Chrome)│  │  & Diff Engine   │  │   (win32com / SAP GUI)    │
          └─────────────┬─────────────┘  └────────┬─────────┘  └─────────────┬─────────────┘
                        │                         │                          │
                        ▼                         ▼                          ▼
               TC14 Active Workspace          3-Way Diff &             SAP Logon 770 /
               http://tcmp3gwb:3000/          Unit Resolver         P1J(ERP60-AWS)-VN (CS12)
```

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| F1 | 14-Column PLM Parser | Parse TC14 Active Workspace Excel exports with exact 14 columns | M1 | Survey E1 |
| F2 | 6-Level BOM Tree Hierarchy | In-memory tree structure supporting parent-child stack indexing | M1 | Survey E1 |
| F3 | Dual-Pass Date Validity Filter | Prune branches by "to <date>" (expired year/month), preserve "UP", empty effectivity rules | M1 | Survey E1 |
| F4 | Model Decomposition & Pruning | 4 action rules on matched nodes across 6 models (Virgo, Libra2, Iris2024, etc.) in `BolocBom` | M1 | Survey E1 |
| F5 | $O(N)$ Unit Resolver | In-memory depth-first stack traversal replacing 5,619-row lookup spreadsheet | M1 | Survey E1 |
| F6 | Three-Way Cross-Reconciliation | Line-by-line reconciliation of CTTT vs PLM vs R3 part quantities and revisions | M2 | Survey E1 |
| F7 | Missing Parts Detection | Reverse lookup on Sheet PLM surfacing parts engineered but omitted in CTTT | M2 | Survey E1 |
| F8 | Cross-Station Total Aggregation | Sum shared components across sub-units and reconcile against machine grand totals | M2 | Survey E1 |
| F9 | Annotation Migration Engine | Automatically preserve and migrate member explanations from old PLM sheets across revisions | M2 | Survey E1 |
| F10 | MSI & Fix Serial Decision Engine | 9-branch decision table evaluating barcode parts, 3-char codes, and service notes vs master file | M2 | Survey E1 |
| F11 | TC14 Headless Browser Session | Headless Edge/Chrome WebDriver with CDP download behavior and session persistence | M3 | Survey E2 |
| F12 | TC14 Automated Authentication | Login with `vn_pe03 / vn_pe03`, cookie handling, and session expiration detection | M3 | Survey E2 |
| F13 | TC14 Hash-Based Search & Open | Direct navigation to item revision avoiding UI input lag and autocomplete delays | M3 | Survey E2 |
| F14 | TC14 BOM Export Pipeline | Trigger `Awp0ExportToExcel` / Content tab export downloading full PLM BOM `.xlsx` | M3 | Survey E2 |
| F15 | SAP GUI 770 COM Automation | Connect to `P1J(ERP60-AWS)-VN` via `win32com.client` with process life-cycle management | M4 | Survey E3 |
| F16 | SAP Multi-Logon Resolution | Detect and handle option 2 (`radMULTI_LOGON_OPT2`) to reclaim orphaned sessions | M4 | Survey E3 |
| F17 | SAP CS12 Transaction Execution | Execute CS12 with Plant 2200, BOM Usage pp01, Alternative 01, and valid date | M4 | Survey E3 |
| F18 | SAP Status Bar Fail-Closed Guard | Check `wnd[0]/sbar` message type 'E' to prevent crashes on missing machine BOMs | M4 | Survey E3 |
| F19 | SAP Spreadsheet Export & Path Routing | Export via `SAPLSPO5:0150` to `R3_<material>_<dd>_<mm>_<yyyy>.xls` in target machine directory | M4 | Survey E3 |
| F20 | Machine Model Code Unification | Unify `ma1` (`110*`) and `maT` (`T10*`) workflows into a configurable engine | M4 | Survey E3 |
| F21 | Dynamic R3 Header Resolution | Resiliently resolve Part Code, Quantity, and Revision columns across ALV layout variations | M4 | Survey E3 |
| F22 | Leader Management Workspace | Create project folders, track member submission statuses, and trigger batch processing | M5 | Survey E1/PPTX |
| F23 | Member Input Workspace | Member CTTT, MSI, and Label 7980/7990 entry with preliminary PLM/R3 self-check | M5 | Survey E1/PPTX |
| F24 | Consolidated Report Generation | Export final consolidated Excel workbook matching `form_ssbom.xlsm` formatting | M5 | Survey E1 |
| F25 | Outlook HTML Email Notification | Automated draft/send of summary notification email to team members via Outlook COM | M5 | Survey E1 |
| F26 | E2E Regression & Validation Suite | 100% match verification against legacy `form_ssbom.xlsm` and test cases across Tiers 1-4 | M6 | Survey E1/E2/E3 |
| F27 | Adversarial Coverage Hardening | White-box stress-testing, boundary condition tests, and edge case coverage (Tier 5) | M6 | Project Pattern |
| F28 | Standalone PyInstaller Executable | Standalone executable distribution (.exe) with embedded assets and zero manual install | M6 | Request R5 |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| M1 | Core BOM Tree & Unit Resolver | F1, F2, F3, F4, F5 | none | DONE |
| M2 | Cross-Reconciliation & MSI Engine | F6, F7, F8, F9, F10 | M1 | DONE |
| M3 | Teamcenter TC14 Web Automation | F11, F12, F13, F14 | none | DONE |
| M4 | SAP R3 CS12 Automation | F15, F16, F17, F18, F19, F20, F21 | none | DONE |
| M5 | PyQt6 Desktop Application & Reports | F22, F23, F24, F25 | M1, M2, M3, M4 | DONE |
| M6 | Final Verification & PyInstaller Packaging | F26, F27, F28 (Pass 100% E2E tests + Tier 5 Hardening + .exe) | M1..M5, TEST_READY.md | DONE |

## Interface Contracts

### `src/core/tree.py` ↔ `src/core/reconciliation.py`
```python
@dataclass
class BOMNode:
    level: int
    item_id: str
    item_name: str
    has_children: bool
    quantity: float
    effectivity: str
    revision: str
    unit_name: str
    children: List['BOMNode']

class BOMTreeFilter:
    def filter_tree(self, root_nodes: List[BOMNode], model_name: str, reference_date: datetime.date) -> List[BOMNode]:
        ...
```

### `src/automation/tc14/client.py` ↔ Application Layer
```python
class TC14Client:
    def download_bom_full(self, part_number: str, output_dir: Path) -> Path:
        """Returns path to downloaded PLM_*.xlsx"""
        ...
```

### `src/automation/sap/cs12.py` ↔ Application Layer
```python
class SAPCS12Client:
    def download_bom(self, material_code: str, valid_date: datetime.date, output_dir: Path) -> Path:
        """Returns path to downloaded R3_*.xls"""
        ...
```

### `src/core/reconciliation.py` ↔ `src/reporting/excel_export.py`
```python
@dataclass
class ReconciliationResult:
    cttt_rows: pd.DataFrame
    plm_missing_rows: pd.DataFrame
    cttt_totals: pd.DataFrame
    msi_results: pd.DataFrame
    overall_status: str # "OK" or "NG"
```

## Code Layout
```
D:\Sandbox\pm_sosanhbom\
├── src/
│   ├── core/
│   │   ├── __init__.py
│   │   ├── models.py
│   │   ├── tree_parser.py
│   │   ├── date_filter.py
│   │   ├── model_pruner.py
│   │   ├── unit_resolver.py
│   │   ├── reconciliation.py
│   │   └── msi_engine.py
│   ├── automation/
│   │   ├── tc14/
│   │   │   ├── __init__.py
│   │   │   ├── client.py
│   │   │   ├── selectors.py
│   │   │   └── session.py
│   │   └── sap/
│   │       ├── __init__.py
│   │       ├── connection.py
│   │       ├── cs12.py
│   │       └── parser.py
│   ├── gui/
│   │   ├── __init__.py
│   │   ├── app.py
│   │   ├── leader_view.py
│   │   ├── member_view.py
│   │   └── settings_dialog.py
│   └── reporting/
│       ├── __init__.py
│       ├── excel_generator.py
│       └── outlook_mailer.py
├── tests/
│   ├── test_infra/
│   ├── tier1_features/
│   ├── tier2_boundaries/
│   ├── tier3_combinations/
│   ├── tier4_real_world/
│   └── tier5_adversarial/
└── dist/
```
