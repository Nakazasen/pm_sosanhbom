# TEST_INFRA.md — E2E Testing Infrastructure & Philosophy

## 1. Testing Philosophy

The test infrastructure for the **BOM Comparison Automation Modernization** project adheres to strict industrial software quality assurance standards, ensuring that the modernization of legacy Excel VBA/VBScript workflows into high-performance Python produces bit-accurate, resilient, and enterprise-grade results.

### 1.1 Core Principles
1. **Fail-Closed & Authenticity (Anti-Cheating Guard)**:
   - Tests MUST NOT use facade mocks that always return True or dummy assertions.
   - Every assertion is derived from the authoritative specifications in `ORIGINAL_REQUEST.md`, `PROJECT.md`, and reverse-engineered legacy VBA ground truth (`form_ssbom.xlsm`, `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`, `tonghop_new12052026_ma1.xlsm`).
2. **Deterministic Output Verification**:
   - Numerical quantities, engineering revisions, part codes, and decision states are checked for exact equivalence.
   - Non-deterministic values (such as file modification timestamps and execution duration) are tested with strict bounds and regex patterns.
3. **Progressive Testability & Adapter Isolation**:
   - Domain logic (`src/core`) is 100% decoupled from external I/O protocols.
   - Adapters for external dependencies (Siemens Teamcenter Active Workspace TC14 via Selenium, SAP Logon 770 via COM interop, Microsoft Outlook via MAPI COM) provide clean interface contracts and high-fidelity test mocks when live servers are offline.
4. **Pyramid & Multi-Tier Stratification**:
   - Testing is structured into 4 formal test tiers, ensuring high coverage from unit features to end-to-end real-world production data.

---

## 2. Multi-Tier Architecture & Directory Layout

The automated test suite is organized into modular directories under `tests/`:

```
tests/
├── conftest.py                   # Shared pytest fixtures, data generators, and mocks
├── tier1_features/               # Feature-by-feature tests in isolation (F1 .. F28)
│   ├── test_f01_plm_parser.py
│   ├── test_f02_bom_hierarchy.py
│   ├── test_f03_date_filter.py
│   ├── test_f04_model_pruner.py
│   ├── test_f05_unit_resolver.py
│   ├── test_f06_reconciliation.py
│   ├── test_f07_missing_parts.py
│   ├── test_f08_cross_station.py
│   ├── test_f09_annotation_migration.py
│   ├── test_f10_msi_decision.py
│   ├── test_f11_tc14_headless_session.py
│   ├── test_f12_tc14_authentication.py
│   ├── test_f13_tc14_search_navigation.py
│   ├── test_f14_tc14_export_pipeline.py
│   ├── test_f15_sap_com_automation.py
│   ├── test_f16_sap_multilogon.py
│   ├── test_f17_sap_cs12_execution.py
│   ├── test_f18_sap_fail_closed_guard.py
│   ├── test_f19_sap_export_routing.py
│   ├── test_f20_machine_code_unification.py
│   ├── test_f21_dynamic_r3_header.py
│   ├── test_f22_leader_workspace.py
│   ├── test_f23_member_workspace.py
│   ├── test_f24_consolidated_report.py
│   ├── test_f25_outlook_notification.py
│   ├── test_f26_e2e_regression.py
│   ├── test_f27_adversarial_coverage.py
│   └── test_f28_standalone_packaging.py
├── tier2_boundaries/             # Boundary conditions and extreme edge cases
│   ├── test_empty_boms.py
│   ├── test_depth_extremes.py
│   ├── test_date_edge_cases.py
│   └── test_character_encoding.py
├── tier3_combinations/           # Combinatorial interactions across pipeline
│   ├── test_pipeline_tree_to_units.py
│   ├── test_pipeline_reconciliation.py
│   └── test_pipeline_msi_crosscheck.py
└── tier4_real_world/             # Ground truth verification with legacy workbooks
    ├── test_legacy_unit_resolver.py
    └── test_legacy_form_ssbom.py
```

### 2.1 Tier Breakdown Details
- **Tier 1 (Isolated Features F1..F28)**:
  - Minimum 5 distinct test cases per feature.
  - Verifies contract compliance, input parsing, state transitions, return values, and failure handling in isolation.
- **Tier 2 (Boundaries & Edge Cases)**:
  - Empty files (0 bytes, 0 rows).
  - Single-row BOMs (root assembly without children).
  - Maximum depth (up to 6 levels per specification) and invalid deep trees (>6 levels).
  - Date transitions: Leap years (e.g. Feb 29), month-end boundaries, same year with difference > 1 month vs <= 1 month, and "UP" preservation.
  - Non-ASCII, Japanese (Kanji/Katakana), Vietnamese diacritics, and symbols in Part Codes and Item Descriptions.
- **Tier 3 (Combinatorial Interactions)**:
  - Multi-stage pipelines: 14-column PLM parsing $\to$ Dual-pass date validity filtering $\to$ 6-machine model decomposition $\to$ $O(N)$ Unit resolution.
  - End-to-end 3-way reconciliation: Member CTTT + PLM BOM + SAP R3 BOM $\to$ line-by-line comparison $\to$ Grand Total aggregation $\to$ Missing parts detection.
  - Combined 3-way check + MSI 9-branch decision engine evaluation.
- **Tier 4 (Real-World Legacy Workbooks)**:
  - Validates output directly against legacy workbooks: `form_ssbom.xlsm` and `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`.
  - Proves 100% bit-accurate matching between the new Python engine and the legacy Excel VBA outputs.

---

## 3. Shared Fixtures & Test Data Generation (`tests/conftest.py`)

All tests share reusable, stateless fixtures managed in `tests/conftest.py`:

| Fixture Name | Type | Description |
|--------------|------|-------------|
| `temp_workspace` | `Path` | Isolated temporary directory automatically cleaned up after test run |
| `sample_14col_plm_df` | `pd.DataFrame` | Valid 14-column DataFrame matching Teamcenter TC14 export schema |
| `sample_plm_excel_file` | `Path` | Valid `.xlsx` workbook on disk containing multi-level BOM hierarchy |
| `sample_cs12_html_file` | `Path` | Valid HTML-disguised `.xls` workbook generated by SAP CS12 export |
| `sample_cttt_df` | `pd.DataFrame` | Member work instruction table (`CTTT`) with quantities and part codes |
| `sample_msi_df` | `pd.DataFrame` | Sub-unit barcode, 3-char MSI code, and service comment test records |
| `mock_tc14_driver` | `MagicMock` | Simulated Selenium WebDriver with TC14 DOM elements and session methods |
| `mock_sap_session` | `MagicMock` | Simulated SAP GUI Scripting COM session (`wnd[0]`, `sbar`, `tbar[1]/btn[8]`) |
| `sample_bolocbom_rules` | `Dict` | Rules for the 6 machine models (Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris) |

---

## 4. Test Execution Commands

### 4.1 Run All Automated Tests
```powershell
# Run the complete test suite
python -m pytest tests/ -v

# Run with concise summary and short traceback
python -m pytest tests/ -q --tb=short
```

### 4.2 Run by Tier
```powershell
# Run Tier 1: Feature Isolation Tests (F1..F28)
python -m pytest tests/tier1_features/ -v

# Run Tier 2: Boundary and Edge Cases
python -m pytest tests/tier2_boundaries/ -v

# Run Tier 3: Combinatorial Pipelines
python -m pytest tests/tier3_combinations/ -v

# Run Tier 4: Real-World Legacy Ground Truth Validation
python -m pytest tests/tier4_real_world/ -v
```

### 4.3 Run by Specific Feature (Examples)
```powershell
# Test F1 (14-Column PLM Parser)
python -m pytest tests/tier1_features/test_f01_plm_parser.py -v

# Test F5 (O(N) Unit Resolver)
python -m pytest tests/tier1_features/test_f05_unit_resolver.py -v

# Test F10 (MSI Decision Engine)
python -m pytest tests/tier1_features/test_f10_msi_decision.py -v

# Test F17 (SAP CS12 Execution)
python -m pytest tests/tier1_features/test_f17_sap_cs12_execution.py -v
```
