# TEST_INFRA.md — E2E Testing Infrastructure & Philosophy

## 1. Testing Philosophy & Invariants

The test infrastructure for the **Phần Mềm So Sánh BOM Tự Động (Kyocera BOM Comparison System)** project adheres to strict industrial software quality assurance standards. It guarantees that modernizing legacy Excel VBA/VBScript workflows into high-performance Python produces bit-accurate, resilient, and enterprise-grade results.

### 1.1 Core Principles
1. **Fail-Closed & Authenticity (Anti-Cheating Guard)**:
   - Tests MUST NOT use facade mocks that always return True or dummy assertions.
   - Every assertion is derived from authoritative specifications in `ORIGINAL_REQUEST.md`, `PROJECT.md`, and reverse-engineered legacy VBA ground truth (`form_ssbom.xlsm`, `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`, `tonghop_new12052026_ma1.xlsm`, `FIX_SERIAL_DLTOOL_VER010.xls`).
2. **Deterministic Output Verification**:
   - Numerical quantities, engineering revisions, part codes, and decision states are checked for exact equivalence.
   - Non-deterministic values (such as timestamps and execution durations) are tested with strict bounds and regex patterns.
3. **Progressive Testability & Adapter Isolation**:
   - Domain logic (`src/core`) is 100% decoupled from external I/O protocols.
   - Adapters for external dependencies (Siemens Teamcenter Active Workspace TC24/TC14, SAP Logon 770 CS12 via COM interop, Microsoft Outlook via MAPI COM) provide clean interface contracts and high-fidelity test mocks when live servers are offline.
4. **Pyramid & Multi-Tier Stratification**:
   - Testing is structured into 4 formal test tiers, ensuring high coverage from unit features to end-to-end real-world production data.

---

## 2. Multi-Tier Architecture & Directory Layout

The automated E2E test suite is organized under `tests/e2e/`:

```
tests/e2e/
├── __init__.py
├── test_tier1_feature_coverage.py       # Tier 1: Feature Isolation (40 tests, R1..R6)
├── test_tier2_boundary_corner.py        # Tier 2: Boundaries & Corner Cases (30 tests, R1..R6)
├── test_tier3_pairwise_combinations.py  # Tier 3: Pairwise Combinations (6 pipelines)
└── test_tier4_production_scenarios.py   # Tier 4: Real-World Production Workflows (5 scenarios)
```

---

## 3. Test Tier Breakdown & Scope

### Tier 1: Feature Coverage (`test_tier1_feature_coverage.py` — 40 Tests)
Covers all primary functional requirements (R1 to R6) in isolation:
- **R1: Leader Workspace Wizard (7 tests)**:
  - `TC-T1-R1-01`: Model folder & stage initialization (`Virgo`, `Iris2024`, `maT`, `ma1`).
  - `TC-T1-R1-02`: Member assignment & distribution via staffing matrix (`Lichsu`).
  - `TC-T1-R1-03`: Exclusion of machines flagged with Column F = `"X"`.
  - `TC-T1-R1-04`: Inbound routing of PLM & SAP R3 BOM files into machine directories.
  - `TC-T1-R1-05`: Real-time scan of member submission packages via cell `CTTT!Q2 = "OK"`.
  - `TC-T1-R1-06`: Fail-closed gate blocking consolidation when any member is pending.
  - `TC-T1-R1-07`: Archival of submitted member files into `phutrach/` subfolder.
- **R2: BOM Filter Engine Level 1..6 (8 tests)**:
  - `TC-T1-R2-01`: Unconditional retention of effectivity containing `"UP"`.
  - `TC-T1-R2-02`: Expiration and filtering of `"to <date>"` effectivities.
  - `TC-T1-R2-03`: Recursive pruning of all descendant child nodes when parent expires.
  - `TC-T1-R2-04`: Pruning of empty effectivity leaves (`has_children = False`).
  - `TC-T1-R2-05`: Pruning of empty effectivity subtrees (`has_children = True`).
  - `TC-T1-R2-06`: BolocBom rule with `Full_name` exact match pruning.
  - `TC-T1-R2-07`: BolocBom rule with `Part_name` substring match pruning.
  - `TC-T1-R2-08`: Automatic backup of raw TC export into `backupTC14full/`.
- **R3: Annotation Inheritance Engine (6 tests)**:
  - `TC-T1-R3-01`: Multi-version sheet backup (`PLM_old`).
  - `TC-T1-R3-02`: Cleaning of old data ranges (`A2:M` and `O2:Q`).
  - `TC-T1-R3-03`: Carryover of Explanations (Col O) for unchanged parts.
  - `TC-T1-R3-04`: Carryover of Responsible Person (Col P) and Manager Check (Col Q).
  - `TC-T1-R3-05`: Clean blank strings for newly introduced parts.
  - `TC-T1-R3-06`: Archiving superseded source files into `capnhat\old\`.
- **R4: MSI Deep Reconciliation Engine (8 tests)**:
  - `TC-T1-R4-01`: Branch 1: Unit not in PLM -> NG, highlight B, K red.
  - `TC-T1-R4-02`: Branch 2: Normalization of blank service comment to `"-"`.
  - `TC-T1-R4-03`: Branch 4: In PLM, missing from master Fix Serial -> NG, B, K red.
  - `TC-T1-R4-04`: Branch 5: Both code and service match -> OK, B, D, E, K green.
  - `TC-T1-R4-05`: Branch 7: Service note required warning -> OK with warning, E red.
  - `TC-T1-R4-06`: Branch 8: 3-character code mismatch -> NG, D, K red.
  - `TC-T1-R4-07`: Branch 9: Service comment mismatch -> NG, E, K red.
  - `TC-T1-R4-08`: Row 36 Hontai machine code fallback to `PLM!C2`.
- **R5: List JIG Master & 4M Evaluation (5 tests)**:
  - `TC-T1-R5-01`: Loading 15 machine model series from JIG master.
  - `TC-T1-R5-02`: Copying JIG table into Sheet `List JIG`.
  - `TC-T1-R5-03`: Preserving formatting and column widths.
  - `TC-T1-R5-04`: Clearing JIG table on reset/model switch.
  - `TC-T1-R5-05`: Recording 4M evaluation with Production Engineering (KTSX).
- **R6: Member Workspace & Self-Check (6 tests)**:
  - `TC-T1-R6-01`: Auto-loading member assignment (engineer, sub-unit, model).
  - `TC-T1-R6-02`: CTTT table operations (add, clear, export records).
  - `TC-T1-R6-03`: Preliminary self-check against local PLM and R3 with instant OK/NG feedback.
  - `TC-T1-R6-04`: Blocking submission when unaddressed NG items exist without explanation.
  - `TC-T1-R6-05`: Exporting submission package containing 3 sheets (`CTTT`, `MSI`, `Label_7980_7990`).
  - `TC-T1-R6-06`: Verifying submission carries `CTTT!Q2 = "OK"` seal.

---

### Tier 2: Boundary & Corner Cases (`test_tier2_boundary_corner.py` — 30 Tests)
Stress-tests extreme edge cases and failure modes:
- **R1 Boundaries**: Empty machine list, single unsubmitted member in 10-person team, Unicode paths with Japanese Kanji and spaces, 0-byte corrupted workbooks, offline Outlook COM handling.
- **R2 Boundaries**: Kyocera grace period (same year month difference <= 1 retained), leap year Feb 29 boundary, 6-level deep tree recursion without stack overflow, unknown model name handling, 0-row empty PLM files.
- **R3 Boundaries**: Consecutive multi-version updates (`PLM_old_1` .. `PLM_old_4`), scrambled row orders, 0% part overlap, multiline explanations with Japanese quotes, recovery from missing `PLM_old`.
- **R4 Boundaries**: Short barcodes (< 9 chars), empty Row 36 and `PLM!C2`, case-insensitive 3-char matching (`1hn` vs `1HN`), whitespace stripping on service notes, missing Sheet `UNIT` in master file.
- **R5 Boundaries**: Unsupported model in JIG master, read-only file locks, merged cells handling, zero-JIG models, rejected 4M assessments.
- **R6 Boundaries**: Empty CTTT submissions, fractional quantities (`0.5`, `1.25`), non-numeric quantity strings (`"2 pcs"` coerced to 0.0), self-check without reference data, repeated submissions.

---

### Tier 3: Pairwise Combinations (`test_tier3_pairwise_combinations.py` — 6 Tests)
Validates interactions between functional subsystems in sequential pipelines:
1. `TC-T3-01`: Wizard Steps 1 & 2 <===> TC24 BOM Filtering (prunes expired items, creates clean PLM, backs up raw).
2. `TC-T3-02`: Member Submissions <===> Leader Tracking & Consolidation (detects `Q2="OK"` across all units, moves files to `phutrach/`).
3. `TC-T3-03`: Consolidated Data <===> MSI Deep Reconciliation (evaluates 9 branches across all units and Hontai).
4. `TC-T3-04`: Master Comparison Workbook <===> List JIG & 4M Confirmation (builds `form_ssbom` with `List JIG`).
5. `TC-T3-05`: BOM Filtering <===> Annotation Inheritance (filters Rev 2 tree and migrates Rev 1 annotations).
6. `TC-T3-06`: Annotation Inheritance <===> Member Workspace Re-Check (member opens updated BOM with inherited notes).

---

### Tier 4: Real-World Production Scenarios (`test_tier4_production_scenarios.py` — 5 Tests)
End-to-end execution matching factory workflows:
1. `TC-T4-01`: DMT/PMT Phase Workflow for Model `Virgo` (`maT`) with common date `2024/08/25`.
2. `TC-T4-02`: Mass Production MP Phase Workflow for Model `Libra2` (`ma1`) with individual dates.
3. `TC-T4-03`: TC2412 24-column Excel formula protection benchmark (`Tongket!C5`, `CTTT!A1`, `CTTT!I3`, `CTTT!G3`, `PLM!R2`, `PLM!S2`).
4. `TC-T4-04`: Rapid $O(N)$ UnitResolver algorithm benchmark (5,000+ nodes, 100% parity, < 100 ms).
5. `TC-T4-05`: Real-world ECN annotation migration lifecycle (500 parts, 120 notes, 20 added, 10 deleted).

---

## 4. Execution Commands

### Run Full E2E Test Suite (81 Tests)
```bash
py -m pytest tests/e2e/ -v
```

### Run by Specific Tier
```bash
# Tier 1: Feature Isolation (40 tests)
py -m pytest tests/e2e/test_tier1_feature_coverage.py -v

# Tier 2: Boundaries & Corner Cases (30 tests)
py -m pytest tests/e2e/test_tier2_boundary_corner.py -v

# Tier 3: Pairwise Combinations (6 tests)
py -m pytest tests/e2e/test_tier3_pairwise_combinations.py -v

# Tier 4: Production Workflows (5 tests)
py -m pytest tests/e2e/test_tier4_production_scenarios.py -v
```
