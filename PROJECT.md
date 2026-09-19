# Project: Phần Mềm So Sánh BOM Tự Động (Kyocera Desktop App)

## Architecture
- **GUI Layer** (`src/gui/`, `src/ui/`): PyQt6 desktop interface.
  - `LeaderView` (`src/gui/leader_view.py`): 4-step sequential Wizard workflow (Setup & Staffing, Data Sourcing, Tracking & Consolidation, Comparison & Reporting).
  - `MemberView` (`src/gui/member_view.py`): Member workspace conforming to `formnguoidung` (auto-load engineer/machine, 3 tables CTTT/MSI/7980, local self-check against BOM, Q2="OK" submission seal).
- **Core Processing Engines** (`src/core/`):
  - `BOMTreeParser`, `date_filter.py`, `model_pruner.py`: Recursive BOM PLM parsing, date effectivity pruning, and model-specific pruning (Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris) with backup to `backupTC14full/`.
  - `reconciliation.py`: 3-way line-by-line reconciliation (CTTT vs PLM vs R3), station totals, reverse missing lookup.
  - `msi_engine.py`: MSI deep cross-check against `FIX_SERIAL_DLTOOL_VER010.xls` (Unit 9-char & Machine 10-char matching, contrast styling).
  - `jig_manager.py`: JIG master catalog loader from `List JIG thay doi, khi bo sung ma hang.xlsx` across 15 machine models & 4M assessment confirmation.
  - `inheritance_engine.py` / `workbook_updater.py`: Legacy `ham_match_index_mix` algorithm, preserving Explanations, Responsible Person, Manager Check from `PLM_old` to `PLM`, archiving to `capnhat\old\`, and refreshing Pivot Tables.
- **Reporting & Services** (`src/reporting/`, `src/automation/`):
  - `excel_generator.py`: Generating `form_ssbom.xlsm` workbook with formula networks, formatting, Pivot Tables.
  - `outlook_mailer.py`: 2-tier Outlook email generation (Member notification & Management reporting).
  - `automation/sap/`: CS12 transaction automation with common or per-machine dates.
  - `automation/tc2412/`: Teamcenter Active Workspace web download automation.

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Leader Step 1: Project Setup | Select Model, Phase (maT for DMT/PMT, ma1 for PP+), input BOM codes | M1 | ORIGINAL_REQUEST §R1 |
| 2 | Leader Step 1: Staffing Matrix | Staff assignment Cơ 1, 2, 3 (Sheet Lichsu) & auto-create machine folders + member packages | M1 | ORIGINAL_REQUEST §R1, VBA Lichsu |
| 3 | Leader Step 2: PLM & SAP Sourcing | Integrated download of PLM TC24 & SAP R3 CS12 (common or individual dates), auto-routing to machine folders | M1 | ORIGINAL_REQUEST §R1 |
| 4 | Leader Step 3: Real-time Q2 Tracking | Real-time scan of member submission packages via cell CTTT!Q2 = "OK", live status table | M1 | ORIGINAL_REQUEST §R1, VBA kt_trangthai |
| 5 | Leader Step 3: Fail-Closed Consolidation | Enforce 100% OK check before consolidation, consolidate CTTT, MSI, 7980/7990, move member files to phutrach | M1 | ORIGINAL_REQUEST §R1, VBA tonghopdl |
| 6 | Leader Step 4: BOM Compare & Pivot Refresh | Generate form_ssbom workbook, populate comparison sheets, refresh Pivot Tables | M1 | ORIGINAL_REQUEST §R1, form_ssbom.xlsm |
| 7 | Leader Step 4: 2-Tier Outlook Email | Preview and send 2-tier Outlook emails (to members and to managers) with attachments | M1 | ORIGINAL_REQUEST §R1, VBA Mail |
| 8 | PLM BOM Tree Parsing | Parse Teamcenter multi-level BOM export into structured hierarchical tree | M2 | ORIGINAL_REQUEST §R2 |
| 9 | Occurrence Effectivities Date Filter | Scan column I (from...to...), compare against current date, eliminate expired items | M2 | ORIGINAL_REQUEST §R2, VBA locbomfull |
| 10 | Recursive Child Subtree Deletion | Delete all descendant child nodes (Level 1..6) when parent is expired or pruned | M2 | ORIGINAL_REQUEST §R2, VBA locbomfull |
| 11 | Model-Specific Pruner (BolocBom) | Apply BolocBom rules for Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris (exact Full_name, partial Part_name, unexpanded parts) | M2 | ORIGINAL_REQUEST §R2, VBA BolocBom |
| 12 | Backup to backupTC14full | Automatic backup of raw TC BOM export to backupTC14full folder before filtering | M2 | ORIGINAL_REQUEST §R2 |
| 13 | Multi-Version Sheet Archiving | Backup current PLM sheet as PLM_old upon new BOM ingestion | M3 | ORIGINAL_REQUEST §R3, VBA capnhat_PLM_R3 |
| 14 | ham_match_index_mix Inheritance | Match unchanged parts by Item ID and carry over Explanations (Col O), Responsible Person (Col P), Manager Check (Col Q) | M3 | ORIGINAL_REQUEST §R3, VBA capnhat_PLM_R3 |
| 15 | File Archive to capnhat\old\ | Move superseded PLM/R3 raw files into capnhat\old\ archive folder | M3 | ORIGINAL_REQUEST §R3 |
| 16 | Workbook Pivot Table Refresh | Refresh data cache and Pivot Tables in form_ssbom after updating BOM data | M3 | ORIGINAL_REQUEST §R3 |
| 17 | FIX_SERIAL Master Loader | Load FIX_SERIAL_DLTOOL_VER010.xls (Sheet UNIT and Sheet MACHINE) | M4 | ORIGINAL_REQUEST §R4, VBA msi |
| 18 | MSI Unit & Machine 3-Char Cross-Check | Match 9-char Unit MSI & Label Comment, match 10-char Machine HONTAI, evaluate OK/NG | M4 | ORIGINAL_REQUEST §R4, VBA msi |
| 19 | MSI Contrast Visual Styling | Contrasting color styling: Red (255) for NG/Missing, Green (6750054) for OK | M4 | ORIGINAL_REQUEST §R4 |
| 20 | List JIG Master Loader | Link with 'List JIG thay doi, khi bo sung ma hang.xlsx', filter and load 15 machine series | M5 | ORIGINAL_REQUEST §R5, VBA uf_jig |
| 21 | List JIG Formula & Format Preservation | Preserve formulas and cell styles in Sheet List JIG | M5 | ORIGINAL_REQUEST §R5 |
| 22 | 4M Assessment & KTSX Confirmation | Interactive checkboxes / fields to assess 4M changes and record confirmation with KTSX | M5 | ORIGINAL_REQUEST §R5 |
| 23 | Member Assignment Auto-Loading | Auto-populate engineer name, machine code, department from assignment package, eliminate sample text boxes | M6 | ORIGINAL_REQUEST §R6, formnguoidung.xlsm |
| 24 | Member 3-Table Input & Validation | Input and validation for 3 tables: CTTT, MSI, Label 7980/7990 | M6 | ORIGINAL_REQUEST §R6, formnguoidung.xlsm |
| 25 | Member Preliminary Self-Check | Auto-load machine BOMs (PLM & R3) for engineer self-check before submission | M6 | ORIGINAL_REQUEST §R6 |
| 26 | Member Submission Seal (Q2 = OK) | "Xác nhận Nộp" button to stamp CTTT!Q2 = "OK" green status seal for Leader tracking | M6 | ORIGINAL_REQUEST §R6, VBA uf_ssbnpt |
| 27 | E2E Testing Suite (Tiers 1-4) | Comprehensive 81+ test case opaque-box test suite covering all features and boundaries | E2E | ORIGINAL_REQUEST §Acceptance Criteria |
| 28 | Final Hardening & App Packaging | Tier 5 adversarial testing, package_app.py build verification, launcher health check | M7 | ORIGINAL_REQUEST §Acceptance Criteria |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| E2E | E2E Testing Track | Independent requirement-driven test suite (Tiers 1-4) & TEST_READY.md | none | DONE (81/81 pass) |
| M1 | Leader Workspace 4-Step Wizard | Refactor Leader Workspace into sequential 4-step wizard with full automation | none | DONE (62/62 pass) |
| M2 | BOM Filter Engine Level 1..6 | Recursive date effectivity & BolocBom model pruner with backupTC14full | none | DONE (verified) |
| M3 | Explanation Inheritance Pipeline | ham_match_index_mix algorithm, PLM_old archiving, capnhat\old\, Pivot refresh | M2 | DONE (16/16 pass) |
| M4 | MSI Deep Cross-Check Engine | FIX_SERIAL_DLTOOL_VER010.xls 3-char match, contrast styling, UI integration | M1 | DONE (verified) |
| M5 | JIG Master Catalog & 4M Assessment | Load List JIG master for 15 series, preserve formulas, 4M evaluation with KTSX | M1 | DONE (11/11 pass) |
| M6 | Member Workspace (formnguoidung) | Auto-load assignments, 3 tables, self-check, Q2="OK" submission seal | M1 | DONE (28/28 pass) |
| M7 | Final E2E Pass, Hardening & Packaging | Pass 100% E2E tests, Tier 5 adversarial tests, package_app.py build | E2E, M1..M6 | DONE (741/741 pass) |

## Interface Contracts
### Leader Wizard ↔ Sourcing & File System
- Step 1 creates directory structure: `<Project_Dir>/<Machine_Code>/` and `<Project_Dir>/<Machine_Code>/<Engineer_Name>.xlsm` (or .xlsx) copied from template `formnguoidung.xlsm`.
- Step 2 downloads or copies:
  * PLM: `<Project_Dir>/<Machine_Code>/PLM_<Machine_Code>.xlsx`
  * SAP R3: `<Project_Dir>/<Machine_Code>/R3_<Machine_Code>.xlsx`
- Step 3 scans: `<Project_Dir>/<Machine_Code>/*` for sheet `CTTT`, cell `Q2`. Returns `Status: OK` if `Q2 == "OK"`, otherwise `Pending`.
  * Consolidation moves submitted files to `<Project_Dir>/phutrach/<Machine_Code>/`.
- Step 4 outputs: `<Project_Dir>/form_ssbom_<Machine_Code>.xlsm`.

### Member Workspace ↔ Submission File
- Member opens `<Machine_Code>/<Engineer_Name>.xlsm`.
- Auto-populates: `Engineer Name`, `Machine Code`, `Department` into header cells.
- Tables: `CTTT` (Cols A:N), `MSI` (Cols A:H), `Label_7980_7990` (Cols A:F).
- "Xác nhận Nộp" action writes `"OK"` into `CTTT!Q2`, applies green background, and saves file.

### BOM Filter Engine ↔ Output
- Input: Raw TC14/TC24 Excel/CSV export.
- Output: Standardized PLM BOM dataframe / Excel sheet with expired dates pruned and children (Level 1..6) deleted recursively. Raw file backed up to `backupTC14full/`.

### Explanation Inheritance (`ham_match_index_mix`)
- Input: `PLM_old` sheet (Cols C: Item ID, O: Explanation, P: Person, Q: Manager Check), new `PLM` sheet.
- Output: Updated `PLM` sheet preserving O, P, Q for matching Item IDs. Old file moved to `capnhat\old\`. Pivot Table cache refreshed.

### MSI Engine ↔ FIX_SERIAL Master
- Input: `FIX_SERIAL_DLTOOL_VER010.xls`, `Sheet PLM` and `Sheet MSI`.
- Logic: Match 9 chars on `UNIT!A`, 10 chars on `MACHINE!A`.
- Output: Verdict `OK` / `NG`, styling Red `Color=255`, Green `Color=6750054`.

### JIG Manager ↔ Master Catalog & 4M
- Input: `List JIG thay doi, khi bo sung ma hang.xlsx`.
- Output: Populated Sheet `List JIG` for selected machine model with preserved Excel formulas and 4M checklist (Có/Không, OK/NG).

## Code Layout
- `src/gui/leader_view.py`: Leader 4-Step Wizard UI & Controller.
- `src/gui/member_view.py`: Member Workspace UI & Controller.
- `src/core/models.py`: Data models for Staffing, BOM, MSI, JIG, 4M.
- `src/core/tree_parser.py`: BOM hierarchical parser.
- `src/core/date_filter.py`: Column I date effectivity filter.
- `src/core/model_pruner.py`: BolocBom rule-based pruner.
- `src/core/reconciliation.py`: 3-way reconciliation engine.
- `src/core/msi_engine.py`: MSI deep cross-check engine.
- `src/core/jig_manager.py`: JIG catalog loader & 4M assessment manager.
- `src/core/inheritance_engine.py`: ham_match_index_mix & multi-version updater.
- `src/reporting/excel_generator.py`: form_ssbom workbook generator.
- `src/reporting/outlook_mailer.py`: 2-tier Outlook mailer.
- `tests/e2e/`: E2E test suite (Tiers 1-4).
- `scripts/package_app.py`: Application packager.
