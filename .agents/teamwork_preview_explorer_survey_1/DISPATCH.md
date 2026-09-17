# Task Assignment: Legacy VBA & Algorithm Architecture Investigation

## Identity
- Archetype: teamwork_preview_explorer
- Role: Legacy VBA & Algorithm Investigator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Objective
Thoroughly inspect all legacy code, VBA modules, macros, Excel workbooks, and lookup sheets in `D:\Sandbox\pm_sosanhbom` to reverse-engineer and document:
1. 6-level BOM tree filter logic (Level 1..6) including validity date handling ("to", "UP", comparison with current date).
2. Model machine decomposition logic (Virgo, Libra2, Iris2024...) and rules for pruning/skipping unneeded sub-assemblies.
3. Component-to-Unit resolution algorithm (replacing the 5,619-row lookup sheet `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx` with an optimal in-memory tree traversal).
4. CTTT vs PLM vs R3 cross-reconciliation matrix: quantity discrepancies, replacement parts, additions/deletions.
5. MSI data management and 3-character fixed code logic (`fix_serial`).
6. Exact input and output data schemas of legacy forms (`form_ssbom.xlsm`, etc.).

## Inputs & Context
- Read `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md`
- Inspect existing files in `D:\Sandbox\pm_sosanhbom` (e.g. `form_ssbom.xlsm`, `.vbs` files, `.xlsx` lookup files, extracted VBA scripts)

## Deliverables
- Write comprehensive report to `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\handoff.md`
- Report back to parent orchestrator with a summary of findings.
