# Task Assignment: Explorer Fix 2 (Automation & Provider Adapters Remediation)

## Identity
- Archetype: teamwork_preview_explorer
- Role: Automation & Provider Adapters Remediation Investigator
- Working Directory: D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_2
- Parent Conversation ID: 8a26cf43-3f4f-42ea-ac18-3875de8c9a43

## Context
Gate Iteration 1 FAILED due to Forensic Auditor INTEGRITY VIOLATION, Reviewer REQUEST_CHANGES, and Challenger DEFECTS_DETECTED.
You MUST read:
- `D:\Sandbox\pm_sosanhbom\.agents\ORIGINAL_REQUEST.md` (authoritative user requirements, especially Section R1, R2, and R5 Provider Adapters)
- `D:\Sandbox\pm_sosanhbom\PROJECT.md`
- `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_orchestrator_1\DEAD_ENDS.md`
- `D:\Sandbox\pm_sosanhbom\.agents\auditor_1\handoff.md` (Finding 1.3 Group A: TC14 NameError)
- `D:\Sandbox\pm_sosanhbom\.agents\reviewer_1\handoff.md` (CRIT-01)
- `D:\Sandbox\pm_sosanhbom\.agents\reviewer_2\handoff.md` (F-03, F-04 R5 Adapters, F-05 import)
- `D:\Sandbox\pm_sosanhbom\.agents\challenger_2\handoff.md` (Adapter fault injection findings)

## Objectives & Scope
Investigate and design exact remediation code changes for:
1. **Audit Finding 1.3 Group A — TC14 Client Syntax/Scoping Bug**:
   - Inspect `src/automation/tc14/client.py:133-170` (and `apps/1.0.0/src/automation/tc14/client.py` mirror).
   - Design exact fix for `login()`: add `timeout: Optional[float] = None` parameter and properly initialize `wait = WebDriverWait(driver, op_timeout)` before line 164.
2. **Reviewer Finding F-04 — Implement R5 Provider Adapters**:
   - Design `src/core/adapters.py` (or `src/services/providers.py`):
     - `PLMProvider` (ABC) with methods `fetch_bom(item_id, rev)` and `export_excel(item_id, rev)`.
     - `TeamcenterSeleniumAdapter(PLMProvider)` wrapping `TC14AutomationClient`.
     - `ExcelPLMAdapter(PLMProvider)` reading local PLM Excel files directly.
     - `ERPProvider` (ABC) with methods `fetch_multilevel_bom(material, plant, date)`.
     - `SAPR3COMAdapter(ERPProvider)` wrapping `CS12Service`.
     - `ExcelR3Adapter(ERPProvider)` reading local R3 exports.
     - Design unit tests demonstrating core engine independence by swapping providers.
3. **Challenger 2 Finding — Adapter Fault Blindness**:
   - Inspect `TC14SessionManager.is_session_alive()` and `SAPConnectionManager.is_connected()`. Design reconnection logic on dropped sockets/COM severance.
4. **Reviewer Finding F-05 — Minor Import Bug**:
   - In `src/gui/leader_view.py:100`, fix `from src.automation.sap.parser import SAPBOMParser` to use `ResilientR3Parser` or `parse_r3_cs12_file`.

Write your comprehensive remediation report to `D:\Sandbox\pm_sosanhbom\.agents\explorer_fix_2\handoff.md` and message parent orchestrator.
