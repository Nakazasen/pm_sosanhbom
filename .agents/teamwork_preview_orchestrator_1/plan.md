# Orchestrator Execution Plan: BOM Comparison Automation Rewrite

## Objective
Modernize the legacy BOM Comparison system into a modular, production-grade Python solution covering Teamcenter Active Workspace (TC14) automation, SAP R3 CS12 BOM automation, high-performance BOM tree filter/reconciliation engine, and PyQt6 desktop interface, verified against legacy VBA/VBScript test cases.

## Phase 0: Survey & Scope Mapping
1. Dispatch 3 parallel Explorers:
   - Explorer 1 (Legacy Extraction & Algorithm Specs): Inspect legacy VBA/VBS files (`form_ssbom.xlsm`, `tudongdangnhapR3.vbs`, lookup tables, module macros) to extract business logic, unit resolution, 6-level filter rules, and reconciliation matrix.
   - Explorer 2 (Teamcenter Web Automation Specs): Inspect Teamcenter Active Workspace interaction, URL `http://tcmp3gwb:3000/`, authentication, DOM structures, React SPA behaviors, and session resilience requirements.
   - Explorer 3 (SAP R3 Scripting & Integration Specs): Inspect SAP GUI scripting COM interface, CS12 parameters, plant 2200, BOM usage pp01, and export file management.
2. Merge Explorer reports into `PROJECT.md` Feature Inventory & Architecture.

## Phase 1: Decomposition & Track Setup
1. Define Milestones (Target 4-6 modular milestones):
   - M1: Core BOM Data Structures, Tree Filter (6-level), and Rapid Unit Resolver Engine.
   - M2: CTTT vs PLM vs R3 Cross-Reconciliation & Discrepancy Matrix Engine.
   - M3: Teamcenter TC14 Web Automation Module (Selenium WebDriver).
   - M4: SAP R3 CS12 Automation Module (win32com / SAP GUI Scripting).
   - M5: PyQt6 Desktop GUI Application (Leader & Member Sub-modules).
   - M6: E2E Integration, 100% Verification vs Legacy Data, & Standalone Packaging + LAN Auto-Update strictly following `D:\Sandbox\MP2027\huongdansetup_autoupdate.md` (Dual distribution: Inno Setup installer + `<App>_Launcher.exe` + `.mpupdate` atomic catalog).
2. Parallel Dual Track:
   - E2E Testing Track: Setup test infrastructure, ground-truth oracle tests from `form_ssbom.xlsm`.
   - Implementation Track: Milestone execution with Sub-orchestrators and Worker/Reviewer/Challenger/Auditor loops.

## Phase 2: Execution & Monitoring
1. Monitor subagents and sub-orchestrators.
2. Heartbeat cron checks every 10 minutes.
3. Gate checks with binary Forensic Auditor integrity veto.

## Phase 3: Final Acceptance & Reporting
1. 100% E2E test passage (Tiers 1-4).
2. Tier 5 adversarial test hardening.
3. Standalone executable build verification.
4. Comprehensive human report to Sentinel.
