# Dead Ends Tracking — BOM Comparison Automation Modernization

| Iteration | Approach Tried | Why It Failed | Files Touched |
|---|---|---|---|
| Iteration 1 | Creating a secondary GUI facade (`src/ui/main_window.py`) that returns hardcoded dictionary metrics (`total_parts: 1248, ok_count: 1240, ng_count: 8`) and launches from `SSBOM_Launcher.py` | Rejected by Forensic Auditor as Prohibited Pattern #2 (Facade Implementation) and #1 (Hardcoded Test Results). Excel report was never created on disk. | `src/ui/main_window.py`, `SSBOM_Launcher.py`, `apps/1.0.0/SSBOM_App.py` |
| Iteration 1 | Writing Tier 1 feature tests in `tests/tier1_features/` that define local reference mock functions instead of importing from `src/` | Rejected by Forensic Auditor as Prohibited Pattern #4 (Self-Certifying Tests). Failed audit with 0 src imports across 10 test files. | `tests/tier1_features/test_f07..f10, f22..f26, f28` |
| Iteration 1 | Limiting hierarchy level validation in `BOMNode` to `le=10` | Crashes with Pydantic ValidationError when parsing deep industrial BOMs (11 and 12 levels). | `src/core/models.py` |
| Iteration 1 | Undefined `timeout` parameter and uninitialized `wait` variable in `TC14AutomationClient.login()` | Crashes at runtime with `NameError: name 'timeout' is not defined` on login. | `src/automation/tc14/client.py` |
| Iteration 1 | Missing cycle detection in recursive date filter and model pruner | Crashes with RecursionError when encountering cyclic BOM occurrence data. | `src/core/date_filter.py`, `src/core/model_pruner.py` |
