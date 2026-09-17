# Handoff Report: Teamcenter TC14 Web Automation Investigation

**Explorer**: Explorer 2 (Teamcenter TC14 Web Automation Investigator)  
**Working Directory**: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_2`  
**Target System**: Siemens Teamcenter Active Workspace (TC14) at `http://tcmp3gwb:3000/`  
**Date**: 2026-09-17  
**Type**: Hard Handoff (Investigation Complete)

---

## 1. Observation

### 1.1 Target Endpoint & Architecture
- **URL**: `http://tcmp3gwb:3000/`
- **HTTP Probe Result**: Status `200 OK`, Server returns HTML bundle for Siemens Active Workspace Client.
- **Client Build & Version**: Teamcenter Active Workspace Version 2412 (Build December 2024, Copyright © 2024 Siemens Industry Software Inc.).
- **Application Type**: Single Page Application (SPA) built with React / Declarative UI (`afx-app-tcaw`, `aw-layout-mainView`, `dynamic-declreact.fab15efbfbd1c74.js`, `react-dom.production.min.fab06efbfbdefbf.js`).
- **Initial Cookies**:
  - `_csrf`: Anti-CSRF token (Path=/, HttpOnly, SameSite=Lax).
  - `XSRF-TOKEN`: Client CSRF token for SOA requests.
  - `JSESSIONID`: Java Servlet Session ID (domain=tcmp3gwb).
- **LocalStorage Keys (Initial)**:
  - `['locale:/', 'sessionID:/', 'soaService.timeOfLastCall:/', 'theme:/']`

### 1.2 Authentication Mechanism & Selectors
- **Form Selectors**:
  - Username Input: `input[name="username"]` (Attributes: `placeholder="User Name"`, `data-locator="User Name"`, `class="sw-property-val"`, `aria-label="username"`).
  - Password Input: `input[name="password"]` (Attributes: `placeholder="Password"`, `data-locator="Password"`, `class="sw-property-val"`, `aria-label="password"`).
  - Login Button: `div.aw-login-signInButton button[type="submit"]` (Classes: initially `sw-button accent-caution disabled`, transitions to `sw-button accent-caution` once inputs are filled).
- **Authentication Credentials**: `vn_pe03` / `vn_pe03` (Verified working).
- **Post-Login State**:
  - Destination URL: `http://tcmp3gwb:3000/#/showHome`
  - Page Title: `Teamcenter - Home`
  - User Division: `tcBanner_vn_pe03__VN Manufacturing Engineer Division`
  - Active Cookies (9 total): `sanEID`, `sanSID`, `sanOOC`, `JSESSIONID`, `_csrf`, `XSRF-TOKEN`, `sanEID-expire`, `sanSID-expire`, `sanOOC-expire`.
  - Session Tracker in LocalStorage: `soaService.timeOfLastCall:/`, `awSession:/`, `sessionDiscriminator:/`, `sessionID:/`.

### 1.3 Search & Navigation Flow
- **Global Search Input**:
  - Selector: `input.aw-uiwidgets-searchBox[name="searchBox"]` (Parent: `div.aw-uiwidgets-searchBoxContainer`).
  - Search Execution: Enter text + `Keys.ENTER` triggers hash navigation.
- **Direct Hash URL Pattern**:
  - Direct Search Route: `http://tcmp3gwb:3000/#/teamcenter.search.search?searchCriteria=<PART_NUMBER>&secondaryCriteria=*&isGlobalSearch=true`
  - Verified with part `110C103NL0`: Direct navigation immediately returned 3 matching items:
    1. `110C103NL0` / `PBA-00001711` / `Revision:A` (Product Revision)
    2. `110C103NL0` / `RBA-00003025` / `Revision:A`
    3. `110C103NL0` / `MBC_90284614` / `Revision:04`
- **Search Result Item Selectors**:
  - Result Container: `li.aw-widgets-cellListItem`
  - Clickable Title Link: `a.aw-widgets-cellListCellTitle[data-locator="aw-clickable-title"]`
  - Inline Open Button: `button.aw-commandId-Awp0ShowObjectCell[command-id="Awp0ShowObjectCell"]`
- **Object View Navigation**:
  - Target URL Pattern: `http://tcmp3gwb:3000/#/com.siemens.splm.clientapi.tcui.xrt.showObject?uid=<OBJECT_UID>`
  - Verified Page Title: `Teamcenter - PBA-00001711/A;1-110C103NL0;ECOSYS MA4500ifx 220-240V50/60HZ`
  - Available Object Tabs: `Overview`, `Content`, `Classification`, `Where Used`, `Attachments`, `Changes`, `History`, `Relations`, `Reports`, `Audit Logs`.

### 1.4 BOM Tree Hierarchy & Export Mechanisms
- **Content Tab (`sw-tab` with text "Content")**:
  - Renders the occurrence tree table: `div.aw-splm-table`, `occTreeTable`.
  - Tree Rows: `<div class="aw-splm-tableRow ui-grid-row aw-splm-tablePinnedRow" aria-level="<1..6>" aria-expanded="true|false">`.
  - Expand/Collapse Toggle: `div.aw-jswidgets-treeExpandCollapseCmd` with `title="Show Children"` (icon `miscCollapsedTree`) or `title="Hide Children"` (icon `miscExpandedTree`).
  - Table Columns: `Element`, `ID`, `Type`, `Assembly Indicator` (represents `hasChildren` True/False), `Object`, `Release Status`, `Design Item`, `Owner`.
- **Export To Excel (`button[command-id="Awp0ExportToExcel"]`)**:
  - Opens right slide-out popup: `div.sw-popup.sw-right-dialog` / `div.aw-layout-popup`.
  - Action Button: `button.sw-button` with text `Export`.
  - Downloads an `.xlsx` file generated directly by Teamcenter SOA backend.
  - Verified: Downloaded workbook is valid OpenXML (`.xlsx`), containing exported structure columns.
- **Legacy Macro Data Ground Truth**:
  - Inspecting `locbomfull.bas.bas` (lines 43-188) and PPTX slide 27 confirms the exact columns expected by downstream filters:
    - Col B: Level BOM (integer 1 to 6)
    - Col D: Part Number / Item ID
    - Col E: `hasChildren` ("True" or "False" / Assembly Indicator)
    - Col I: Effectivity date string (`chuoihieuluc`, containing "to" or "UP")
    - Col K: Item Name

### 1.5 Browser Drivers on System
- **Microsoft Edge**:
  - Path: `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
  - Version: `153.0.4234.32`
  - Driver: Selenium Manager auto-provisions matching `msedgedriver.exe`.
- **Google Chrome**:
  - Path: `C:\Users\tvn183660\AppData\Local\Google\Chrome\Application\chrome.exe`
  - Version: `152.0.7977.83`
  - Driver: Selenium Manager auto-provisions matching `chromedriver.exe`.
- **Headless Execution Verification**:
  - Both Edge and Chrome initialize and run headlessly with `--headless=new` without UI popups or crashing.
  - Chrome DevTools Protocol (`Page.setDownloadBehavior`) verified to enable automatic background file downloads.

---

## 2. Logic Chain

1. **SPA Route Determinism**:
   - *Observation*: Active Workspace uses hash routing (`#/showHome`, `#/teamcenter.search.search?...`, `#/com.siemens.splm.clientapi.tcui.xrt.showObject?...`).
   - *Deduction*: Instead of relying solely on DOM clicks and keystroke delays in search inputs, the automation driver can navigate directly via URL hash parameters. This eliminates input lag, autocomplete popups, and race conditions.

2. **BOM Full Acquisition Strategy: Export vs. Interactive Tree Scraping**:
   - *Observation A*: In the Content tab, large printer assemblies have thousands of parts virtualized in `aw-splm-tableRow`. Clicking every `miscCollapsedTree` toggle across 6 levels requires hundreds of sequential AJAX/SOA requests and DOM re-renders.
   - *Observation B*: Teamcenter provides built-in `Awp0ExportToExcel` and standard PLM reports which generate the complete BOM Full in `.xlsx` format.
   - *Observation C*: The entire legacy downstream pipeline (`locbomfull.bas`, `tonghop_new12052026_ma1.xlsm`) expects an Excel file named `PLM_*.xlsx` with columns B (Level), D (Item ID), E (hasChildren), I (Effectivity), K (Item Name).
   - *Deduction*: The primary automated path should trigger the TC14 Export To Excel / PLM Report action, downloading `PLM_<machine>.xlsx` directly into the designated export folder. An interactive tree expander module can serve as a secondary fallback if the export command is restricted by role.

3. **Session Lifecycle & Auto-Relogin Architecture**:
   - *Observation*: TC14 maintains session activity via `soaService.timeOfLastCall:/` and cookies `JSESSIONID`, `XSRF-TOKEN`. When a session times out, the SPA redirects to `http://tcmp3gwb:3000/` and renders `input[name="username"]`.
   - *Deduction*: The automation framework must implement a fail-safe session guard: before any operation, it checks `driver.current_url` and verifies that the username field is absent. If the login form is detected, it automatically executes the re-login routine with `vn_pe03 / vn_pe03` and resumes the queued task.

4. **Headless Download Reliability**:
   - *Observation*: Standard headless Chromium blocks or stores downloads as temporary files (`.tmp` / `.crdownload`) if download behavior is not explicitly permitted.
   - *Deduction*: When initializing the WebDriver, the module must configure Chrome/Edge experimental prefs (`download.default_directory`, `prompt_for_download=False`) AND execute CDP command `Page.setDownloadBehavior` with `"behavior": "allow"`.

---

## 3. Caveats

1. **Role/Rule Variations**:
   - Slide 16 mentioned: "Khi download PLM phải sử dụng tài khoản VT_KTCT1 và rule KTCT". In our live test on `http://tcmp3gwb:3000/`, credentials `vn_pe03 / vn_pe03` logged in with group/division `VN Manufacturing Engineer Division` and successfully accessed the BOM and export features. If specific models require custom closure rules (e.g., "Skip inactive occurrences"), those closure rules can be set via SOA or UI settings.
2. **Virtual DOM Scrolling in Fallback Scraper**:
   - If interactive tree scraping is used rather than Excel export, the virtual scroll container (`aw-splm-table`) only mounts DOM nodes visible in the viewport. Virtual scrolling simulation would be needed to traverse all rows.
3. **Network Latency**:
   - The TC14 server runs on the local intranet (`tcmp3gwb:3000`). While network ping is fast (<5ms), complex SOA responses (such as expanding multi-thousand part assemblies) can take 5-15 seconds. Explicit waits must be set with a minimum timeout of 30-45 seconds for heavy BOM operations.

---

## 4. Conclusion & Architectural Recommendation

We recommend implementing the **Teamcenter TC14 Web Automation Module (`tc14_client.py`)** with the following production architecture:

### 4.1 Class Design
```
┌────────────────────────────────────────────────────────┐
│                   TC14AutomationClient                 │
├────────────────────────────────────────────────────────┤
│ - driver: WebDriver (Edge or Chrome via Factory)       │
│ - base_url: "http://tcmp3gwb:3000/"                    │
│ - download_dir: Path                                   │
├────────────────────────────────────────────────────────┤
│ + login(username="vn_pe03", password="vn_pe03")        │
│ + ensure_authenticated() -> bool                       │
│ + search_part(part_number: str) -> List[SearchResult]   │
│ + open_item_revision(uid or index=0)                   │
│ + export_bom_full(output_path: Path) -> Path           │
│ + expand_bom_tree_interactive() -> List[BOMNode]       │
│ + close()                                              │
└────────────────────────────────────────────────────────┘
```

### 4.2 Core Configuration Blueprint
- **Driver**: Edge WebDriver by default (`webdriver.Edge`), with Chrome fallback (`webdriver.Chrome`).
- **Headless Mode**: `--headless=new`, `--window-size=1920,1080`, `--disable-gpu`, `--no-sandbox`.
- **CDP Download Configuration**:
  ```python
  driver.execute_cdp_cmd("Page.setDownloadBehavior", {
      "behavior": "allow",
      "downloadPath": str(download_dir.resolve())
  })
  ```
- **Selectors Dictionary**:
  ```python
  TC14_SELECTORS = {
      "username": (By.CSS_SELECTOR, "input[name='username']"),
      "password": (By.CSS_SELECTOR, "input[name='password']"),
      "login_btn": (By.CSS_SELECTOR, ".aw-login-signInButton button"),
      "search_box": (By.CSS_SELECTOR, "input.aw-uiwidgets-searchBox"),
      "cell_item": (By.CSS_SELECTOR, "li.aw-widgets-cellListItem"),
      "open_btn": (By.CSS_SELECTOR, "button[command-id='Awp0ShowObjectCell']"),
      "content_tab": (By.XPATH, "//li[contains(@class,'sw-tab') and .//span[text()='Content']]"),
      "export_excel_cmd": (By.CSS_SELECTOR, "button[command-id='Awp0ExportToExcel']"),
      "export_dialog_btn": (By.XPATH, "//div[contains(@class,'sw-right-dialog')]//button[normalize-space()='Export']"),
      "progress_spinner": (By.CSS_SELECTOR, ".aw-layout-progressIndicator, .sw-progress-indicator"),
      "tree_rows": (By.CSS_SELECTOR, ".aw-splm-tableRow.aw-splm-tablePinnedRow"),
      "tree_expand_btn": (By.CSS_SELECTOR, "div.aw-jswidgets-treeExpandCollapseCmd[title='Show Children']"),
  }
  ```

---

## 5. Verification Method

To independently verify the findings in this report:

1. **Verify WebDriver Initialization**:
   Run:
   ```powershell
   python .agents/teamwork_preview_explorer_survey_2/probe_env.py
   ```
   *Expected*: Both Edge and Chrome initialize successfully with `browserVersion` ~152-153.

2. **Verify Headless Login & Session Creation**:
   Run:
   ```powershell
   python .agents/teamwork_preview_explorer_survey_2/test_tc14_login.py
   ```
   *Expected*: URL transitions to `http://tcmp3gwb:3000/#/showHome`, page title `Teamcenter - Home`, 9 cookies issued.

3. **Verify Search & BOM Navigation**:
   Run:
   ```powershell
   python .agents/teamwork_preview_explorer_survey_2/test_tc14_content_and_export.py
   ```
   *Expected*: Search for `110C103NL0` returns ECOSYS MA4500ifx item revision, Content tab displays 52 BOM rows with tree indicators, and Export to Excel triggers valid `.xlsx` download.

4. **Invalidation Conditions**:
   - If `http://tcmp3gwb:3000/` returns HTTP 403 / 502 / Connection Refused.
   - If Active Workspace UI theme changes and removes `sw-property-val` or `aw-commandId-*` selectors.
   - If credentials `vn_pe03` are deactivated on Active Directory / Teamcenter infodba.
