"""
Unit tests for Milestone M1 Specification Contract Verification
specs/SPEC_PLM_AUTO_DOWNLOAD.md
"""
import re
from pathlib import Path
import pytest
from bs4 import BeautifulSoup
import lxml.html

WORKSPACE = Path(__file__).resolve().parents[2]
SPEC_PATH = WORKSPACE / "specs" / "SPEC_PLM_AUTO_DOWNLOAD.md"
ORIGINAL_REQUEST_PATH = WORKSPACE / ".agents" / "ORIGINAL_REQUEST.md"

EXPECTED_14_COLUMNS = [
    "Home",
    "Level",
    "Item Type",
    "Item Id",
    "Has Children",
    "Quantity",
    "1st Parts",
    "2nd BOM Flag",
    "Occurrence Effectivities",
    "Item Revision Project List",
    "Item Name",
    "Notice No",
    "Revision",
    "Item Rev Status",
]


def test_spec_14_columns_contract():
    """Verify that specs/SPEC_PLM_AUTO_DOWNLOAD.md defines the exact 14 columns matching ORIGINAL_REQUEST.md."""
    assert SPEC_PATH.exists(), f"Spec file missing: {SPEC_PATH}"
    assert ORIGINAL_REQUEST_PATH.exists(), f"Original request missing: {ORIGINAL_REQUEST_PATH}"

    spec_text = SPEC_PATH.read_text(encoding="utf-8")
    orig_text = ORIGINAL_REQUEST_PATH.read_text(encoding="utf-8")

    # 1. Table 3 extraction
    table_cols = re.findall(r'\|\s*\*\*\d+\*\*\s*\|\s*\*\*[A-N]\*\*\s*\|\s*\*\*([^*]+)\*\*', spec_text)
    assert len(table_cols) == 14, f"Expected 14 columns in Table 3, got {len(table_cols)}"
    assert table_cols == EXPECTED_14_COLUMNS, f"Table 3 columns mismatch: {table_cols} vs {EXPECTED_14_COLUMNS}"

    # 2. CANONICAL_14_COLUMNS extraction
    canonical_match = re.search(r'CANONICAL_14_COLUMNS\s*:\s*List\[str\]\s*=\s*\[(.*?)\]', spec_text, re.DOTALL)
    assert canonical_match, "CANONICAL_14_COLUMNS not found in code"
    canonical_cols = re.findall(r'"([^"]+)"', canonical_match.group(1))
    assert canonical_cols == EXPECTED_14_COLUMNS, f"Code CANONICAL_14_COLUMNS mismatch: {canonical_cols}"

    # 3. EXPECTED_HEADERS extraction
    expected_match = re.search(r'EXPECTED_HEADERS\s*=\s*\[(.*?)\]', spec_text, re.DOTALL)
    assert expected_match, "EXPECTED_HEADERS not found in openpyxl code"
    expected_cols = re.findall(r'"([^"]+)"', expected_match.group(1))
    assert expected_cols == EXPECTED_14_COLUMNS, f"EXPECTED_HEADERS mismatch: {expected_cols}"

    # 4. AC-CONFIG-04 numbered list extraction
    ac_config_cols = re.findall(
        r'^\s*(\d+)\.\s+([A-Za-z0-9\s]+?)\s*$',
        spec_text[spec_text.find("AC-CONFIG-04"):spec_text.find("AC-CONFIG-05")],
        re.MULTILINE
    )
    ac_config_sorted = [name.strip() for num, name in sorted([(int(n), s) for n, s in ac_config_cols], key=lambda x: x[0])]
    assert ac_config_sorted == EXPECTED_14_COLUMNS, f"AC-CONFIG-04 mismatch: {ac_config_sorted}"

    # 5. Check against ORIGINAL_REQUEST.md
    orig_cols = re.findall(r'^\s*(\d+)\.\s+([A-Za-z0-9\s]+?)\s*$', orig_text, re.MULTILINE)
    orig_14 = [name.strip() for num, name in sorted([(int(n), s) for n, s in orig_cols if 1 <= int(n) <= 14], key=lambda x: x[0])]
    assert orig_14 == EXPECTED_14_COLUMNS, f"ORIGINAL_REQUEST.md mismatch: {orig_14}"


def test_spec_all_phases_user_stories_and_gwt():
    """Verify all 7 phases + R8 have User Stories, complete Given-When-Then blocks, and schemas."""
    spec_text = SPEC_PATH.read_text(encoding="utf-8")

    phases = [
        ("Phase 1: AUTH-01", r'## 4\.\s*PHASE 1:\s*AUTH-01', "AC-AUTH-", ["AuthInputSchema", "AuthOutputSchema"], 5),
        ("Phase 2: BOM-SEARCH-01", r'## 5\.\s*PHASE 2:\s*BOM-SEARCH-01', "AC-SEARCH-", ["SearchInputSchema", "SearchOutputSchema"], 4),
        ("Phase 3: BOM-EXPAND-01", r'## 6\.\s*PHASE 3:\s*BOM-EXPAND-01', "AC-EXPAND-", ["ExpandInputSchema", "ExpandOutputSchema"], 4),
        ("Phase 4: BOM-SELECT-01", r'## 7\.\s*PHASE 4:\s*BOM-SELECT-01', "AC-SELECT-", ["SelectInputSchema", "SelectOutputSchema"], 3),
        ("Phase 5: BOM-EXPORT-OPEN-01", r'## 8\.\s*PHASE 5:\s*BOM-EXPORT-OPEN-01', "AC-EXPORT-OPEN-", ["ExportOpenInputSchema", "ExportOpenOutputSchema"], 2),
        ("Phase 6: BOM-EXPORT-CONFIG-01", r'## 9\.\s*PHASE 6:\s*BOM-EXPORT-CONFIG-01', "AC-CONFIG-", ["ExportConfigInputSchema", "ExportConfigOutputSchema"], 5),
        ("Phase 7: BOM-EXPORT-RUN-01", r'## 10\.\s*PHASE 7:\s*BOM-EXPORT-RUN-01', "AC-RUN-", ["ExportRunInputSchema", "ExportRunOutputSchema"], 5),
        ("Phase 8: REP-01 (R8)", r'## 11\.\s*PHASE 8:\s*REP-01', "AC-REP-", ["ProgressReporter"], 3),
    ]

    total_ac_count = 0

    for name, regex, ac_prefix, schemas, expected_count in phases:
        start_m = re.search(regex, spec_text)
        assert start_m, f"Header for {name} not found"
        pos = start_m.start()

        next_m = re.search(r'\n## \d+\. ', spec_text[pos + 10:])
        section_text = spec_text[pos:pos + 10 + next_m.start()] if next_m else spec_text[pos:]

        # 1. User Story check
        assert "###" in section_text and "User Story" in section_text, f"{name} is missing User Story section"

        # 2. Acceptance criteria check
        ac_titles = re.findall(rf'-\s*\*\*({ac_prefix}\d+[^*]*)\*\*:', section_text)
        assert len(ac_titles) == expected_count, f"{name} expected {expected_count} ACs, found {len(ac_titles)}"

        for title in ac_titles:
            escaped = re.escape(title)
            ac_block_m = re.search(rf'-\s*\*\*{escaped}\*\*:\s*\r?\n(.*?)(?=\r?\n-\s*\*\*|\r?\n\r?\n###|\Z)', section_text, re.DOTALL)
            assert ac_block_m, f"AC block not found for {title}"
            content = ac_block_m.group(1)
            assert "**Given**" in content or "**given**" in content.lower(), f"{title} missing Given"
            assert "**When**" in content or "**when**" in content.lower(), f"{title} missing When"
            assert "**Then**" in content or "**then**" in content.lower(), f"{title} missing Then"
            total_ac_count += 1

        # 3. Schemas check
        for schema in schemas:
            assert schema in section_text, f"{name} missing schema: {schema}"

    assert total_ac_count == 31, f"Expected 31 total ACs across all phases, verified {total_ac_count}"


def test_spec_dom_selectors_in_snapshots():
    """Verify core DOM selectors from spec actually match elements in available DOM snapshots using BeautifulSoup & XPath."""
    login_html = (WORKSPACE / "tc14_login_page_dom.html").read_text(encoding="utf-8", errors="replace")
    export_html = (WORKSPACE / "tc14_export_dialog_dom.html").read_text(encoding="utf-8", errors="replace")

    login_soup = BeautifulSoup(login_html, "html.parser")
    export_soup = BeautifulSoup(export_html, "html.parser")
    login_tree = lxml.html.fromstring(login_html)
    export_tree = lxml.html.fromstring(export_html)

    # 1. Verify Login Page Locators (Teamcenter 2412)
    ver_elem = login_soup.select("p.aw-login-copyrightTitle")
    assert len(ver_elem) > 0, "p.aw-login-copyrightTitle not found"
    assert "Version 2412" in ver_elem[0].get_text()
    assert len(login_soup.select('input[name="username"]')) == 1
    assert len(login_soup.select('input[name="password"]')) == 1
    assert len(login_soup.select('button[type="submit"]')) == 1
    assert len(login_soup.select('div.aw-login-progressContainer')) == 1
    assert len(login_tree.xpath("//div[contains(@class,'signin-form')]")) >= 1

    # 2. Verify Export Dialog Locators
    assert len(export_soup.select('div.sw-popup.sw-dialog#ui-id-3')) == 1
    assert len(export_soup.select('button[command-id="Awp0CloseCommandPanel"]')) == 1
    assert len(export_soup.select('button[command-id="Awp0MoveUpExcelColumn"]')) == 1
    assert len(export_soup.select('button[command-id="Awp0MoveDownExcelColumn"]')) == 1
    assert len(export_soup.select('button[command-id="Awp0ExportSelectedColumnsAdd"]')) == 1
    assert len(export_soup.select('details[caption="Selected Properties"] li.aw-widgets-cellListItem')) == 6
    assert len(export_soup.select('form.sw-command-panel button.sw-button')) >= 1

    # 3. Verify Toolbar / Global Command IDs in Active Workspace
    assert len(export_soup.select('button[command-id="Awb0Expand"]')) == 1
    assert len(export_soup.select('button[command-id="Awp0SelectAll"]')) == 1
    assert len(export_soup.select('button[command-id="Arm0ExportImport"]')) == 1
    assert len(export_soup.select('button[command-id="Awp0ExportToExcel"]')) == 1
    assert len(export_soup.select('button[command-id="Awp0ShowAlertWithBubble"]')) == 1
