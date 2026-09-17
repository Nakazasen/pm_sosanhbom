"""Shared Pytest Fixtures and Mock Generators for BOM Comparison Automation Test Suite.

Provides:
- Isolated temporary filesystem workspaces.
- Synthetic 14-column PLM BOM fixtures (DataFrames and real .xlsx files).
- Synthetic SAP R3 CS12 BOM fixtures (HTML .xls and CSV/TSV).
- Simulated CTTT work instruction and MSI/Fix Serial master tables.
- High-fidelity mocks for Selenium WebDriver (TC14) and SAP GUI Scripting COM sessions.
- Ground truth rule tables for BolocBom machine models.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any, Dict, List
from unittest.mock import MagicMock

import openpyxl
import pandas as pd
import pytest


# ============================================================================
# Filesystem & Workspace Fixtures
# ============================================================================

@pytest.fixture
def temp_workspace(tmp_path: Path) -> Path:
    """Provide a dedicated, clean temporary workspace folder."""
    ws = tmp_path / "workspace"
    ws.mkdir(parents=True, exist_ok=True)
    return ws


# ============================================================================
# Synthetic PLM Data Fixtures (14-Column TC14 Specification)
# ============================================================================

@pytest.fixture
def sample_14col_plm_records() -> List[Dict[str, Any]]:
    """Generates standard multi-level 14-column TC14 BOM records.
    
    Structure:
    - Level 1: LSU UNIT (Has Children: True, Effectivity: '01-Jan-2024 UP')
      - Level 2: POLYGON MOTOR ASSY (Has Children: True, Effectivity: '01-Jan-2024 UP')
        - Level 3: MOTOR BRACKET (Has Children: False, Qty: 1.0, Effectivity: '01-Jan-2024 UP')
        - Level 3: SCREW M3X6 (Has Children: False, Qty: 4.0, Effectivity: '01-Jan-2024 UP')
      - Level 2: F-THETA LENS (Has Children: False, Qty: 2.0, Effectivity: '01-Jan-2024 UP')
    - Level 1: FUSER UNIT (Has Children: True, Effectivity: '01-Jan-2024 UP')
      - Level 2: HEATER LAMP 220V (Has Children: False, Qty: 1.0, Effectivity: '01-Jan-2024 UP')
      - Level 2: THERMISTOR SUB ASSY (Has Children: True, Effectivity: '01-Jan-2024 UP')
        - Level 3: SENSOR CHIP (Has Children: False, Qty: 1.0, Effectivity: '01-Jan-2024 UP')
    - Level 1: EXPIRED SUBASSY (Has Children: True, Effectivity: '01-Jan-2020 to 31-Dec-2021')
      - Level 2: OLD COMPONENT (Has Children: False, Qty: 1.0, Effectivity: '01-Jan-2020 to 31-Dec-2021')
    """
    return [
        {
            "row_idx": 1,
            "level": 1,
            "item_type": "Assembly",
            "item_id": "302FP93010",
            "has_children": "True",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "LSU UNIT",
            "notice_no": "ECN-1001",
            "revision": "A",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 2,
            "level": 2,
            "item_type": "Assembly",
            "item_id": "302FP94010",
            "has_children": "True",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "POLYGON MOTOR ASSY",
            "notice_no": "ECN-1001",
            "revision": "A",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 3,
            "level": 3,
            "item_type": "Part",
            "item_id": "302FP02010",
            "has_children": "False",
            "quantity": 1.0,
            "first_parts": "P1",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "MOTOR BRACKET",
            "notice_no": "ECN-1001",
            "revision": "01",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 4,
            "level": 3,
            "item_type": "Part",
            "item_id": "B1303060",
            "has_children": "False",
            "quantity": 4.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "SCREW M3X6",
            "notice_no": "ECN-1001",
            "revision": "-",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 5,
            "level": 2,
            "item_type": "Part",
            "item_id": "302FP02050",
            "has_children": "False",
            "quantity": 2.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "F-THETA LENS",
            "notice_no": "ECN-1001",
            "revision": "B",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 6,
            "level": 1,
            "item_type": "Assembly",
            "item_id": "302FP93020",
            "has_children": "True",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "FUSER UNIT",
            "notice_no": "ECN-1002",
            "revision": "02",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 7,
            "level": 2,
            "item_type": "Part",
            "item_id": "302FP04010",
            "has_children": "False",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "HEATER LAMP 220V",
            "notice_no": "ECN-1002",
            "revision": "01",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 8,
            "level": 2,
            "item_type": "Assembly",
            "item_id": "302FP94030",
            "has_children": "True",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "THERMISTOR SUB ASSY",
            "notice_no": "ECN-1002",
            "revision": "A",
            "item_rev_status": "Released",
        },
        {
            "row_idx": 9,
            "level": 3,
            "item_type": "Part",
            "item_id": "302FP04050",
            "has_children": "False",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2024 UP",
            "item_revision_projects_list": "PRJ_ECOSYS",
            "item_name": "SENSOR CHIP",
            "notice_no": "ECN-1002",
            "revision": "01",
            "item_rev_status": "Released",
        },
        # Expired subtree for date filtering testing
        {
            "row_idx": 10,
            "level": 1,
            "item_type": "Assembly",
            "item_id": "302FP99990",
            "has_children": "True",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2020 to 31-Dec-2021",
            "item_revision_projects_list": "PRJ_OLD",
            "item_name": "EXPIRED SUBASSY",
            "notice_no": "ECN-0999",
            "revision": "01",
            "item_rev_status": "Obsolete",
        },
        {
            "row_idx": 11,
            "level": 2,
            "item_type": "Part",
            "item_id": "302FP09990",
            "has_children": "False",
            "quantity": 1.0,
            "first_parts": "",
            "second_bom_flag": "",
            "occurrence_effectivities": "01-Jan-2020 to 31-Dec-2021",
            "item_revision_projects_list": "PRJ_OLD",
            "item_name": "OLD COMPONENT",
            "notice_no": "ECN-0999",
            "revision": "01",
            "item_rev_status": "Obsolete",
        },
    ]


@pytest.fixture
def sample_14col_plm_df(sample_14col_plm_records: List[Dict[str, Any]]) -> pd.DataFrame:
    """Return synthetic 14-column PLM data as a pandas DataFrame."""
    return pd.DataFrame(sample_14col_plm_records)


@pytest.fixture
def sample_plm_excel_file(tmp_path: Path, sample_14col_plm_records: List[Dict[str, Any]]) -> Path:
    """Create a physical OpenXML .xlsx file matching the 14-column PLM export format."""
    file_path = tmp_path / "PLM_TEST_BOM.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PLM"

    headers = [
        "Row Index", "Level", "Item Type", "Item Id", "Has Children", "Quantity",
        "1st Parts", "2nd BOM Flag", "Occurrence Effectivities",
        "Item Revision Projects List", "Item Name", "Notice No", "Revision", "Item Rev Status"
    ]
    ws.append(headers)

    for rec in sample_14col_plm_records:
        ws.append([
            rec["row_idx"],
            rec["level"],
            rec["item_type"],
            rec["item_id"],
            rec["has_children"],
            rec["quantity"],
            rec["first_parts"],
            rec["second_bom_flag"],
            rec["occurrence_effectivities"],
            rec["item_revision_projects_list"],
            rec["item_name"],
            rec["notice_no"],
            rec["revision"],
            rec["item_rev_status"],
        ])

    wb.save(file_path)
    wb.close()
    return file_path


# ============================================================================
# Synthetic SAP CS12 Data Fixtures
# ============================================================================

@pytest.fixture
def sample_cs12_html_content() -> str:
    """Generate realistic HTML-in-XLS string produced by SAPLSPO5:0150 export."""
    return """
    <html>
    <head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
    <body>
    <table border="1">
        <tr><td colspan="10">SAP AG - CS12 Multilevel BOM</td></tr>
        <tr><td>Material: 110C103NL0</td><td>Plant: 2200</td><td>Usage: pp01</td><td>Alt: 01</td></tr>
        <tr><td colspan="10">-------------------------------------------------------------</td></tr>
        <tr>
            <th>Level</th>
            <th>Item</th>
            <th>Component</th>
            <th>Object Description</th>
            <th>Quantity</th>
            <th>Un</th>
            <th>RevLev</th>
            <th>Valid From</th>
            <th>Valid To</th>
            <th>Change No</th>
        </tr>
        <tr>
            <td>.1</td>
            <td>0010</td>
            <td>302FP93010</td>
            <td>LSU UNIT</td>
            <td>1.0</td>
            <td>PC</td>
            <td>A</td>
            <td>2024/01/01</td>
            <td>9999/12/31</td>
            <td>ECN1001</td>
        </tr>
        <tr>
            <td>..2</td>
            <td>0020</td>
            <td>302FP94010</td>
            <td>POLYGON MOTOR ASSY</td>
            <td>1.0</td>
            <td>PC</td>
            <td>A</td>
            <td>2024/01/01</td>
            <td>9999/12/31</td>
            <td>ECN1001</td>
        </tr>
        <tr>
            <td>...3</td>
            <td>0030</td>
            <td>302FP02010</td>
            <td>MOTOR BRACKET</td>
            <td>1.0</td>
            <td>PC</td>
            <td>01</td>
            <td>2024/01/01</td>
            <td>9999/12/31</td>
            <td>ECN1001</td>
        </tr>
        <tr>
            <td>...3</td>
            <td>0040</td>
            <td>B1303060</td>
            <td>SCREW M3X6</td>
            <td>4.0</td>
            <td>PC</td>
            <td>-</td>
            <td>2024/01/01</td>
            <td>9999/12/31</td>
            <td>ECN1001</td>
        </tr>
        <tr>
            <td>.1</td>
            <td>0050</td>
            <td>302FP93020</td>
            <td>FUSER UNIT</td>
            <td>1.0</td>
            <td>PC</td>
            <td>02</td>
            <td>2024/01/01</td>
            <td>9999/12/31</td>
            <td>ECN1002</td>
        </tr>
        <tr>
            <td>..2</td>
            <td>0060</td>
            <td>302FP04010</td>
            <td>HEATER LAMP 220V</td>
            <td>1.0</td>
            <td>PC</td>
            <td>01</td>
            <td>2024/01/01</td>
            <td>9999/12/31</td>
            <td>ECN1002</td>
        </tr>
    </table>
    </body>
    </html>
    """


@pytest.fixture
def sample_cs12_file(tmp_path: Path, sample_cs12_html_content: str) -> Path:
    """Create a physical R3 export file containing HTML spreadsheet data."""
    dest = tmp_path / "R3_110C103NL0_17_09_2026.xls"
    dest.write_text(sample_cs12_html_content, encoding="utf-8")
    return dest


# ============================================================================
# Synthetic CTTT & Reconciliation Fixtures
# ============================================================================

@pytest.fixture
def sample_cttt_df() -> pd.DataFrame:
    """Generates member work instructions (CTTT) input table."""
    return pd.DataFrame([
        {
            "SUB": "LSU",
            "TRANG CTTT": "P.01",
            "MÃ LINH KIỆN": "302FP02010",
            "TÊN LINH KIỆN": "MOTOR BRACKET",
            "SỐ LƯỢNG": 1.0,
            "PHỤ TRÁCH": "Nguyen Van A",
            "GIAI_THICH": "",
        },
        {
            "SUB": "LSU",
            "TRANG CTTT": "P.02",
            "MÃ LINH KIỆN": "B1303060",
            "TÊN LINH KIỆN": "SCREW M3X6",
            "SỐ LƯỢNG": 4.0,
            "PHỤ TRÁCH": "Nguyen Van A",
            "GIAI_THICH": "",
        },
        {
            "SUB": "FUSER",
            "TRANG CTTT": "P.10",
            "MÃ LINH KIỆN": "302FP04010",
            "TÊN LINH KIỆN": "HEATER LAMP 220V",
            "SỐ LƯỢNG": 1.0,
            "PHỤ TRÁCH": "Tran Van B",
            "GIAI_THICH": "",
        },
    ])


@pytest.fixture
def sample_msi_df() -> pd.DataFrame:
    """Generates member MSI / barcode input records."""
    return pd.DataFrame([
        {
            "Mã LK Barcode": "302FP93010",
            "Mã UNIT/linh kiện bản mạch": "302FP93010-01",
            "Tên UNIT": "LSU UNIT",
            "3 ký tự MSI": "2NL",
            "SEVICE": "-",
            "ABS": "OK",
            "Phụ trách": "Nguyen Van A",
        },
        {
            "Mã LK Barcode": "302FP93020",
            "Mã UNIT/linh kiện bản mạch": "302FP93020-02",
            "Tên UNIT": "FUSER UNIT",
            "3 ký tự MSI": "1HN",
            "SEVICE": "SERVICE REQUIRED",
            "ABS": "OK",
            "Phụ trách": "Tran Van B",
        },
    ])


# ============================================================================
# Machine Model Pruning Rules Fixture (BolocBom)
# ============================================================================

@pytest.fixture
def sample_bolocbom_rules() -> Dict[str, List[Dict[str, Any]]]:
    """Authentic rules from Sheet BolocBom across the 6 machine models."""
    return {
        "Virgo": [
            {"item_name": "BOTTLE WASTE", "match_mode": "Full_name", "part_code": None},
            {"item_name": "FILL UP CONTAINER ASSY", "match_mode": "Part_name", "part_code": None},
            {"item_name": "PWB PANEL MAIN ASSY WITH SOFTWARE", "match_mode": "Full_name", "part_code": None},
        ],
        "Libra2": [
            {"item_name": "FRAME CONVEYING", "match_mode": "Full_name", "part_code": None},
            {"item_name": "HOLDER BELT TENSION R", "match_mode": "Full_name", "part_code": None},
        ],
        "Iris2024": [
            {"item_name": "PWB ENGINE ASSY WITH SOFTWARE", "match_mode": "Part_name", "part_code": None},
            {"item_name": "PWB FEED DRIVE ASSY", "match_mode": "Part_name", "part_code": None},
        ],
        "Sirius2": [
            {"item_name": "PWB SWITCH ASSY", "match_mode": "Full_name", "part_code": None},
            {"item_name": "P.W.BOARD ASSY THERMISTOR", "match_mode": "Full_name", "part_code": None},
        ],
        "Mebius": [
            {"item_name": "BOTTLE WASTE", "match_mode": "Full_name", "part_code": None},
            {"item_name": None, "match_mode": "Full_name", "part_code": "302FP20250"},
        ],
        "Polaris": [
            {"item_name": "BOTTLE WASTE", "match_mode": "Full_name", "part_code": None},
            {"item_name": "FILL UP CONT ASSY", "match_mode": "Part_name", "part_code": None},
        ],
    }


# ============================================================================
# External Automation Mocks (TC14 WebDriver & SAP COM Session)
# ============================================================================

@pytest.fixture
def mock_tc14_driver() -> MagicMock:
    """Create a simulated Selenium WebDriver pre-configured for TC14 Active Workspace."""
    driver = MagicMock()
    driver.current_url = "http://tcmp3gwb:3000/#/showHome"
    driver.page_source = "<html><body><div class='aw-layout-mainView'></div></body></html>"
    driver.get_cookies.return_value = [
        {"name": "JSESSIONID", "value": "TEST_SESSION_123", "domain": "tcmp3gwb"},
        {"name": "XSRF-TOKEN", "value": "TEST_XSRF_TOKEN", "domain": "tcmp3gwb"},
        {"name": "sanSID", "value": "AUTH_COOKIE_SID", "domain": "tcmp3gwb"},
    ]
    return driver


@pytest.fixture
def mock_sap_session() -> MagicMock:
    """Create a simulated SAP GUI Scripting COM session."""
    session = MagicMock()
    session.Busy = False

    def find_by_id_mock(elem_id: str):
        elem = MagicMock()
        elem.Text = ""
        elem.MessageType = "S"  # Success status
        if "okcd" in elem_id:
            elem.Text = "/nCS12"
        elif "sbar" in elem_id:
            elem.MessageType = "S"
            elem.Text = "BOM displayed successfully"
        return elem

    session.findById.side_effect = find_by_id_mock
    return session
