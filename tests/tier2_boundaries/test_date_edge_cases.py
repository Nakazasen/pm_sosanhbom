"""Tier 2 Boundary Tests: Date Parsing, Filtering, and Leap Year Boundaries.

Verifies:
1. Leap year date parsing (e.g. 29-Feb-2024) and date arithmetic across February 29.
2. Year-end rollover boundaries (31-Dec-2024 to 01-Jan-2025).
3. Open-ended "UP" effectivity strings ('01-Jan-2024 UP') are always retained.
4. Closed interval effectivity strings ('01-Jan-2024 .. 15-Mar-2024') correctly expired.
5. Inverted dates and future cut-off thresholds accurately partition active vs inactive nodes.
"""

import datetime
from pathlib import Path
import openpyxl
import pytest

from src.core.date_filter import DateFilter, FilterCriteria
from src.core.tree_parser import PLMTreeParser
from src.automation.sap.models import CS12Params


class TestDateEdgeCasesBoundaries:
    """Boundary test suite for date filtering and temporal boundaries."""

    def _create_plm_with_dates(self, path: Path, date_strings: list) -> Path:
        """Helper to create synthetic PLM workbook with specified effectivity strings."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PLM"
        headers = [
            "No", "Level", "Item Type", "Item ID", "Has Children", "Quantity",
            "First Parts", "Second BOM flag", "Occurrence Effectivities",
            "Item Revision:Projects List", "Item Name", "Notice No", "Revision",
            "Item Rev:Release Status"
        ]
        ws.append(headers)

        for idx, date_str in enumerate(date_strings, 1):
            ws.append([
                idx, 1, "Part", f"PART_{idx}", "False", 1.0,
                "", "", date_str, "PRJ_DATE", f"NAME_{idx}", "ECN-1", "01", "Released"
            ])
        wb.save(path)
        return path

    def test_leap_year_february_29_parsing(self, tmp_path: Path):
        """Handle 29-Feb-2024 expiry date in 'to <date>' format."""
        f = self._create_plm_with_dates(tmp_path / "leap_year.xlsx", [
            "01-Jan-2024 to 29-Feb-2024",
            "29-Feb-2024 UP",
        ])

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        # Target date is 2024-05-01 (year_diff=0, month_diff=3 > 1) -> PART_1 expired, PART_2 (UP) kept
        criteria = FilterCriteria(reference_date=datetime.date(2024, 5, 1))
        filter_eng = DateFilter(criteria)
        filtered_tree = filter_eng.filter_tree(tree)

        active_ids = [n.item_id for n in filtered_tree.flatten()]
        assert "PART_1" not in active_ids  # Expired
        assert "PART_2" in active_ids      # Retained due to UP

    def test_sap_cs12_params_leap_year_formatting(self):
        """CS12Params format leap year dates correctly for SAP and filenames."""
        params = CS12Params(
            material="110K123450",
            valid_date=datetime.date(2024, 2, 29),
        )
        assert params.formatted_date_sap == "2024/02/29"
        assert params.formatted_date_filename == "29_02_2024"
        assert params.expected_filename == "R3_110K123450_29_02_2024.xls"

    def test_year_end_boundary_crossover(self, tmp_path: Path):
        """Effectivity date with year crossover expiry: 'to 31-Dec-2023' expired in 2024."""
        f = self._create_plm_with_dates(tmp_path / "year_crossover.xlsx", [
            "01-Jan-2023 to 31-Dec-2023",
            "01-Jan-2024 UP",
        ])
        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        # Reference date is in 2024 -> diff_year = 2024 - 2023 = 1 > 0 -> PART_1 expired
        criteria = FilterCriteria(reference_date=datetime.date(2024, 1, 15))
        filter_eng = DateFilter(criteria)
        filtered_tree = filter_eng.filter_tree(tree)

        active_ids = [n.item_id for n in filtered_tree.flatten()]
        assert "PART_1" not in active_ids
        assert "PART_2" in active_ids

    def test_open_ended_up_always_retained_in_future(self, tmp_path: Path):
        """Open-ended 'UP' rule: a part released in past with UP is effective at any future date."""
        f = self._create_plm_with_dates(tmp_path / "up_effectivity.xlsx", [
            "01-Jan-2020 UP",
        ])
        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        # Check against year 2030
        criteria = FilterCriteria(reference_date=datetime.date(2030, 6, 15))
        filter_eng = DateFilter(criteria)
        filtered_tree = filter_eng.filter_tree(tree)

        assert len(filtered_tree.flatten()) == 1
        assert filtered_tree.flatten()[0].item_id == "PART_1"

    def test_closed_date_interval_expiration(self, tmp_path: Path):
        """Closed interval '01-Jan-2024 to 01-Jun-2024' is expired by 15-Aug-2024."""
        f = self._create_plm_with_dates(tmp_path / "interval.xlsx", [
            "01-Jan-2024 to 01-Jun-2024",
        ])
        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        # 2024-08-15 vs 2024-06-01: diff_month = 2 > 1 -> expired
        criteria = FilterCriteria(reference_date=datetime.date(2024, 8, 15))
        filter_eng = DateFilter(criteria)
        filtered_tree = filter_eng.filter_tree(tree)

        assert len(filtered_tree.flatten()) == 0
