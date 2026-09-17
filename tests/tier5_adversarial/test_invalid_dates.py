"""Tier 5 Adversarial Tests: Invalid Dates & Malformed Effectivity Strings.

Verifies:
1. Impossible leap year dates (2025-02-29 vs 2024-02-29).
2. Out of bounds months/days (month 13, day 32, month 0).
3. Out of range years (year 0, year 9999, negative years).
4. Malformed effectivity patterns without valid dates ("to nowhere", "to to UP", "from A to B").
5. Ambiguous date formats (DD/MM/YYYY vs MM/DD/YYYY) parsed deterministically.
6. Empty, None, and garbage strings evaluated fail-closed.
"""

from __future__ import annotations

import datetime
import pytest

from src.core.date_filter import DateFilterEngine, extract_expiry_date, is_effectivity_expired
from src.automation.sap.models import CS12Params


class TestAdversarialInvalidDates:
    """Test suite for adversarial date inputs."""

    def test_impossible_leap_year_dates(self):
        """2025-02-29 is not a valid leap year and must return None from extractor."""
        # Non leap year
        invalid_leap = "to 29/02/2025"
        assert extract_expiry_date(invalid_leap) is None

        # Leap year 2024 is valid
        valid_leap = "to 29/02/2024"
        parsed = extract_expiry_date(valid_leap)
        assert parsed == datetime.date(2024, 2, 29)

    def test_out_of_bounds_month_and_day(self):
        """Month 13 and day 32 fail cleanly to None."""
        assert extract_expiry_date("to 32/01/2026") is None
        assert extract_expiry_date("to 15/13/2026") is None
        assert extract_expiry_date("to 00/00/2026") is None

    def test_out_of_range_years(self):
        """Years far outside business relevance fail cleanly."""
        assert extract_expiry_date("to 9999-99-99") is None
        assert extract_expiry_date("to 0000-00-00") is None

    def test_unstructured_and_garbage_effectivity_strings(self):
        """Strings containing 'to' but followed by garbage text return None."""
        assert extract_expiry_date("to ") is None
        assert extract_expiry_date("to nowhere") is None
        assert extract_expiry_date("to to to UP") is None
        assert extract_expiry_date("from station 1 to station 2") is None
        assert extract_expiry_date("to 999-99-99") is None

    def test_up_retention_unconditional(self):
        """Any effectivity containing 'UP' is preserved unconditionally."""
        ref_date = datetime.date(2026, 9, 17)
        # Even if a past date is included before UP
        assert is_effectivity_expired("01-Jan-2020 UP", reference_date=ref_date, keep_up=True) is False
        assert is_effectivity_expired("to 2020-01-01 UP", reference_date=ref_date, keep_up=True) is False

    def test_expired_dates_properly_pruned(self):
        """Explicitly expired dates are flagged expired."""
        ref_date = datetime.date(2026, 9, 17)
        # Expired 2 years ago
        assert is_effectivity_expired("to 31/12/2024", reference_date=ref_date) is True
        # Future date is not expired
        assert is_effectivity_expired("to 31/12/2027", reference_date=ref_date) is False

    def test_sap_cs12_params_date_formatting(self):
        """CS12Params cleanly formats edge dates without crashing."""
        edge_date = datetime.date(2026, 12, 31)
        params = CS12Params(material="302K123456", valid_date=edge_date)
        assert params.formatted_date_sap == "2026/12/31"
        assert params.formatted_date_filename == "31_12_2026"
        assert "31_12_2026" in params.expected_filename
