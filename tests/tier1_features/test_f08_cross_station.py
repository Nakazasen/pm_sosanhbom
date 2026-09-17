"""Feature F8: Cross-Station Total Aggregation Isolation Tests.

Verifies:
1. Components used across multiple stations/sub-units (e.g. standard screws) are summed correctly.
2. Cross-station sum matching R3 total yields 'OK'.
3. Cross-station sum differing from R3 total yields 'NG'.
4. Single-station parts pass through correctly.
5. Floating-point quantities (e.g. grease, wire lengths) aggregate with decimal precision.
"""

from __future__ import annotations

import pandas as pd
import pytest

from src.core.reconciliation import (
    ReconciliationEngine,
    aggregate_cross_station,
    aggregate_cross_station_totals,
)


class TestF08CrossStation:
    """Test suite for Feature F8: Cross-Station Total Aggregation."""

    def test_f08_sum_shared_screws_across_units(self) -> None:
        """Test 1: Sum shared screw quantities used across multiple stations."""
        cttt = pd.DataFrame([
            {"MÃ LINH KIỆN": "B1303060", "SUB": "LSU", "SỐ LƯỢNG": 4.0},
            {"MÃ LINH KIỆN": "B1303060", "SUB": "FUSER", "SỐ LƯỢNG": 6.0},
            {"MÃ LINH KIỆN": "B1303060", "SUB": "FRAME", "SỐ LƯỢNG": 10.0},
        ])
        r3 = pd.DataFrame([
            {"part_code": "B1303060", "r3_total_qty": 20.0},
        ])

        totals = aggregate_cross_station_totals(cttt, r3)
        assert len(totals) == 1
        assert totals.iloc[0]["CTTT_SUM_QTY"] == 20.0
        assert totals.iloc[0]["STATUS"] == "OK"

    def test_f08_cross_station_matching_yields_ok(self) -> None:
        """Test 2: When CTTT aggregated total equals R3 grand total, status is 'OK'."""
        cttt = pd.DataFrame([
            {"MÃ LINH KIỆN": "PART1", "SỐ LƯỢNG": 2.0},
            {"MÃ LINH KIỆN": "PART1", "SỐ LƯỢNG": 3.0},
        ])
        r3 = pd.DataFrame([
            {"part_code": "PART1", "r3_total_qty": 5.0},
        ])

        engine = ReconciliationEngine()
        totals = engine.aggregate_cross_station(cttt, r3)
        assert totals.iloc[0]["STATUS"] == "OK"

    def test_f08_cross_station_mismatch_yields_ng(self) -> None:
        """Test 3: When CTTT sum does not equal R3 grand total, status is 'NG'."""
        cttt = pd.DataFrame([
            {"MÃ LINH KIỆN": "PART1", "SỐ LƯỢNG": 2.0},
            {"MÃ LINH KIỆN": "PART1", "SỐ LƯỢNG": 2.0},
        ])
        r3 = pd.DataFrame([
            {"part_code": "PART1", "r3_total_qty": 5.0},
        ])

        totals = aggregate_cross_station(cttt, r3)
        assert totals.iloc[0]["CTTT_SUM_QTY"] == 4.0
        assert totals.iloc[0]["r3_total_qty"] == 5.0
        assert totals.iloc[0]["STATUS"] == "NG"

    def test_f08_single_station_components_preserved(self) -> None:
        """Test 4: Components used in only a single sub-unit aggregate without distortion."""
        cttt = pd.DataFrame([
            {"MÃ LINH KIỆN": "UNIQUE_LENS", "SỐ LƯỢNG": 1.0},
        ])
        r3 = pd.DataFrame([
            {"part_code": "UNIQUE_LENS", "r3_total_qty": 1.0},
        ])

        totals = aggregate_cross_station_totals(cttt, r3)
        assert totals.iloc[0]["CTTT_SUM_QTY"] == 1.0
        assert totals.iloc[0]["STATUS"] == "OK"

    def test_f08_roundtrip_precision_floating_quantities(self) -> None:
        """Test 5: Accurate summation of fractional quantities."""
        cttt = pd.DataFrame([
            {"MÃ LINH KIỆN": "GREASE", "SỐ LƯỢNG": 0.25},
            {"MÃ LINH KIỆN": "GREASE", "SỐ LƯỢNG": 0.75},
        ])
        r3 = pd.DataFrame([
            {"part_code": "GREASE", "r3_total_qty": 1.0},
        ])

        totals = aggregate_cross_station(cttt, r3)
        assert totals.iloc[0]["CTTT_SUM_QTY"] == 1.0
        assert totals.iloc[0]["STATUS"] == "OK"
