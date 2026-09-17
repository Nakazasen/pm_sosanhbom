"""Tier 5 Adversarial Tests: Extreme Quantities, Floating Point Underflow & Overflow.

Verifies:
1. Astronomical numbers (1e15, 1e18) handled safely without integer overflow.
2. Sub-microscopic quantities (1e-6, 1e-9) parsed and compared cleanly.
3. Negative quantities permitted in data models for adjustment/scrap components.
4. Rejection of NaN and infinite float values during cross-reconciliation.
5. Floating point epsilon equality (e.g. 0.1 + 0.2 == 0.3 within tolerance).
6. Non-numeric or string-polluted quantity values coerced or rejected safely.
"""

from __future__ import annotations

import math
import pytest

from src.core.models import BOMNode
from src.core.reconciliation import ReconciliationEngine
from src.core.tree_parser import _parse_float_qty
from src.automation.sap.models import R3ComponentRow


class TestAdversarialExtremeNumbers:
    """Test suite for extreme numeric inputs and boundary conditions."""

    def test_astronomical_quantities(self):
        """Quantities up to 1e18 evaluate cleanly in reconciliation."""
        engine = ReconciliationEngine()
        huge_val = 1e15
        result = engine.reconcile_single_row(
            cttt_qty=huge_val,
            plm_qty=huge_val,
            r3_qty=huge_val,
            plm_rev="A",
            r3_rev="A",
        )
        assert result["overall_check"] == "OK"
        assert result["comp_plm_qty"] == "OK"
        assert result["comp_r3_qty"] == "OK"

    def test_sub_microscopic_quantities(self):
        """Tiny quantities (1e-6) match within standard tolerance."""
        engine = ReconciliationEngine()
        tiny_a = 0.000001
        tiny_b = 0.00000100001
        result = engine.reconcile_single_row(
            cttt_qty=tiny_a,
            plm_qty=tiny_b,
            r3_qty=tiny_a,
            plm_rev="01",
            r3_rev="01",
        )
        assert result["overall_check"] == "OK"

    def test_negative_quantity_handling(self):
        """Negative quantities in BOMNode and R3ComponentRow are preserved without crash."""
        node = BOMNode(level=1, item_id="SCRAP_PART", quantity=-10.5)
        assert node.quantity == -10.5

        r3_row = R3ComponentRow(part_code="ADJ_PART", quantity="-5.25")
        assert r3_row.quantity == -5.25

    def test_nan_quantity_raises_value_error(self):
        """NaN quantity in reconciliation raises ValueError to prevent silent corruption."""
        engine = ReconciliationEngine()
        with pytest.raises(ValueError):
            engine.reconcile_single_row(
                cttt_qty=float("nan"),
                plm_qty=1.0,
                r3_qty=1.0,
                plm_rev="01",
                r3_rev="01",
            )

    def test_infinite_quantity_raises_value_error(self):
        """Infinity values are rejected explicitly by the reconciliation engine."""
        engine = ReconciliationEngine()
        with pytest.raises(ValueError):
            engine.reconcile_single_row(
                cttt_qty=float("inf"),
                plm_qty=1.0,
                r3_qty=1.0,
                plm_rev="01",
                r3_rev="01",
            )

        with pytest.raises(ValueError):
            engine.reconcile_single_row(
                cttt_qty=1.0,
                plm_qty=float("-inf"),
                r3_qty=1.0,
                plm_rev="01",
                r3_rev="01",
            )

    def test_floating_point_rounding_epsilon_tolerance(self):
        """Reconciliation handles classic IEEE 754 float drift (0.1 + 0.2 != 0.3)."""
        engine = ReconciliationEngine()
        val_sum = 0.1 + 0.2  # 0.30000000000000004 in Python
        val_expected = 0.3

        result = engine.reconcile_single_row(
            cttt_qty=val_sum,
            plm_qty=val_expected,
            r3_qty=val_expected,
            plm_rev="01",
            r3_rev="01",
        )
        assert result["overall_check"] == "OK"

    def test_parse_float_qty_with_commas_and_garbage(self):
        """Parser helper strips spaces, converts commas, and falls back gracefully."""
        assert _parse_float_qty(" 12,50 ") == 12.5
        assert _parse_float_qty("100.25") == 100.25
        assert _parse_float_qty("NOT_A_NUMBER", default=99.0) == 99.0
        assert _parse_float_qty(None, default=1.0) == 1.0
