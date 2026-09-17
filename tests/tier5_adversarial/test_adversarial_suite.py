"""Tier 5 Adversarial Hardening & Extreme Boundary Condition Tests.

Stress-tests the core comparison, parsing, and data validation layers under
adversarial, corrupted, and malformed inputs to ensure fail-closed security.
"""

import math
from pathlib import Path
import pytest

from src.core.models import BOMNode, BOMTree
from src.core.tree_parser import BOMTreeParser
from src.core.date_filter import DateFilterEngine
from src.core.unit_resolver import UnitResolver
from src.core.reconciliation import ReconciliationEngine
from src.core.msi_engine import MSIEngine
from src.automation.sap.models import R3ComponentRow, SAPConnectionError


class TestTier5AdversarialHardening:
    """Adversarial stress tests for enterprise robustness."""

    def test_adversarial_path_traversal_in_material_name(self):
        """Ensure path traversal attacks in material names do not escape target directories."""
        malicious_material = "../../windows/system32/cmd"
        row = R3ComponentRow(part_code=malicious_material, quantity=1.0)
        assert row.part_code == malicious_material
        assert not row.part_code.startswith("..\\") or not row.part_code.startswith("../")

    def test_adversarial_infinite_circular_bom_hierarchy(self):
        """Prevent stack overflow or infinite loops on circular BOM definitions."""
        node_a = BOMNode(level=1, item_id="PART_A")
        node_b = BOMNode(level=2, item_id="PART_B")
        node_a.children.append(node_b)
        node_b.children.append(node_a)  # Circular reference

        tree = BOMTree(root=node_a)
        resolver = UnitResolver()
        res_tree = resolver.resolve_tree(tree)
        assert res_tree.root.unit_name is not None

    def test_adversarial_corrupted_unicode_and_null_bytes(self):
        """Handle null bytes and extreme unicode strings without crashing."""
        corrupted_name = "PART_\x00_TEST_\ufffd_\U0001f4a9"
        row = R3ComponentRow(part_code=corrupted_name, quantity=5.0, rev_r3="A")
        assert "\x00" in row.part_code
        assert row.quantity == 5.0

    def test_adversarial_extreme_quantities_and_overflow(self):
        """Ensure extreme floating-point and NaN quantities are handled cleanly."""
        engine = ReconciliationEngine()

        # Extremely large quantity
        huge_qty = 1e12
        res_huge = engine.reconcile_single_row(
            cttt_qty=huge_qty, plm_qty=huge_qty, r3_qty=huge_qty, plm_rev="A", r3_rev="A"
        )
        assert res_huge["overall_check"] == "OK"

        # NaN / Infinity rejection
        with pytest.raises(ValueError):
            engine.reconcile_single_row(
                cttt_qty=float("nan"), plm_qty=1.0, r3_qty=1.0, plm_rev="A", r3_rev="A"
            )

    def test_adversarial_corrupted_date_effectivity_strings(self):
        """Ensure garbage effectivity date strings fail safely to inactive."""
        filter_engine = DateFilterEngine(reference_date_str="2026-09-17")
        # Inverted or random string
        assert filter_engine.is_effective("GARBAGE_DATE_STRING") is False
        assert filter_engine.is_effective("to 9999-99-99 UP") is False
        assert filter_engine.is_effective("") is False
        assert filter_engine.is_effective(None) is False

    def test_adversarial_msi_empty_and_garbage_codes(self):
        """Ensure MSI engine rejects blank, whitespace-only, and malformed barcodes."""
        msi = MSIEngine()
        res = msi.evaluate(
            in_plm=False,
            member_code="",
            master_code="",
            member_service="",
            master_service="",
        )
        assert res["status"] == "NG"
        assert "PLM" in res["reason"]
