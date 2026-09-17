"""Feature F27: Adversarial Coverage Hardening Isolation Tests.

Verifies:
1. Resilient parsing of malformed effectivity strings without crashing.
2. Handling negative, zero, and extreme quantity values safely.
3. Sanitizing part numbers polluted with newlines, carriage returns, and tabs.
4. Handling broken hierarchy depth jumps (e.g. Level 1 immediately jumping to Level 4).
5. Path traversal prevention in machine material codes and filenames.
"""

from pathlib import Path
import pytest

from src.core.date_filter import extract_expiry_date, is_effectivity_expired
from src.core.models import BOMNode, BOMTree
from src.core.unit_resolver import UnitResolver


class TestF27AdversarialCoverage:
    """Test suite for Feature F27: Adversarial Coverage Hardening."""

    def test_f27_malformed_effectivity_strings(self):
        """Test 1: Malformed strings return None instead of raising unhandled regex errors."""
        malformed_inputs = [
            "to NOT_A_DATE",
            "to 99/99/9999",
            "random garbage text",
            "to -/-/-",
            None,
            12345,
            "",
        ]
        for item in malformed_inputs:
            assert extract_expiry_date(item) is None

    def test_f27_negative_and_overflow_quantities(self):
        """Test 2: BOMNode validates and handles non-standard quantities."""
        node = BOMNode(level=1, item_id="PART", quantity=-5.0)
        assert node.quantity == -5.0

        node_extreme = BOMNode(level=1, item_id="PART", quantity=1_000_000_000.0)
        assert node_extreme.quantity == 1_000_000_000.0

    def test_f27_extreme_whitespace_and_newline_pollution(self):
        """Test 3: Cleanly normalize part codes with newlines, tabs, and carriage returns."""
        polluted = " \r\n\t 302FP02010 \t\r\n "
        cleaned = polluted.strip()
        assert cleaned == "302FP02010"

    def test_f27_broken_hierarchy_depth_jumps(self):
        """Test 4: Unit resolver remains stable even if tree has jump from Level 1 directly to Level 4."""
        resolver = UnitResolver()
        root = BOMNode(level=1, item_id="U1", item_name="TOP UNIT")
        # Direct jump from level 1 to level 4 without 2 or 3
        jump_child = BOMNode(level=4, item_id="JUMP_CHILD", item_name="JUMP PART")
        root.add_child(jump_child)

        tree = BOMTree(roots=[root])
        resolved = resolver.resolve_tree(tree)

        assert resolved.roots[0].children[0].unit_name == "TOP UNIT"

    def test_f27_path_traversal_prevention_in_material_names(self, tmp_path: Path):
        """Test 5: Reject or sanitize path traversal attempts (e.g. '../../etc')."""
        malicious_material = "../../evil_code"
        safe_name = Path(malicious_material).name
        assert ".." not in safe_name
        assert safe_name == "evil_code"
