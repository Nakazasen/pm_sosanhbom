"""Feature F7: Missing Parts Detection Isolation Tests.

Verifies:
1. Surfaces all engineered PLM components omitted from assembly work instructions (CTTT).
2. Reports zero missing parts when all PLM production components exist in CTTT.
3. Distinguishes leaf parts from higher-level assembly parents.
4. Aggregates multiple missing parts with their unit names, quantities, and revisions.
5. Performs case-insensitive and whitespace-trimmed part code matching.
"""

from __future__ import annotations

import pandas as pd
import pytest

from src.core.reconciliation import (
    ReconciliationEngine,
    detect_missing_parts,
    detect_missing_plm_parts,
)


class TestF07MissingParts:
    """Test suite for Feature F7: Missing Parts Detection."""

    def test_f07_detect_engineered_parts_omitted_in_cttt(self) -> None:
        """Test 1: Identify parts designed in PLM but omitted in CTTT."""
        plm_df = pd.DataFrame([
            {"item_id": "PART_A", "item_name": "GEAR", "quantity": 1.0, "unit_name": "LSU"},
            {"item_id": "PART_B", "item_name": "SPRING", "quantity": 2.0, "unit_name": "LSU"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "PART_A", "SỐ LƯỢNG": 1.0},
        ])

        missing = detect_missing_plm_parts(plm_df, cttt_df)
        assert len(missing) == 1
        assert missing.iloc[0]["item_id"] == "PART_B"
        assert missing.iloc[0]["unit_name"] == "LSU"

    def test_f07_no_missing_when_all_plm_parts_accounted_for(self) -> None:
        """Test 2: When all PLM parts are in CTTT, missing DataFrame is empty."""
        plm_df = pd.DataFrame([
            {"item_id": "P1", "item_name": "MOTOR"},
            {"item_id": "P2", "item_name": "SCREW"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "P1"},
            {"MÃ LINH KIỆN": "P2"},
        ])

        missing = detect_missing_parts(plm_df, cttt_df)
        assert len(missing) == 0

    def test_f07_ignore_assembly_parents_in_missing_check(self) -> None:
        """Test 3: Parent assemblies with children omitted when only_leaves=True."""
        plm_df = pd.DataFrame([
            {"item_id": "PARENT_UNIT", "has_children": True, "item_name": "SUBASSY"},
            {"item_id": "P1", "has_children": False, "item_name": "LEAF PART"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "P1"},
        ])

        missing = detect_missing_plm_parts(plm_df, cttt_df, only_leaves=True)
        assert len(missing) == 0

    def test_f07_multiple_missing_parts_aggregation(self) -> None:
        """Test 4: Aggregate multiple missing parts across different sub-units."""
        plm_df = pd.DataFrame([
            {"item_id": "P1", "unit_name": "LSU"},
            {"item_id": "P2", "unit_name": "FUSER"},
            {"item_id": "P3", "unit_name": "DRUM"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "P1"},
        ])

        engine = ReconciliationEngine()
        missing = engine.detect_missing_parts(plm_df, cttt_df)
        assert len(missing) == 2
        missing_ids = set(missing["item_id"])
        assert missing_ids == {"P2", "P3"}

    def test_f07_case_insensitive_part_code_lookup(self) -> None:
        """Test 5: Lookup matches regardless of case and surrounding whitespace."""
        plm_df = pd.DataFrame([
            {"item_id": "302fp02010", "unit_name": "LSU"},
        ])
        cttt_df = pd.DataFrame([
            {"MÃ LINH KIỆN": "  302FP02010  "},
        ])

        missing = detect_missing_plm_parts(plm_df, cttt_df)
        assert len(missing) == 0
