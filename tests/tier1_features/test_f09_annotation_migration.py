"""Feature F9: Annotation Migration Engine Isolation Tests.

Verifies:
1. Member explanations ('Giải thích') migrate seamlessly from previous PLM sheets to new exports.
2. Person in charge ('Phụ trách') annotations are preserved.
3. Manager verification ('Quản lý check') marks are preserved.
4. New engineering additions in new revisions receive clean empty strings (not 'nan').
5. Removed / deprecated components remain intact in the backup history sheet.
"""

from __future__ import annotations

import pandas as pd
import pytest

from src.core.reconciliation import ReconciliationEngine, migrate_annotations


class TestF09AnnotationMigration:
    """Test suite for Feature F9: Annotation Migration Engine."""

    def test_f09_preserve_member_explanations_across_revisions(self) -> None:
        """Test 1: Member explanations in old PLM sheet migrate to new PLM sheet."""
        old_df = pd.DataFrame([
            {"part_code": "302FP02010", "giai_thich": "Dung cho model moi", "phu_trach": "An", "quan_ly_check": "OK"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "302FP02010", "item_name": "MOTOR BRACKET", "quantity": 1.0},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["giai_thich"] == "Dung cho model moi"

    def test_f09_preserve_person_in_charge(self) -> None:
        """Test 2: Preserve person in charge ('Phụ trách') field."""
        old_df = pd.DataFrame([
            {"part_code": "P100", "giai_thich": "", "phu_trach": "Nguyen Van A", "quan_ly_check": ""},
        ])
        new_df = pd.DataFrame([
            {"part_code": "P100", "item_name": "LENS"},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["phu_trach"] == "Nguyen Van A"

    def test_f09_preserve_manager_check_status(self) -> None:
        """Test 3: Preserve manager check ('Quản lý check') verification."""
        old_df = pd.DataFrame([
            {"part_code": "P100", "giai_thich": "", "phu_trach": "", "quan_ly_check": "APPROVED_BY_LEADER"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "P100", "item_name": "LENS"},
        ])

        engine = ReconciliationEngine()
        migrated = engine.migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["quan_ly_check"] == "APPROVED_BY_LEADER"

    def test_f09_handle_newly_introduced_parts(self) -> None:
        """Test 4: Brand new parts introduced in new revision receive empty string annotations."""
        old_df = pd.DataFrame([
            {"part_code": "OLD_PART", "giai_thich": "Old note", "phu_trach": "A", "quan_ly_check": "OK"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "NEW_PART_XYZ", "item_name": "NEW SENSOR"},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert migrated.iloc[0]["giai_thich"] == ""
        assert migrated.iloc[0]["phu_trach"] == ""
        assert migrated.iloc[0]["quan_ly_check"] == ""

    def test_f09_removed_parts_archived_in_old_sheet(self) -> None:
        """Test 5: Parts existing only in old sheet remain available in old backup DataFrame."""
        old_df = pd.DataFrame([
            {"part_code": "REMOVED_PART", "giai_thich": "Discontinued", "phu_trach": "B", "quan_ly_check": "OK"},
        ])
        new_df = pd.DataFrame([
            {"part_code": "ACTIVE_PART", "item_name": "ACTIVE"},
        ])

        migrated = migrate_annotations(new_df, old_df)
        assert len(migrated) == 1
        assert "REMOVED_PART" not in set(migrated["part_code"])
        assert "REMOVED_PART" in set(old_df["part_code"])
