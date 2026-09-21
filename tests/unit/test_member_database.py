"""Unit tests for MemberDatabaseManager (SQLite persistence & offline caching)."""

from __future__ import annotations

from pathlib import Path
import pytest

from src.core.member_database import MemberDatabaseManager, MemberRecord


class TestMemberDatabaseManager:
    """Test suite for MemberDatabaseManager CRUD and seeding."""

    def test_initialize_and_seed_defaults(self, tmp_path: Path) -> None:
        """Verify that a newly created DB is seeded with 38 canonical engineers."""
        local_dir = tmp_path / "app"
        remote_file = tmp_path / "remote" / "ssbom_master.db"

        mgr = MemberDatabaseManager(base_dir=local_dir, remote_db_path=remote_file)

        members = mgr.get_members()
        assert len(members) == 38

        # Check departments
        depts = mgr.get_departments()
        assert "Cơ 1" in depts
        assert "Cơ 2" in depts
        assert "Cơ 3" in depts

        # Check Son_mecha1
        son = next((m for m in members if m.account_id == "Son_mecha1"), None)
        assert son is not None
        assert son.department == "Cơ 1"
        assert son.is_active is True

    def test_add_update_delete_member(self, tmp_path: Path) -> None:
        """Test full CRUD operations on member records."""
        mgr = MemberDatabaseManager(base_dir=tmp_path / "app", remote_db_path=tmp_path / "remote.db")

        # 1. Add new member
        new_eng = MemberRecord(
            account_id="Vinh_mecha_lead",
            full_name="Bùi Đức Vinh",
            department="Cơ 1",
            default_sub_unit="DRUM",
            is_active=True,
            notes="Leader dự án",
        )
        ok, msg = mgr.add_member(new_eng)
        assert ok is True, msg

        # Duplicate account_id must fail
        ok_dup, _ = mgr.add_member(new_eng)
        assert ok_dup is False

        # Verify added
        members = mgr.get_members(department="Cơ 1")
        vinh = next((m for m in members if m.account_id == "Vinh_mecha_lead"), None)
        assert vinh is not None
        assert vinh.full_name == "Bùi Đức Vinh"
        assert vinh.default_sub_unit == "DRUM"

        # 2. Update member (test multi-subunits and account_id modification)
        vinh.default_sub_unit = "LSU, DRUM"
        vinh.notes = "Chuyển sang LSU, DRUM"
        # Test modifying account_id
        vinh.account_id = "Vinh_mecha_lead_v2"
        ok_up, msg_up = mgr.update_member(vinh)
        assert ok_up is True, msg_up

        members_after = mgr.get_members(department="Cơ 1")
        vinh_after = next((m for m in members_after if m.account_id == "Vinh_mecha_lead_v2"), None)
        assert vinh_after is not None
        assert vinh_after.default_sub_unit == "LSU, DRUM"
        assert vinh_after.notes == "Chuyển sang LSU, DRUM"

        # Check collision prevention: attempting to rename to an existing member should fail
        vinh_after.account_id = "Son_mecha1"
        ok_coll, msg_coll = mgr.update_member(vinh_after)
        assert ok_coll is False
        assert "trùng" in msg_coll.lower()

        # 3. Delete member
        ok_del, msg_del = mgr.delete_member("Vinh_mecha_lead_v2", member_id=vinh_after.id)
        assert ok_del is True, msg_del

        members_final = mgr.get_members(department="Cơ 1")
        assert not any(m.account_id == "Vinh_mecha_lead_v2" for m in members_final)

    def test_offline_fallback_when_remote_unavailable(self, tmp_path: Path) -> None:
        """Verify fallback to local cache when remote UNC is not accessible."""
        non_existent_remote = Path(r"\\non_existent_server_999\data\ssbom_master.db")
        local_base = tmp_path / "client_app"

        mgr = MemberDatabaseManager(base_dir=local_base, remote_db_path=non_existent_remote)

        assert mgr.is_remote_available() is False
        active_path, is_remote = mgr.get_active_db_path()
        assert is_remote is False
        assert active_path == mgr.local_cache_path
        assert active_path.exists()

        members = mgr.get_members()
        assert len(members) == 38
