"""Integration tests for dynamic Member Roster, Settings, and Autocomplete."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import patch
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.core.member_database import MemberDatabaseManager, MemberRecord
from src.gui.leader_view import LeaderWorkspaceView
from src.gui.member_view import MemberWorkspaceView
from src.gui.settings_dialog import SettingsDialog, DEFAULT_SETTINGS


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestMemberRosterIntegration:
    """Test suite for roster integration across Leader View, Member View, and Settings."""

    def test_leader_view_has_manage_roster_button_and_reloads(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify Leader View Step 1 has btn_manage_roster and responds to DB changes."""
        db_path = tmp_path / "ssbom_master.db"
        db = MemberDatabaseManager(base_dir=tmp_path, remote_db_path=db_path)
        # Add a custom member
        db.add_member(
            MemberRecord(
                account_id="ZCustom_Eng",
                full_name="Custom Engineer",
                department="Cơ 1",
                default_sub_unit="LASER",
            )
        )

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # Verify button exists
        assert hasattr(step1, "btn_manage_roster")
        assert step1.btn_manage_roster.text() == "Quản lý nhân sự..."

        # Trigger reload from DB
        step1._reload_roster_from_db()

        # Check that table contains 38 default + 1 new = 39 members
        account_names = [step1.staff_table.item(r, 1).text() for r in range(step1.staff_table.rowCount())]
        assert "ZCustom_Eng" in account_names
        assert len(account_names) == 39

    def test_member_view_completer_loads_from_database(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify MemberWorkspaceView autocompletes from MemberDatabaseManager."""
        db_path = tmp_path / "ssbom_master.db"
        db = MemberDatabaseManager(base_dir=tmp_path, remote_db_path=db_path)
        db.add_member(
            MemberRecord(
                account_id="AutoCompleter_Test",
                full_name="Tester Autocomplete",
                department="Cơ 2",
            )
        )

        member_view = MemberWorkspaceView(base_dir=tmp_path)
        completer = member_view.author_edit.completer()
        assert completer is not None
        model = completer.model()
        items = [model.index(i, 0).data() for i in range(model.rowCount())]
        assert "Son_mecha1" in items
        assert "AutoCompleter_Test" in items

    def test_settings_dialog_loads_and_saves_shared_db_and_update_dir(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Verify SettingsDialog configures shared_db_path and update_dir."""
        config_file = tmp_path / "config.json"
        custom_db = r"\\fstvn01\Data\Test\ssbom_master.db"
        custom_update = r"\\fstvn01\Data\Test\release_update"

        init_cfg = {
            "paths": {
                "shared_db_path": custom_db,
                "update_dir": custom_update,
            }
        }
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(init_cfg, f)

        dlg = SettingsDialog(config_path=config_file)
        assert dlg.shared_db_edit.text() == custom_db
        assert dlg.update_dir_edit.text() == custom_update

        # Change paths and save
        new_db = r"\\fstvn01\Data\Production\ssbom_master.db"
        new_update = r"\\fstvn01\Data\Production\release_update"
        dlg.shared_db_edit.setText(new_db)
        dlg.update_dir_edit.setText(new_update)

        with patch.object(QMessageBox, "information"):
            dlg.save_and_close()

        with open(config_file, "r", encoding="utf-8") as f:
            saved = json.load(f)

        assert saved["paths"]["shared_db_path"] == new_db
        assert saved["paths"]["update_dir"] == new_update
