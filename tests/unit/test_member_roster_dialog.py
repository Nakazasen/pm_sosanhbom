"""Unit tests for MemberRosterDialog GUI component."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch
import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.member_roster_dialog import MemberRosterDialog
from src.core.member_database import MemberRecord


class TestMemberRosterDialog:
    """Test suite for MemberRosterDialog."""

    def test_dialog_loads_roster_and_filters(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify dialog populates table and filters correctly."""
        local_dir = tmp_path / "app"
        remote_file = tmp_path / "remote" / "ssbom_master.db"

        dlg = MemberRosterDialog(base_dir=local_dir, remote_db_path=remote_file)

        # Initially loaded 38 canonical engineers
        assert dlg.member_table.rowCount() == 38

        # Filter by department "Cơ 1"
        dlg.combo_dept_filter.setCurrentText("Cơ 1")
        assert dlg.member_table.rowCount() == 12

        # Search filter
        dlg.search_edit.setText("Son_mecha1")
        assert dlg.member_table.rowCount() == 1

        # Clear search
        dlg.search_edit.clear()
        dlg.combo_dept_filter.setCurrentText("Tất cả")
        assert dlg.member_table.rowCount() == 38

    def test_dialog_add_and_delete_member_flow(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify adding and deleting a member via dialog methods."""
        local_dir = tmp_path / "app"
        remote_file = tmp_path / "remote" / "ssbom_master.db"

        dlg = MemberRosterDialog(base_dir=local_dir, remote_db_path=remote_file)

        # Set form data
        dlg.txt_account_id.setText("Test_NewMember")
        dlg.txt_full_name.setText("Kỹ sư Kiểm thử")
        dlg.combo_dept.setCurrentText("Cơ 1")
        dlg.combo_sub_unit.setCurrentText("DRUM")

        # Mock QMessageBox to prevent popups during test
        with patch("src.gui.member_roster_dialog.QMessageBox.information"):
            dlg._on_add_member()

        assert dlg.member_table.rowCount() == 39

        # Find and select the row containing Test_NewMember
        target_row = -1
        for r in range(dlg.member_table.rowCount()):
            if dlg.member_table.item(r, 1).text() == "Test_NewMember":
                target_row = r
                break
        assert target_row >= 0
        dlg.member_table.selectRow(target_row)
        assert dlg.txt_account_id.text() == "Test_NewMember"
        assert dlg.txt_full_name.text() == "Kỹ sư Kiểm thử"

        # Delete the member
        with patch("src.gui.member_roster_dialog.QMessageBox.question", return_value=16384):  # StandardButton.Yes
            with patch("src.gui.member_roster_dialog.QMessageBox.information"):
                dlg._on_delete_member()

        assert dlg.member_table.rowCount() == 38

    def test_dialog_edit_account_id_and_multi_subunits(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify that account_id can be edited (unlocked) and multi-subunits can be saved."""
        local_dir = tmp_path / "app"
        remote_file = tmp_path / "remote" / "ssbom_master.db"

        dlg = MemberRosterDialog(base_dir=local_dir, remote_db_path=remote_file)

        # Select row for Son_mecha1
        target_row = -1
        for r in range(dlg.member_table.rowCount()):
            if dlg.member_table.item(r, 1).text() == "Son_mecha1":
                target_row = r
                break
        assert target_row >= 0
        dlg.member_table.selectRow(target_row)
        assert dlg.txt_account_id.isReadOnly() is False
        assert dlg.txt_account_id.text() == "Son_mecha1"

        # Edit account_id and multi-subunits
        dlg.txt_account_id.setText("Son_mecha1_edited")
        dlg.txt_full_name.setText("Nguyễn Văn Sơn Cải Tiến")
        dlg.combo_sub_unit.setEditText("LSU, DRUM")

        with patch("src.gui.member_roster_dialog.QMessageBox.information"):
            dlg._on_update_member()

        # Check that table updated with new account_id and multi-subunits
        found = False
        for r in range(dlg.member_table.rowCount()):
            if dlg.member_table.item(r, 1).text() == "Son_mecha1_edited":
                assert dlg.member_table.item(r, 2).text() == "Nguyễn Văn Sơn Cải Tiến"
                assert dlg.member_table.item(r, 4).text() == "LSU, DRUM"
                found = True
                break
        assert found is True

    def test_dialog_machine_names_and_picker(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify machine_names column, multi-model assignment, and search filtering."""
        from src.gui.member_roster_dialog import SelectMachinesDialog

        local_dir = tmp_path / "app"
        remote_file = tmp_path / "remote" / "ssbom_master.db"

        dlg = MemberRosterDialog(base_dir=local_dir, remote_db_path=remote_file)

        # Verify 8 columns with "Dòng máy" at column 5
        assert dlg.member_table.columnCount() == 8
        assert dlg.member_table.horizontalHeaderItem(5).text() == "Dòng máy"

        # Select first member
        dlg.member_table.selectRow(0)
        acc_id = dlg.txt_account_id.text()
        assert bool(acc_id) is True

        # Assign multiple machine models
        dlg.txt_machine_names.setText("Virgo, 6th Next")
        with patch("src.gui.member_roster_dialog.QMessageBox.information"):
            dlg._on_update_member()

        # Check that table updated column 5 with normalized name
        assert dlg.member_table.item(0, 5).text() == "Virgo, 6thNext"

        # Test searching by machine model name
        dlg.search_edit.setText("6thNext")
        assert dlg.member_table.rowCount() == 1
        assert dlg.member_table.item(0, 1).text() == acc_id

        # Test SelectMachinesDialog
        selector = SelectMachinesDialog(["Virgo", "Iris 2024", "6th Next"], ["virgo"])
        assert "Virgo" in selector.get_selected()
        assert len(selector.get_selected()) == 1

        selector._select_all()
        assert len(selector.get_selected()) == 3
        assert "Iris2024" in selector.get_selected()
        assert "6thNext" in selector.get_selected()

        selector._deselect_all()
        assert len(selector.get_selected()) == 0

