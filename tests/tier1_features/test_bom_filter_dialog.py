"""GUI Tests for BOMFilterConfigDialog."""

import os
import tempfile
from pathlib import Path
import pytest
from PyQt6.QtWidgets import QApplication

from src.core.bom_filter_manager import BOMFilterManager
from src.gui.bom_filter_dialog import BOMFilterConfigDialog


@pytest.fixture(scope="session")
def qapp():
    """Ensure a single QApplication instance exists for GUI tests."""
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def temp_mgr_and_dialog(qapp):
    """Create isolated BOMFilterManager and BOMFilterConfigDialog."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "test_ssbom_master.db"
        mgr = BOMFilterManager(base_dir=tmp_path, remote_db_path=db_path)
        dlg = BOMFilterConfigDialog(filter_manager=mgr, initial_model="Virgo")
        yield mgr, dlg
        dlg.close()


class TestBOMFilterConfigDialog:
    """Test suite for BOMFilterConfigDialog UI operations."""

    def test_dialog_initialization(self, temp_mgr_and_dialog):
        """Verify dialog initializes with models and loads rules for Virgo."""
        mgr, dlg = temp_mgr_and_dialog

        assert dlg.model_list.count() >= 6
        assert dlg.current_model == "Virgo"
        assert dlg.rules_table.rowCount() == 50

    def test_search_rules_filtering(self, temp_mgr_and_dialog):
        """Verify filtering rules by typing into search box."""
        mgr, dlg = temp_mgr_and_dialog

        # Search for 'WASTE'
        dlg.rule_search_box.setText("WASTE")
        assert dlg.rules_table.rowCount() >= 1

        # Search for non-existent item
        dlg.rule_search_box.setText("NON_EXISTENT_PART_XYZ_123")
        assert dlg.rules_table.rowCount() == 0

        # Clear search
        dlg.rule_search_box.setText("")
        assert dlg.rules_table.rowCount() == 50

    def test_add_and_edit_rule_via_form(self, temp_mgr_and_dialog):
        """Verify adding and editing a rule using the dialog's form."""
        mgr, dlg = temp_mgr_and_dialog

        # 1. Add new rule
        dlg.txt_item_name.setText("TEST NEW SENSOR")
        dlg.combo_match_mode.setCurrentIndex(1)  # Part_name
        dlg.txt_part_code.setText("302TEST001")
        dlg.txt_notes.setText("Test note from unit test")
        dlg._on_save_rule()

        assert dlg.rules_table.rowCount() == 51

        # Check DB
        virgo_rules = mgr.get_rules_for_model("Virgo")
        assert len(virgo_rules) == 51
        new_rule = [r for r in virgo_rules if r.item_name == "TEST NEW SENSOR"][0]
        assert new_rule.match_mode == "Part_name"
        assert new_rule.part_code == "302TEST001"

        # 2. Edit the rule
        dlg._start_edit_rule(new_rule.id)
        assert dlg.editing_rule_id == new_rule.id
        dlg.txt_item_name.setText("TEST NEW SENSOR UPDATED")
        dlg._on_save_rule()

        updated_rule = [r for r in mgr.get_rules_for_model("Virgo") if r.id == new_rule.id][0]
        assert updated_rule.item_name == "TEST NEW SENSOR UPDATED"

        # 3. Delete rule
        del_success = mgr.delete_rule(new_rule.id)
        assert del_success is True
        dlg._refresh_rules_table()
        assert dlg.rules_table.rowCount() == 50
