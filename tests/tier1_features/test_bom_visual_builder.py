"""Unit and GUI tests for BOMVisualRuleBuilderDialog."""

import os
import tempfile
from pathlib import Path
from unittest.mock import patch
import pytest
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QComboBox

from src.core.bom_filter_manager import BOMFilterManager
from src.gui.bom_visual_builder_dialog import BOMVisualRuleBuilderDialog


@pytest.fixture(scope="session")
def qapp():
    """Ensure a single QApplication instance exists for GUI tests."""
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def visual_builder_fixture(qapp, sample_plm_excel_file: Path):
    """Create isolated BOMFilterManager and BOMVisualRuleBuilderDialog with sample PLM file."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "test_ssbom_master.db"
        mgr = BOMFilterManager(base_dir=tmp_path, remote_db_path=db_path)
        dlg = BOMVisualRuleBuilderDialog(
            filter_manager=mgr,
            initial_model="TestModel",
            initial_file=sample_plm_excel_file,
        )
        yield mgr, dlg, sample_plm_excel_file
        dlg.close()


class TestBOMVisualRuleBuilderDialog:
    """Test suite for BOMVisualRuleBuilderDialog."""

    def test_tree_construction_from_plm_file(self, visual_builder_fixture):
        """Verify PLM file is parsed and rendered accurately into QTreeWidget."""
        mgr, dlg, sample_file = visual_builder_fixture

        assert dlg.tree_widget.topLevelItemCount() == 3
        # Root 1: LSU UNIT
        root0 = dlg.tree_widget.topLevelItem(0)
        assert root0.text(0) == "LSU UNIT"
        assert root0.text(1) == "L1"
        assert root0.text(2) == "302FP93010"
        assert root0.text(5) == "Có"
        assert root0.childCount() == 2

        # Child: POLYGON MOTOR ASSY
        polygon = root0.child(0)
        assert polygon.text(0) == "POLYGON MOTOR ASSY"
        assert polygon.text(1) == "L2"
        assert polygon.childCount() == 2

    def test_rule2_prune_children_simulation(self, visual_builder_fixture):
        """Verify Rule 2 (cut children, keep assembly) dims child nodes and updates stats."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Mark POLYGON MOTOR ASSY as 'prune_children'
        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")

        assert dlg.rule_actions.get("POLYGON MOTOR ASSY") == "prune_children"
        assert len(dlg.rule_actions) == 1

        # Check that children of polygon are dimmed (strikethrough font)
        root0 = dlg.tree_widget.topLevelItem(0)
        polygon = root0.child(0)
        bracket = polygon.child(0)
        screw = polygon.child(1)

        assert bracket.font(0).strikeOut() is True
        assert screw.font(0).strikeOut() is True
        # Polygon assembly itself is kept (not strikethrough)
        assert polygon.font(0).strikeOut() is False

        # Check KPI label
        assert "Quy tắc đã chọn: 1 quy tắc" in dlg.lbl_stats_rules.text()
        assert "Sau khi lọc:" in dlg.lbl_stats_filtered.text()

    def test_rule1_prune_node_simulation(self, visual_builder_fixture):
        """Verify Rule 1 (prune assembly and its children) dims the assembly itself."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Mark FUSER UNIT as 'prune_node'
        dlg._on_action_changed("FUSER UNIT", "prune_node")

        assert dlg.rule_actions.get("FUSER UNIT") == "prune_node"

        root1 = dlg.tree_widget.topLevelItem(1)
        assert root1.text(0) == "FUSER UNIT"
        assert root1.font(0).strikeOut() is True

        # Child HEATER LAMP must also be strikethrough
        heater = root1.child(0)
        assert heater.font(0).strikeOut() is True

    def test_hide_pruned_nodes_toggle(self, visual_builder_fixture):
        """Verify checking 'Hide pruned nodes' actually sets isHidden=True on pruned nodes."""
        mgr, dlg, sample_file = visual_builder_fixture

        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")
        root0 = dlg.tree_widget.topLevelItem(0)
        polygon = root0.child(0)
        bracket = polygon.child(0)

        # Before checking hide: bracket is visible but dimmed
        dlg.chk_hide_pruned.setChecked(False)
        assert bracket.isHidden() is False

        # After checking hide: bracket is hidden
        dlg.chk_hide_pruned.setChecked(True)
        assert bracket.isHidden() is True

    def test_save_rules_to_db(self, visual_builder_fixture):
        """Verify saving rules persists records into SQLite bolocbom_rules."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Add 2 rules
        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")
        dlg._on_action_changed("FUSER UNIT", "prune_node")

        # Mock QMessageBox.question and QMessageBox.information
        with patch("PyQt6.QtWidgets.QMessageBox.question", return_value=pytest.importorskip("PyQt6.QtWidgets").QMessageBox.StandardButton.Yes):
            with patch("PyQt6.QtWidgets.QMessageBox.information"):
                dlg._on_save_rules()

        # Verify in DB
        saved_rules = mgr.get_rules_for_model("TestModel")
        assert len(saved_rules) == 2
        names = {r.item_name for r in saved_rules}
        assert "POLYGON MOTOR ASSY" in names
        assert "FUSER UNIT" in names
