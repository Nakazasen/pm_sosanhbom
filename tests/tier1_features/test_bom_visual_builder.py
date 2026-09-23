"""Unit and GUI tests for BOMVisualRuleBuilderDialog."""

import os
from pathlib import Path
import tempfile
from unittest.mock import patch

import openpyxl
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

    def test_branch_isolated_rule_selection(self, visual_builder_fixture):
        """Verify Issue 2 fix: selecting a rule on one branch does NOT affect other branches."""
        mgr, dlg, sample_file = visual_builder_fixture

        root0 = dlg.tree_widget.topLevelItem(0)  # LSU UNIT
        polygon = root0.child(0)                 # POLYGON MOTOR ASSY
        lens = root0.child(1)                    # F-THETA LENS
        root1 = dlg.tree_widget.topLevelItem(1)  # FUSER UNIT
        heater = root1.child(0)                  # HEATER LAMP 220V

        combo_polygon = dlg.tree_widget.itemWidget(polygon, 6)
        combo_lens = dlg.tree_widget.itemWidget(lens, 6)
        combo_root1 = dlg.tree_widget.itemWidget(root1, 6)
        combo_heater = dlg.tree_widget.itemWidget(heater, 6)

        # Before action change: all are none
        assert combo_polygon.currentIndex() == 0
        assert combo_lens.currentIndex() == 0
        assert combo_root1.currentIndex() == 0
        assert combo_heater.currentIndex() == 0

        # Change polygon via combobox (user interaction in GUI)
        combo_polygon.setCurrentIndex(1)  # "✂️ Cắt con (Rule 2)"

        # Verify ONLY polygon combo changed; others remain completely untouched
        assert combo_polygon.currentIndex() == 1
        assert combo_lens.currentIndex() == 0
        assert combo_root1.currentIndex() == 0
        assert combo_heater.currentIndex() == 0

        # Verify visual impact is strictly confined to polygon's branch:
        # Polygon itself is kept (highlighted, not strikeOut)
        assert polygon.font(0).strikeOut() is False
        # Polygon's children are pruned (strikeOut)
        assert polygon.child(0).font(0).strikeOut() is True
        assert polygon.child(1).font(0).strikeOut() is True

        # Sibling lens is unaffected (NOT strikeOut)
        assert lens.font(0).strikeOut() is False

        # Other tree root and its children are unaffected (NOT strikeOut)
        assert root1.font(0).strikeOut() is False
        assert heater.font(0).strikeOut() is False

        # Verify rule count reflects only 1 rule configured
        assert "Quy tắc đã chọn: 1 quy tắc" in dlg.lbl_stats_rules.text()

    def test_tree_construction_from_tc2412_24col_file(self, qapp, tmp_path: Path):
        """Verify Issue 1 fix: TC2412 24-col file extracts real item names, not 'Released'."""
        excel_path = tmp_path / "tc2412_sample.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        headers = [
            "Level", "Level", "Item Type", "Name", "1st parts", "Quantity",
            "2nd BOM Flag", "Parts Text", "Notice No", "Revision", "Release Status",
            "Unit Of Measure", "Date Released", "Is Variant Item", "Assembly Indicator",
            "Parts Text", "Element Effectivities", "Occurrence Name", "Finish Type",
            "Safety Part", "Multibody", "Spare Part", "Technology Intent", "TCUID"
        ]
        ws.append(headers)

        rows = [
            ["0", 0, "Parts", "110C3F3NL0", None, None, None, "ECOSYS MA3500fx 220-240V", None, "01", "Released", None, "29-Aug-2024 09:22", "False", "Fixed Assembly", None, None, None, "None", "False", "False", "True", "In-house Design", "uid0"],
            ["1", 1, "Parts", "3VC3FP0010", None, "1.000", "False", "SET ASSY ELEMENTS NL", "N00132693", "01", "Released", "/Piece", "02-Sep-2024 10:31", "False", "Fixed Assembly", None, None, None, "None", "False", "False", "True", "In-house Design", "uid1"],
            ["2", 2, "Parts", "3VC1B58010", None, "1.000", "False", "FRAME UNIT", "N00132043", "05", "Released", "/Piece", "08-Sep-2023 09:00", "False", "Fixed Assembly", None, None, None, "None", "False", "False", "True", "In-house Design", "uid2"],
            ["1", 1, "Parts", "3VC3FP1010", None, "1.000", "False", "PACKING MATTERS NL", "N00133434", "03", None, "/Piece", None, "False", "Fixed Assembly", None, None, None, "None", "False", "False", "True", "In-house Design", "uid3"],
        ]
        for r in rows:
            ws.append(r)
        wb.save(excel_path)
        wb.close()

        db_path = tmp_path / "test_tc2412.db"
        mgr = BOMFilterManager(base_dir=tmp_path, remote_db_path=db_path)
        dlg = BOMVisualRuleBuilderDialog(
            filter_manager=mgr,
            initial_model="Virgo",
            initial_file=excel_path,
        )

        try:
            assert dlg.tree_widget.topLevelItemCount() == 1
            root = dlg.tree_widget.topLevelItem(0)

            # Root checks
            assert root.text(0) == "ECOSYS MA3500fx 220-240V"
            assert root.text(0) != "Released"
            assert root.text(1) == "L0"
            assert root.text(2) == "110C3F3NL0"
            assert root.text(3) == "01"
            assert root.text(3) != "29-Aug-2024 09:22"

            # Child 0: SET ASSY ELEMENTS NL
            c0 = root.child(0)
            assert c0.text(0) == "SET ASSY ELEMENTS NL"
            assert c0.text(0) != "Released"
            assert c0.text(2) == "3VC3FP0010"
            assert c0.text(3) == "01"

            # Child 1: PACKING MATTERS NL (even when Release Status was None)
            c1 = root.child(1)
            assert c1.text(0) == "PACKING MATTERS NL"
            assert c1.text(0) != "(Chưa đặt tên)"
            assert c1.text(2) == "3VC3FP1010"
            assert c1.text(3) == "03"
        finally:
            dlg.close()

    def test_branch_targeted_saving_and_model_pruner(self, visual_builder_fixture):
        """Verify that branch-targeted rules save parent branch in notes and ModelPruner respects it."""
        from src.core.model_pruner import ModelPruner

        mgr, dlg, sample_file = visual_builder_fixture

        root0 = dlg.tree_widget.topLevelItem(0)  # LSU UNIT (302FP93010)
        polygon = root0.child(0)                 # POLYGON MOTOR ASSY (302FP94010)

        combo_polygon = dlg.tree_widget.itemWidget(polygon, 6)
        combo_polygon.setCurrentIndex(1)  # prune_children

        with patch("PyQt6.QtWidgets.QMessageBox.question", return_value=pytest.importorskip("PyQt6.QtWidgets").QMessageBox.StandardButton.Yes):
            with patch("PyQt6.QtWidgets.QMessageBox.information"):
                dlg._on_save_rules()

        # Check in DB
        rules = mgr.get_rules_for_model("TestModel")
        assert len(rules) == 1
        r = rules[0]
        assert r.item_name == "POLYGON MOTOR ASSY"
        assert r.part_code == "302FP94010"
        assert "[branch: 302FP93010]" in r.notes

        # Convert to ModelRule
        m_rule = r.to_model_rule()
        assert m_rule.parent_part_code == "302FP93010"
        assert m_rule.action == "prune_children"

        # Verify with ModelPruner
        pruner = ModelPruner(custom_rules={"TestModel": [m_rule]})
        tree = dlg.raw_tree.clone()
        pruned = pruner.prune_tree(tree, model_name="TestModel")

        # The polygon motor under LSU UNIT has its children pruned
        lsu = pruned.roots[0]
        poly_pruned = lsu.children[0]
        assert poly_pruned.item_id == "302FP94010"
        assert len(poly_pruned.children) == 0  # children pruned!
