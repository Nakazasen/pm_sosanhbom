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

    def test_fast_lazy_combo_creation_and_column_text(self, visual_builder_fixture):
        """Verify on-demand QComboBox creation and column 6 display text."""
        from src.gui.bom_visual_builder_dialog import FastBOMTreeWidget, ACTION_DISPLAY_TEXT

        mgr, dlg, sample_file = visual_builder_fixture
        root0 = dlg.tree_widget.topLevelItem(0)

        # Before calling itemWidget, no heavyweight QComboBox widget was allocated upfront
        raw_widget = super(FastBOMTreeWidget, dlg.tree_widget).itemWidget(root0, 6)
        assert raw_widget is None
        # But text is present in column 6
        assert root0.text(6) == ACTION_DISPLAY_TEXT["none"]

        # Accessing itemWidget creates it on-demand
        combo = dlg.tree_widget.itemWidget(root0, 6)
        assert isinstance(combo, QComboBox)
        assert combo.currentIndex() == 0

        # Changing index updates column 6 text
        combo.setCurrentIndex(1)
        assert root0.text(6) == ACTION_DISPLAY_TEXT["prune_children"]

    def test_large_scale_performance_boost(self, qapp, tmp_path: Path):
        """Verify 2000-node synthetic BOM builds, styles, and expands in under 1 second."""
        import time
        from src.core.models import BOMNode, BOMTree

        # Construct synthetic tree with 2000 nodes (10 roots, each with 199 descendants)
        roots = []
        node_id_counter = 1000
        for r_idx in range(10):
            root = BOMNode(level=1, item_id=f"ROOT_{r_idx}", item_name=f"ROOT ASSEMBLY {r_idx}", has_children=True)
            for c_idx in range(10):
                child = BOMNode(level=2, item_id=f"SUB_{node_id_counter}", item_name=f"SUB ASSY {node_id_counter}", has_children=True)
                node_id_counter += 1
                for g_idx in range(18):
                    leaf = BOMNode(level=3, item_id=f"PART_{node_id_counter}", item_name=f"LEAF PART {node_id_counter}", has_children=False)
                    node_id_counter += 1
                    child.add_child(leaf)
                root.add_child(child)
            roots.append(root)

        synthetic_tree = BOMTree(roots=roots)
        assert len(synthetic_tree.flatten()) > 1900

        dlg = BOMVisualRuleBuilderDialog(initial_model="TestModel")
        try:
            dlg.raw_tree = synthetic_tree
            dlg._load_existing_model_rules()

            t0 = time.time()
            dlg._build_tree_widget()
            dlg._expand_to_level_2()
            dlg._apply_simulation_preview()
            elapsed = time.time() - t0

            # Must build, expand, and simulate thousands of nodes in < 0.8s
            assert elapsed < 0.8, f"Tree population took too long: {elapsed:.3f}s"
            assert dlg.tree_widget.topLevelItemCount() == 10
        finally:
            dlg.close()

    def test_search_tree_clearing_restores_hidden_nodes(self, visual_builder_fixture):
        """Verify that clearing search text unhides nodes and clears search highlights."""
        mgr, dlg, sample_file = visual_builder_fixture
        root0 = dlg.tree_widget.topLevelItem(0)

        # Before search: root0 is visible
        assert root0.isHidden() is False

        # Search for non-matching query
        dlg._on_search_tree("NONEXISTENT_QUERY_12345")
        assert root0.isHidden() is True

        # Clear search query: root0 must be restored to visible
        dlg._on_search_tree("")
        assert root0.isHidden() is False

    def test_model_change_reloads_tree_rules(self, visual_builder_fixture):
        """Verify that changing model in combo box re-evaluates tree rules on loaded BOM."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Add a rule for AnotherModel
        mgr.add_rule(
            model_name="AnotherModel",
            item_name="LSU UNIT",
            match_mode="Full_name",
            notes="Tạo tự động (Cắt con)",
        )

        # Switch model in dialog
        dlg.combo_model.setCurrentText("AnotherModel")
        assert dlg.current_model == "AnotherModel"
        assert len(dlg.existing_model_rules) == 1
        assert "Quy tắc đã chọn: 1 quy tắc" in dlg.lbl_stats_rules.text()

    def test_tc_active_workspace_27col_tree_parsing(self, tmp_path: Path):
        """Verify that 27-column Teamcenter Active Workspace files parse correctly."""
        from src.core.tree_parser import BOMTreeParser
        excel_path = tmp_path / "tc_aw_27col.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        headers = [
            "Level", "ID", "Revision", "Revision Name", "Description", "Reference Designator",
            "Find Number", "Quantity", "Unit Of Measure", "Date Released", "Release Effectivity",
            "Owner", "Group ID", "Last Modifying User", "Assembly Indicator", "Part Required",
            "Is Part Aligned", "Aligned Parts", "Aligned Part Contexts", "Aligned Part Release Status",
            "Design Required", "Is Design Aligned", "Aligned Designs", "Aligned Design Contexts",
            "Aligned Design Release Status", "Release Status", "TCUID"
        ]
        ws.append(headers)
        rows = [
            ["0", "MBC_001", "01", "110C2K2US0", "VIRGO MAIN MACHINE", None, None, None, None, "06-Mar-2024", None, None, None, None, "Fixed Assembly", None, None, None, None, None, None, None, None, None, None, "Released", "uid0"],
            ["1", "MBC_002", "01", "3VC2KP0020", "SET ASSY ELEMENTS US", None, "10", "1.000", "/Piece", "25-Mar-2024", None, None, None, None, "Fixed Assembly", None, None, None, None, None, None, None, None, None, None, "Released", "uid1"],
            ["2", "MBC_003", "01", "302XC00243", None, None, "30", "1.000", "/Piece", "27-Jul-2023", None, None, None, None, "Fixed Assembly", None, None, None, None, None, None, None, None, None, None, "Released", "uid2"],
        ]
        for r in rows:
            ws.append(r)
        wb.save(excel_path)
        wb.close()

        parser = BOMTreeParser()
        tree = parser.parse_file(excel_path)
        assert len(tree.roots) == 1
        root = tree.roots[0]
        assert root.item_id == "110C2K2US0"
        assert root.item_name == "VIRGO MAIN MACHINE"
        assert root.level == 0
        assert len(root.children) == 1

        child = root.children[0]
        assert child.item_id == "3VC2KP0020"
        assert child.item_name == "SET ASSY ELEMENTS US"
        assert child.level == 1
        assert len(child.children) == 1

        subchild = child.children[0]
        assert subchild.item_id == "302XC00243"
        assert subchild.level == 2

    def test_fast_tree_widget_keyboard_popup(self, qapp):
        """Verify Space and Return keys invoke action popup on selected tree item."""
        from PyQt6.QtGui import QKeyEvent
        from PyQt6.QtWidgets import QTreeWidgetItem
        from src.gui.bom_visual_builder_dialog import FastBOMTreeWidget

        tree = FastBOMTreeWidget()
        item = QTreeWidgetItem(["Item 1", "L1", "P001", "01", "1", "No", "— Bình thường (Giữ)"])
        item.setData(0, Qt.ItemDataRole.UserRole, {"node_key": "k1", "item_name": "Item 1", "item_id": "P001"})
        tree.addTopLevelItem(item)
        tree.setCurrentItem(item)

        # Trigger Space key
        key_event = QKeyEvent(QKeyEvent.Type.KeyPress, Qt.Key.Key_Space, Qt.KeyboardModifier.NoModifier)
        tree.keyPressEvent(key_event)

        combo = tree.itemWidget(item, 6)
        assert isinstance(combo, QComboBox)

    def test_search_by_level_variations(self, visual_builder_fixture):
        """Verify search box supports L1, L2, cấp 1, cấp 2, level 2, and combined level+keyword."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Search 'L1'
        dlg._on_search_tree("L1")
        assert len(dlg._search_matching_items) == 3
        names_l1 = {it.text(0) for it in dlg._search_matching_items}
        assert "LSU UNIT" in names_l1
        assert "FUSER UNIT" in names_l1

        # Search 'L2'
        dlg._on_search_tree("L2")
        assert len(dlg._search_matching_items) >= 2
        for it in dlg._search_matching_items:
            assert it.text(1) == "L2"

        # Search 'cấp 2'
        dlg._on_search_tree("cấp 2")
        assert len(dlg._search_matching_items) >= 2
        for it in dlg._search_matching_items:
            assert it.text(1) == "L2"

        # Search 'cap 1' (unaccented Vietnamese)
        dlg._on_search_tree("cap 1")
        assert len(dlg._search_matching_items) == 3

        # Search 'level 2'
        dlg._on_search_tree("level 2")
        assert len(dlg._search_matching_items) >= 2

        # Combined level + keyword: 'L2 POLYGON'
        dlg._on_search_tree("L2 POLYGON")
        assert len(dlg._search_matching_items) == 1
        assert dlg._search_matching_items[0].text(0) == "POLYGON MOTOR ASSY"

    def test_search_phantom_and_has_children(self, visual_builder_fixture):
        """Verify search box supports #phantom, phantom, #has_children, có con, and không con."""
        mgr, dlg, sample_file = visual_builder_fixture

        # '#phantom'
        dlg._on_search_tree("#phantom")
        assert len(dlg._search_matching_items) > 0
        for it in dlg._search_matching_items:
            data = it.data(0, Qt.ItemDataRole.UserRole)
            assert data.get("has_children") is True or "phantom" in it.text(0).lower()

        # 'có con' (Vietnamese for has children)
        dlg._on_search_tree("có con")
        assert len(dlg._search_matching_items) > 0
        for it in dlg._search_matching_items:
            assert it.text(5) == "Có"

        # 'không con' (Vietnamese for leaf nodes)
        dlg._on_search_tree("không con")
        assert len(dlg._search_matching_items) > 0
        for it in dlg._search_matching_items:
            assert it.text(5) == "Không"

    def test_search_by_rule_action(self, visual_builder_fixture):
        """Verify search box matches configured rule actions like cắt con, bỏ cả cụm, đã chọn."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Configure 2 rules
        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")
        dlg._on_action_changed("FUSER UNIT", "prune_node")

        # Search 'cắt con'
        dlg._on_search_tree("cắt con")
        assert len(dlg._search_matching_items) == 1
        assert dlg._search_matching_items[0].text(0) == "POLYGON MOTOR ASSY"

        # Search 'bỏ cả cụm'
        dlg._on_search_tree("bỏ cả cụm")
        assert len(dlg._search_matching_items) == 1
        assert dlg._search_matching_items[0].text(0) == "FUSER UNIT"

        # Search 'đã chọn'
        dlg._on_search_tree("đã chọn")
        assert len(dlg._search_matching_items) == 2
        names = {it.text(0) for it in dlg._search_matching_items}
        assert names == {"POLYGON MOTOR ASSY", "FUSER UNIT"}

    def test_enter_key_cycles_matches_and_never_triggers_file_browser(self, visual_builder_fixture):
        """Verify Enter cycles through search matches without opening file browser."""
        from unittest.mock import MagicMock
        from PyQt6.QtGui import QKeyEvent

        mgr, dlg, sample_file = visual_builder_fixture

        # Verify autoDefault safety
        assert dlg.btn_load_file.autoDefault() is False
        assert dlg.btn_load_file.isDefault() is False
        assert dlg.btn_save.autoDefault() is False
        assert dlg.btn_save.isDefault() is False

        # Mock _on_browse_file to ensure it is NEVER called
        browse_mock = MagicMock()
        dlg._on_browse_file = browse_mock

        # Search for 'ASSY' which has multiple matches
        dlg.txt_search.setText("ASSY")
        total_matches = len(dlg._search_matching_items)
        assert total_matches >= 2
        assert dlg._search_match_index == 0

        # Press Enter on txt_search
        dlg.txt_search.setFocus()
        enter_event = QKeyEvent(QKeyEvent.Type.KeyPress, Qt.Key.Key_Return, Qt.KeyboardModifier.NoModifier)
        dlg.keyPressEvent(enter_event)

        # Advanced to next match
        assert dlg._search_match_index == 1
        # browse_file was NOT called
        browse_mock.assert_not_called()

        # Press Shift+Enter to cycle backwards
        shift_enter_event = QKeyEvent(
            QKeyEvent.Type.KeyPress, Qt.Key.Key_Return, Qt.KeyboardModifier.ShiftModifier
        )
        dlg.keyPressEvent(shift_enter_event)
        assert dlg._search_match_index == 0
        browse_mock.assert_not_called()

    def test_rule_review_audit_mode_toggle_and_filter(self, visual_builder_fixture):
        """Verify Rule Review / Audit Mode shows only configured nodes and updates count."""
        mgr, dlg, sample_file = visual_builder_fixture

        # Initially 0 rules
        assert dlg.chk_review_rules.text() == "📋 Chỉ xem cụm đã chọn quy tắc (0)"

        # Set 2 rules
        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")
        dlg._on_action_changed("FUSER UNIT", "prune_node")
        assert dlg.chk_review_rules.text() == "📋 Chỉ xem cụm đã chọn quy tắc (2)"

        # Enable Review Mode
        dlg.chk_review_rules.setChecked(True)
        assert dlg.combo_rule_filter.isEnabled() is True

        root0 = dlg.tree_widget.topLevelItem(0)  # LSU UNIT (ancestor of polygon)
        polygon = root0.child(0)                 # POLYGON MOTOR ASSY (configured)
        lens = root0.child(1)                    # F-THETA LENS (unconfigured sibling)
        root1 = dlg.tree_widget.topLevelItem(1)  # FUSER UNIT (configured)
        root2 = dlg.tree_widget.topLevelItem(2)  # High voltage (unconfigured root)

        # Polygon and FUSER are visible
        assert polygon.isHidden() is False
        assert root1.isHidden() is False
        # Ancestor LSU UNIT is visible and expanded to reveal polygon
        assert root0.isHidden() is False
        assert root0.isExpanded() is True
        # Unconfigured sibling lens and unconfigured root2 are hidden
        assert lens.isHidden() is True
        assert root2.isHidden() is True

        # Filter specifically to "Chỉ Cắt con" (Rule 2)
        dlg.combo_rule_filter.setCurrentIndex(1)
        assert polygon.isHidden() is False
        assert root1.isHidden() is True

        # Filter specifically to "Chỉ Bỏ cả cụm" (Rule 1)
        dlg.combo_rule_filter.setCurrentIndex(2)
        assert polygon.isHidden() is True
        assert root1.isHidden() is False

        # Turn OFF review mode: everything is restored
        dlg.chk_review_rules.setChecked(False)
        assert root2.isHidden() is False
        assert lens.isHidden() is False

    def test_visual_builder_syncs_rules_to_parent_filter_dialog(self, qapp, sample_plm_excel_file: Path, tmp_path: Path):
        """Verify saving rules from Visual Builder updates parent BOMFilterConfigDialog rules table instantly."""
        from src.gui.bom_filter_dialog import BOMFilterConfigDialog

        db_path = tmp_path / "sync_test.db"
        mgr = BOMFilterManager(base_dir=tmp_path, remote_db_path=db_path)

        parent_dlg = BOMFilterConfigDialog(filter_manager=mgr, initial_model="TestModel")
        try:
            assert parent_dlg.current_model != "TestModel"

            builder_dlg = BOMVisualRuleBuilderDialog(
                filter_manager=mgr,
                initial_model="TestModel",
                initial_file=sample_plm_excel_file,
                parent=parent_dlg,
            )
            builder_dlg.rules_saved.connect(parent_dlg._on_visual_rules_saved)

            # Configure a rule on the tree
            builder_dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")

            # Save rules with mocked confirmation
            with patch("PyQt6.QtWidgets.QMessageBox.question", return_value=pytest.importorskip("PyQt6.QtWidgets").QMessageBox.StandardButton.Yes):
                with patch("PyQt6.QtWidgets.QMessageBox.information"):
                    builder_dlg._on_save_rules()

            # Verify parent dialog immediately received signal and updated table 2
            assert parent_dlg.current_model == "TestModel"
            assert parent_dlg.rules_table.rowCount() == 1
            assert parent_dlg.rules_table.item(0, 1).text() == "POLYGON MOTOR ASSY"
            assert "Cắt con" in parent_dlg.rules_table.item(0, 4).text()
        finally:
            parent_dlg.close()

    def test_part_code_search_not_hijacked_by_level_regex(self, visual_builder_fixture):
        """Verify that part codes like L1024-001, L1-MOTOR, or names like CAP 10UF match without false level rejection."""
        from src.gui.bom_visual_builder_dialog import _node_matches_query

        # Part with L1024-001 at level 2
        assert _node_matches_query("POLYGON MOTOR", "L1024-001", 2, False, "", "none", "L1024-001") is True
        # Part with L1-MOTOR at level 2
        assert _node_matches_query("POLYGON MOTOR", "L1-MOTOR", 2, False, "", "none", "L1-MOTOR") is True
        # Part with CAP 100UF at level 2
        assert _node_matches_query("CAP 100UF", "C100", 2, False, "", "none", "CAP 100UF") is True
        # Cụm ảo synonym matching
        assert _node_matches_query("CỤM TRỐNG DRUM", "DRUM-01", 2, True, "", "none", "cụm ảo") is True
        # Bỏ cụm synonym matching
        assert _node_matches_query("FUSER UNIT", "FU-01", 1, True, "", "prune_node", "bỏ cụm") is True

    def test_enter_key_respects_focused_button_and_does_not_hijack(self, visual_builder_fixture):
        """Verify that when a button or control is focused, Enter activates that button instead of cycling search."""
        from PyQt6.QtGui import QKeyEvent

        mgr, dlg, sample_file = visual_builder_fixture
        dlg.txt_search.setText("ASSY")
        assert len(dlg._search_matching_items) >= 2

        # 1. Test with btn_expand_l2 (non-modal)
        expanded_called = False
        def on_expand():
            nonlocal expanded_called
            expanded_called = True
        dlg.btn_expand_l2.clicked.connect(on_expand)

        dlg.btn_expand_l2.setFocus()
        enter_event = QKeyEvent(QKeyEvent.Type.KeyPress, Qt.Key.Key_Return, Qt.KeyboardModifier.NoModifier)
        dlg.keyPressEvent(enter_event)
        assert expanded_called is True

        # 2. Test with btn_save (with mocked warning)
        with patch("PyQt6.QtWidgets.QMessageBox.warning") as mock_warn:
            dlg.btn_save.setFocus()
            enter_event2 = QKeyEvent(QKeyEvent.Type.KeyPress, Qt.Key.Key_Return, Qt.KeyboardModifier.NoModifier)
            dlg.keyPressEvent(enter_event2)
            mock_warn.assert_called_once()

    def test_rule1_preview_styling_active_on_target_node(self, visual_builder_fixture):
        """Verify that Rule 1 (prune_node) target node receives rule1 styling (_set_item_rule1_highlight), not generic dimmed."""
        mgr, dlg, sample_file = visual_builder_fixture
        root1 = dlg.tree_widget.topLevelItem(1)  # FUSER UNIT

        dlg._on_action_changed("FUSER UNIT", "prune_node")
        assert getattr(root1, "_visual_state", None) == "rule1"
        assert root1.font(0).strikeOut() is True

    def test_review_mode_allows_child_inspection_when_expanded(self, visual_builder_fixture):
        """Verify that in review mode, rule nodes start collapsed, but when expanded their children are visible for inspection."""
        mgr, dlg, sample_file = visual_builder_fixture
        root0 = dlg.tree_widget.topLevelItem(0)  # LSU UNIT
        polygon = root0.child(0)                 # POLYGON MOTOR ASSY

        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")
        dlg.chk_review_rules.setChecked(True)

        # Polygon itself is visible and starts collapsed in review mode
        assert polygon.isHidden() is False
        assert polygon.isExpanded() is False

        # Its child (dimmed pruned component) is unhidden so user can inspect upon expansion
        poly_child = polygon.child(0)
        assert poly_child.isHidden() is False
        # And child has dimmed visual state
        assert getattr(poly_child, "_visual_state", None) == "dimmed"

    def test_save_rules_preserves_external_module_rules(self, visual_builder_fixture):
        """Verify that saving rules from a sub-BOM preserves existing rules belonging to other modules of the model."""
        mgr, dlg, sample_file = visual_builder_fixture
        model = "TestModel"

        # Pre-seed a rule for an unrepresented module (e.g. CASSETTE UNIT not in this BOM file)
        mgr.add_rule(
            model_name=model,
            item_name="PWB CASSETTE ASSY",
            match_mode="Full_name",
            part_code="999CASSETTE",
            notes="Module khác đã tạo trước đó",
        )
        dlg.existing_model_rules = mgr.get_rules_for_model(model)

        # Configure rule in current visual session
        dlg._on_action_changed("POLYGON MOTOR ASSY", "prune_children")

        with patch("PyQt6.QtWidgets.QMessageBox.question", return_value=pytest.importorskip("PyQt6.QtWidgets").QMessageBox.StandardButton.Yes):
            with patch("PyQt6.QtWidgets.QMessageBox.information"):
                dlg._on_save_rules()

        # Verify that BOTH the current visual rule AND the external module rule are preserved in DB
        rules_in_db = mgr.get_rules_for_model(model)
        names = {r.item_name for r in rules_in_db}
        assert "POLYGON MOTOR ASSY" in names
        assert "PWB CASSETTE ASSY" in names


