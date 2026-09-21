"""Tests for BOMFilterManager and dynamic ModelPruner integration."""

import tempfile
from pathlib import Path
import pytest

from src.core.bom_filter_manager import BOMFilterManager
from src.core.model_pruner import ModelPruner
from src.core.models import BOMNode, BOMTree


@pytest.fixture
def temp_filter_manager():
    """Create a BOMFilterManager with an isolated temporary database."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "test_ssbom_master.db"
        mgr = BOMFilterManager(base_dir=tmp_path, remote_db_path=db_path)
        yield mgr


class TestBOMFilterManager:
    """Test suite for BOMFilterManager operations."""

    def test_seeding_default_rules(self, temp_filter_manager):
        """Verify initial database creation seeds all 357 rules across 6 models."""
        models = temp_filter_manager.get_all_models()
        assert len(models) >= 6
        for m in ["Virgo", "Libra2", "Iris2024", "Sirius2", "Mebius", "Polaris"]:
            assert m in models

        virgo_rules = temp_filter_manager.get_rules_for_model("Virgo")
        assert len(virgo_rules) == 50

        iris_rules = temp_filter_manager.get_rules_for_model("Iris2024")
        assert len(iris_rules) == 134

    def test_add_update_delete_rule(self, temp_filter_manager):
        """Verify adding, updating, and deleting a rule for a model."""
        # Add rule for a new model
        rule_id = temp_filter_manager.add_rule(
            model_name="Mercury",
            item_name="COVER FRONT ASSY",
            match_mode="Full_name",
            part_code="302ABC1234",
            notes="Test rule",
        )
        assert rule_id > 0

        rules = temp_filter_manager.get_rules_for_model("Mercury")
        assert len(rules) == 1
        assert rules[0].item_name == "COVER FRONT ASSY"
        assert rules[0].part_code == "302ABC1234"

        # Update rule
        success = temp_filter_manager.update_rule(
            rule_id=rule_id,
            item_name="COVER FRONT ASSY MODIFIED",
            match_mode="Part_name",
            part_code="302XYZ9999",
            notes="Updated note",
        )
        assert success is True

        updated_rules = temp_filter_manager.get_rules_for_model("Mercury")
        assert len(updated_rules) == 1
        assert updated_rules[0].item_name == "COVER FRONT ASSY MODIFIED"
        assert updated_rules[0].match_mode == "Part_name"
        assert updated_rules[0].part_code == "302XYZ9999"

        # Delete rule
        del_success = temp_filter_manager.delete_rule(rule_id)
        assert del_success is True
        assert len(temp_filter_manager.get_rules_for_model("Mercury")) == 0

    def test_export_and_import_excel(self, temp_filter_manager, tmp_path):
        """Verify export to Excel and re-importing rules."""
        export_file = tmp_path / "test_export_bolocbom.xlsx"
        exported_path = temp_filter_manager.export_to_excel(export_file, model_name="Virgo")
        assert exported_path.exists()

        # Add a custom rule to a new model in DB
        temp_filter_manager.add_rule(
            model_name="CustomExcelModel",
            item_name="SPECIAL MOTOR",
            match_mode="Part_name",
            part_code="12345",
        )
        all_export = tmp_path / "all_rules.xlsx"
        temp_filter_manager.export_to_excel(all_export)
        assert all_export.exists()

        # Import into another clean manager
        clean_dir = tmp_path / "clean_mgr"
        clean_dir.mkdir()
        clean_db = clean_dir / "clean.db"
        clean_mgr = BOMFilterManager(base_dir=clean_dir, remote_db_path=clean_db)
        # Clear clean_mgr
        clean_mgr.delete_model("Virgo")
        assert len(clean_mgr.get_rules_for_model("Virgo")) == 0

        # Re-import from exported file
        imported_count = clean_mgr.import_from_excel(export_file)
        assert imported_count == 50
        assert len(clean_mgr.get_rules_for_model("Virgo")) == 50

    def test_reset_model_to_default(self, temp_filter_manager):
        """Verify resetting a modified model back to factory defaults."""
        # Add a dummy rule to Virgo
        temp_filter_manager.add_rule("Virgo", "DUMMY RULE", "Full_name")
        assert len(temp_filter_manager.get_rules_for_model("Virgo")) == 51

        # Reset Virgo
        reset_count = temp_filter_manager.reset_model_to_default("Virgo")
        assert reset_count == 50
        assert len(temp_filter_manager.get_rules_for_model("Virgo")) == 50

    def test_model_pruner_dynamic_integration(self, temp_filter_manager):
        """Verify ModelPruner utilizes rules added to BOMFilterManager."""
        # Add a custom rule for a new model 'Titan'
        temp_filter_manager.add_rule(
            model_name="Titan",
            item_name="TITAN COVER",
            match_mode="Full_name",
        )

        pruner = ModelPruner(filter_manager=temp_filter_manager)
        rules = pruner.get_rules_for_model("Titan")
        assert len(rules) == 1
        assert rules[0].item_name == "TITAN COVER"

        # Apply rule to prune tree
        root = BOMNode(level=1, item_id="ROOT", item_name="ROOT")
        leaf1 = BOMNode(level=2, item_id="T01", item_name="TITAN COVER", has_children=False)
        leaf2 = BOMNode(level=2, item_id="T02", item_name="KEEP ME", has_children=False)
        root.add_child(leaf1)
        root.add_child(leaf2)

        tree = BOMTree(roots=[root])
        pruned = pruner.prune_tree(tree, model_name="Titan")
        assert pruned.size() == 2
        assert pruned.roots[0].children[0].item_id == "T02"
