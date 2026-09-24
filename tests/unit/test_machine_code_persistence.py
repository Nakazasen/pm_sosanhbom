"""Unit tests for machine code persistence and project configuration in Leader View."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.leader_view import LeaderWorkspaceView, ProjectStage


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestMachineCodePersistence:
    """Test suite ensuring deleted machine codes do not reappear after restart."""

    def test_delete_machine_persists_and_does_not_reappear(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """When a machine row is deleted, it is saved and does not reappear upon re-instantiation."""
        config_dir = tmp_path / "config"
        config_dir.mkdir(parents=True, exist_ok=True)

        # 1. First run: Start with defaults
        view1 = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view1.step1_widget

        initial_count = step1.machine_table.rowCount()
        assert initial_count >= 3

        # Record code of first row
        first_code = step1.machine_table.item(0, 1).text()
        second_code = step1.machine_table.item(1, 1).text()

        # Select first row and delete it
        step1.machine_table.setCurrentCell(0, 1)
        step1._del_machine_row()

        assert step1.machine_table.rowCount() == initial_count - 1
        assert step1.machine_table.item(0, 1).text() == second_code

        # Verify config/settings.json exists and has recorded the change
        cfg_file = config_dir / "settings.json"
        assert cfg_file.exists()
        with open(cfg_file, "r", encoding="utf-8") as f:
            cfg = json.load(f)

        assert "project_configs" in cfg
        virgo_cfg = cfg["project_configs"].get("Virgo")
        assert virgo_cfg is not None
        saved_codes = [m["code"] for m in virgo_cfg["machines"]]
        assert first_code not in saved_codes
        assert second_code in saved_codes

        # 2. Simulate next session / reopening the app with same base_dir
        view2 = LeaderWorkspaceView(base_dir=tmp_path)
        step1_session2 = view2.step1_widget

        # Deleted code MUST NOT reappear
        session2_codes = [
            step1_session2.machine_table.item(r, 1).text()
            for r in range(step1_session2.machine_table.rowCount())
        ]
        assert first_code not in session2_codes
        assert second_code in session2_codes
        assert step1_session2.machine_table.rowCount() == initial_count - 1

    def test_delete_all_machines_persists_empty_table(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """If user deletes all machine codes, reopening does not restore defaults."""
        config_dir = tmp_path / "config"
        config_dir.mkdir(parents=True, exist_ok=True)

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        while step1.machine_table.rowCount() > 0:
            step1.machine_table.setCurrentCell(0, 1)
            step1._del_machine_row()

        assert step1.machine_table.rowCount() == 0

        # Next session
        view_reopen = LeaderWorkspaceView(base_dir=tmp_path)
        step1_reopen = view_reopen.step1_widget

        assert step1_reopen.machine_table.rowCount() == 0

    def test_switching_models_persists_and_restores_independent_codes(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Each model maintains its own machine list in configuration."""
        config_dir = tmp_path / "config"
        config_dir.mkdir(parents=True, exist_ok=True)

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # Set Virgo codes
        step1.model_combo.setCurrentText("Virgo")
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110C2L9JP0")
        step1._add_machine_row(code="110C2MTTW0")

        # Switch to Iris2024
        step1.model_combo.setCurrentText("Iris2024")
        step1.machine_table.setRowCount(0)
        step1._add_machine_row(code="110IRIS001")

        # Switch back to Virgo
        step1.model_combo.setCurrentText("Virgo")
        virgo_codes = [
            step1.machine_table.item(r, 1).text()
            for r in range(step1.machine_table.rowCount())
        ]
        assert "110C2L9JP0" in virgo_codes
        assert "110C2MTTW0" in virgo_codes
        assert "110IRIS001" not in virgo_codes

    def test_machine_dict_service_custom_model_and_codes(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """MachineDictService saves custom models and codes to settings.json and retrieves them."""
        from src.services.machine_dict_service import MachineDictService

        settings_file = tmp_path / "config" / "settings.json"
        monkeypatch.setattr(
            MachineDictService,
            "_get_local_settings_path",
            lambda self: settings_file,
        )

        service = MachineDictService(excel_path=tmp_path / "non_existent.xlsx")

        # 1. Register new model
        ok = service.add_model_name("Crux2025")
        assert ok is True
        assert "Crux2025" in service.get_model_names()

        # 2. Add 4-char machine codes
        service.add_or_update_machine_code("Crux2025", "02Z8")
        service.add_or_update_machine_code("Crux2025", "0C0T")

        codes = service.get_machine_codes_for_model("Crux2025")
        assert "02Z8" in codes
        assert "0C0T" in codes

        # 3. Verify settings.json was written correctly
        assert settings_file.exists()
        with open(settings_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        assert "Crux2025" in data.get("custom_models", [])
        assert "02Z8" in data.get("custom_model_codes", {}).get("Crux2025", [])

    def test_model_machine_manager_dialog_populate_table(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """ModelMachineManagerDialog allows managing models and populates Table 1.2 with formatted codes."""
        from src.gui.leader_view import ModelMachineManagerDialog
        from src.services.machine_dict_service import MachineDictService

        settings_file = tmp_path / "config" / "settings.json"
        monkeypatch.setattr(
            MachineDictService,
            "_get_local_settings_path",
            lambda self: settings_file,
        )

        view = LeaderWorkspaceView(base_dir=tmp_path)
        step1 = view.step1_widget

        # Add custom model and code to service
        dict_service = step1.dict_service
        dict_service.add_model_name("Corvus")
        dict_service.add_or_update_machine_code("Corvus", "09AB")
        dict_service.add_or_update_machine_code("Corvus", "09CD")

        # Open dialog instance
        dlg = ModelMachineManagerDialog(
            dict_service=dict_service,
            current_model="Corvus",
            current_stage=step1.stage_combo.currentText(),
            parent=step1,
        )

        # Verify dialog loaded the model and its codes
        assert dlg.current_model == "Corvus"
        codes_text = dlg.text_codes.toPlainText()
        assert "09AB" in codes_text
        assert "09CD" in codes_text

        # Connect dialog signals to step1 slots as leader_view does
        dlg.model_selected.connect(step1._on_model_selected_from_dialog)
        dlg.populate_codes_requested.connect(step1._on_populate_codes_from_dialog)

        # Trigger population into Table 1.2
        dlg._on_populate_table_clicked()

        # Step1 machine_table must now contain the populated full codes
        table_codes = [
            step1.machine_table.item(r, 1).text()
            for r in range(step1.machine_table.rowCount())
        ]
        assert "11009ABNL0" in table_codes
        assert "11009CDNL0" in table_codes
        assert step1.model_combo.currentText() == "Corvus"

        # Verify saved to config
        with open(settings_file, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        corvus_cfg = cfg.get("project_configs", {}).get("Corvus")
        assert corvus_cfg is not None
        saved_codes = [m["code"] for m in corvus_cfg["machines"]]
        assert "11009ABNL0" in saved_codes
        assert "11009CDNL0" in saved_codes

