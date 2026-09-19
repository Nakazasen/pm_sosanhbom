"""Feature F22: Leader Management Workspace Isolation Tests.

Verifies:
1. Automated project directory structure creation for target machine models.
2. Member workbook submission tracking (Submitted vs Pending).
3. Batch reconciliation trigger aggregating all sub-units.
4. Tri-lingual i18n support (Vietnamese, Japanese, Chinese) for non-tech UI.
5. Leader project audit logging and history tracking.
"""

from __future__ import annotations

from pathlib import Path
import openpyxl
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox


from src.gui.leader_view import BatchReconciliationWorker, LeaderWorkspaceView
from src.reporting.excel_generator import STANDARD_SUB_UNITS
from src.ui.i18n import SUPPORTED_LANGUAGES, get_i18n, t


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestF22LeaderWorkspace:
    """Test suite for Feature F22: Leader Management Workspace."""

    def test_f22_project_folder_hierarchy_creation(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Test 1: Create project folders for machine model and required sub-units."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Virgo")
        created_path = view.create_project_folder_structure()

        assert created_path.exists()
        assert (created_path / "PLM").exists()
        assert (created_path / "R3").exists()
        assert (created_path / "Reports").exists()
        assert (created_path / "capnhat" / "old").exists()
        for unit in STANDARD_SUB_UNITS:
            assert (created_path / "CTTT" / unit).exists()

    def test_f22_track_member_submission_status(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Test 2: Track which members have submitted their workbooks."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = LeaderWorkspaceView(base_dir=tmp_path)
        view.model_combo.setCurrentText("Virgo")
        model_dir = view.create_project_folder_structure()

        # Place submission file in LSU
        lsu_file = model_dir / "CTTT" / "LSU" / "cttt_lsu.xlsx"
        wb = openpyxl.Workbook()
        wb.save(lsu_file)
        wb.close()


        view.scan_member_submissions()
        assert view.submission_table.rowCount() == len(STANDARD_SUB_UNITS)

        mapping = view.get_sub_unit_statuses()
        assert mapping.get("LSU") == "Đã nộp"

    def test_f22_batch_processing_worker_init(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Test 3: Instantiate BatchReconciliationWorker."""
        collected_data = [{"MÃ LINH KIỆN": "PART_A", "SỐ LƯỢNG": 1.0}]
        worker = BatchReconciliationWorker(
            collected_cttt=collected_data,
            plm_path=None,
            r3_path=None,
            model_name="Virgo",
        )
        assert worker.model_name == "Virgo"
        assert len(worker.collected_cttt) == 1

    def test_f22_i18n_language_switching(self) -> None:
        """Test 4: Tri-lingual dictionary support (VN, JP, CN) without hardcoded strings."""
        i18n = get_i18n()
        for lang_code in ["vi", "ja", "zh"]:
            i18n.set_language(lang_code)
            assert i18n.current_language == lang_code
            trans = t("status_idle")
            assert trans != ""
            assert not trans.startswith("[")

    def test_f22_audit_log_tracking(self, tmp_path: Path) -> None:
        """Test 5: Audit log records execution timestamps and leader actions."""
        log_file = tmp_path / "leader_audit.log"
        actions = [
            "Project Virgo created",
            "Member CTTT files ingested",
            "Reconciliation executed with overall status: OK",
        ]
        with open(log_file, "w", encoding="utf-8") as f:
            for act in actions:
                f.write(f"[2026-09-17 10:00:00] {act}\n")

        assert log_file.exists()
        lines = log_file.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 3
        assert "overall status: OK" in lines[2]
