"""Feature F20: Machine Model Code Unification Isolation Tests.

Verifies:
1. Unifying 'ma1' (110*) and 'maT' (T10*) legacy split workbooks into a single engine.
2. Material-specific subfolder routing: base_dir / <material> / R3_<material>_<date>.xls.
3. Automatic archival of older export files into capnhat/old/ with timestamps.
4. Sequential batch processing across lists of materials.
5. Graceful omission of empty or whitespace material codes.
"""

import datetime
from pathlib import Path
from unittest.mock import MagicMock
import pytest

from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import ExportResult


class TestF20MachineCodeUnification:
    """Test suite for Feature F20: Machine Model Code Unification."""

    def test_f20_unify_ma1_and_mat_prefixes(self, tmp_path: Path, mock_sap_session):
        """Test 1: Unify mass-production ('110*') and prototype ('T10*') materials."""
        service = CS12Service(mock_sap_session)
        service.execute_cs12_and_export = MagicMock(
            side_effect=lambda p: ExportResult(success=True, material=p.material, file_path=p.destination_dir / p.expected_filename)
        )

        materials = ["110C103NL0", "T10K001NL0"]
        results = service.batch_download(
            materials=materials,
            valid_date=datetime.date(2026, 9, 17),
            base_destination_dir=tmp_path,
        )

        assert len(results) == 2
        assert results[0].material == "110C103NL0"
        assert results[1].material == "T10K001NL0"

    def test_f20_dedicated_subfolder_creation(self, tmp_path: Path, mock_sap_session):
        """Test 2: Each material is routed to its own dedicated subfolder."""
        service = CS12Service(mock_sap_session)
        service.execute_cs12_and_export = MagicMock(
            side_effect=lambda p: ExportResult(success=True, material=p.material, file_path=p.destination_dir / p.expected_filename)
        )

        service.batch_download(
            materials=["110C103NL0"],
            valid_date=datetime.date(2026, 9, 17),
            base_destination_dir=tmp_path,
        )

        expected_folder = tmp_path / "110C103NL0"
        assert expected_folder.is_dir()

    def test_f20_archive_existing_files(self, tmp_path: Path):
        """Test 3: Existing older files are moved to capnhat/old/ before overwriting."""
        mat_folder = tmp_path / "110C103NL0"
        mat_folder.mkdir(parents=True)
        old_file = mat_folder / "R3_110C103NL0_01_01_2026.xls"
        old_file.write_text("OLD VERSION")

        CS12Service._archive_existing_files(mat_folder, "110C103NL0")

        # Original old_file should be moved
        assert not old_file.exists()
        archive_dir = mat_folder / "capnhat" / "old"
        assert archive_dir.is_dir()
        archived_files = list(archive_dir.glob("R3_110C103NL0_*.xls"))
        assert len(archived_files) == 1

    def test_f20_batch_download_multiple_materials(self, tmp_path: Path, mock_sap_session):
        """Test 4: Batch download processes multiple machine models in sequence."""
        service = CS12Service(mock_sap_session)
        call_order = []

        def mock_exec(params):
            call_order.append(params.material)
            return ExportResult(success=True, material=params.material)

        service.execute_cs12_and_export = mock_exec

        materials = ["110C1", "110C2", "T10D1"]
        results = service.batch_download(
            materials=materials,
            valid_date=datetime.date(2026, 9, 17),
            base_destination_dir=tmp_path,
        )

        assert len(results) == 3
        assert call_order == ["110C1", "110C2", "T10D1"]

    def test_f20_skip_blank_material_entries(self, tmp_path: Path, mock_sap_session):
        """Test 5: Skip empty or whitespace entries without crashing."""
        service = CS12Service(mock_sap_session)
        service.execute_cs12_and_export = MagicMock(
            side_effect=lambda p: ExportResult(success=True, material=p.material)
        )

        materials = ["110C1", "", "   ", "T10D2"]
        results = service.batch_download(
            materials=materials,
            valid_date=datetime.date(2026, 9, 17),
            base_destination_dir=tmp_path,
        )

        # Only 2 valid materials processed
        assert len(results) == 2
        assert [r.material for r in results] == ["110C1", "T10D2"]
