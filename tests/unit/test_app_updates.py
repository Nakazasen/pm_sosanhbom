"""Unit tests for Application Update Staging, Activation, and Rollback Service.

Verifies:
- Fail-closed update staging into .staging/<version>
- Manifest hash verification and promotion to apps/<version>
- Isolated health-check execution
- Database backup to backups/before-<version>/
- Atomic current.json pointer update and previous.json recording
- Rollback support via atomic pointer swap
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from src.services.app_updates import AppUpdateError, AppUpdateManager
from src.services.update_security import compute_file_sha256


def create_mock_mpupdate(
    dest_path: Path,
    version: str,
    entrypoint: str = "SSBOM_App.py",
    corrupt_file: bool = False,
) -> Path:
    """Helper creating a valid or corrupted mock .mpupdate archive with health check support."""
    entry_code = """import sys
if "--health-check" in sys.argv:
    print("Health check OK")
    sys.exit(0)
print("Normal execution")
"""
    manifest_files = {}

    with zipfile.ZipFile(dest_path, "w") as zf:
        zf.writestr(entrypoint, entry_code)
        entry_bytes = entry_code.encode("utf-8")
        manifest_files[entrypoint] = {
            "sha256": compute_file_sha256_bytes(entry_bytes),
            "size": len(entry_bytes),
        }

        # Extra file
        zf.writestr("version.txt", f"Version {version}")
        v_bytes = f"Version {version}".encode("utf-8")

        if corrupt_file:
            # Deliberately put bad hash in manifest
            manifest_files["version.txt"] = {
                "sha256": "0000000000000000000000000000000000000000000000000000000000000000",
                "size": len(v_bytes),
            }
        else:
            manifest_files["version.txt"] = {
                "sha256": compute_file_sha256_bytes(v_bytes),
                "size": len(v_bytes),
            }

        manifest = {
            "schema": 1,
            "id": "SSBOM_Manager",
            "version": version,
            "min_app_version": "1.0.0",
            "health_check": "--health-check",
            "entrypoint": entrypoint,
            "files": manifest_files,
        }
        zf.writestr("manifest.json", json.dumps(manifest, indent=2))

    return dest_path


def compute_file_sha256_bytes(data: bytes) -> str:
    import hashlib
    return hashlib.sha256(data).hexdigest().lower()


class TestAppUpdateManager:
    """Tests for AppUpdateManager fail-closed operations."""

    def test_get_current_state_default(self, tmp_path: Path) -> None:
        mgr = AppUpdateManager(install_root=tmp_path)
        state = mgr.get_current_state()
        assert state["active_version"] == "1.0.0"

    def test_full_update_lifecycle_and_activation(self, tmp_path: Path) -> None:
        mgr = AppUpdateManager(install_root=tmp_path)

        # Set initial current.json
        initial_current = {
            "active_version": "1.0.0",
            "entrypoint": "apps/1.0.0/SSBOM_App.py",
            "manifest_sha256": "dummy_sha_v1",
        }
        (tmp_path / "current.json").write_text(json.dumps(initial_current), encoding="utf-8")

        # Create dummy database to test runtime backup
        dummy_db = tmp_path / "production_data.sqlite3"
        dummy_db.write_bytes(b"SQLITE_DATA_BINARY")

        # Create valid v1.0.1 package
        pkg_path = tmp_path / "SSBOM_Manager-1.0.1.mpupdate"
        create_mock_mpupdate(pkg_path, version="1.0.1")

        # Install update
        result = mgr.install_update_package(pkg_path)

        assert result["version"] == "1.0.1"
        assert result["entrypoint"] == "apps/1.0.1/SSBOM_App.py"
        assert (tmp_path / "apps" / "1.0.1" / "SSBOM_App.py").is_file()

        # Verify current.json updated
        with open(tmp_path / "current.json", "r", encoding="utf-8") as fp:
            new_current = json.load(fp)
        assert new_current["active_version"] == "1.0.1"
        assert new_current["entrypoint"] == "apps/1.0.1/SSBOM_App.py"

        # Verify previous.json created with old state
        with open(tmp_path / "previous.json", "r", encoding="utf-8") as fp:
            prev = json.load(fp)
        assert prev["active_version"] == "1.0.0"

        # Verify backup was created
        backup_dir = tmp_path / "backups" / "before-1.0.1"
        assert backup_dir.is_dir()
        assert (backup_dir / "production_data.sqlite3").is_file()
        assert (backup_dir / "backup.json").is_file()

    def test_update_rejected_if_older_version(self, tmp_path: Path) -> None:
        mgr = AppUpdateManager(install_root=tmp_path)
        current = {
            "active_version": "1.5.0",
            "entrypoint": "apps/1.5.0/SSBOM_App.py",
            "manifest_sha256": "sha",
        }
        (tmp_path / "current.json").write_text(json.dumps(current), encoding="utf-8")

        pkg_path = tmp_path / "SSBOM_Manager-1.2.0.mpupdate"
        create_mock_mpupdate(pkg_path, version="1.2.0")

        with pytest.raises(AppUpdateError) as excinfo:
            mgr.install_update_package(pkg_path)

        assert "not newer than current" in str(excinfo.value)
        # Ensure current.json was unchanged
        assert mgr.get_active_version() == "1.5.0"

    def test_update_aborts_on_corrupted_manifest(self, tmp_path: Path) -> None:
        mgr = AppUpdateManager(install_root=tmp_path)
        pkg_path = tmp_path / "SSBOM_Manager-1.0.2.mpupdate"
        create_mock_mpupdate(pkg_path, version="1.0.2", corrupt_file=True)

        with pytest.raises(AppUpdateError) as excinfo:
            mgr.install_update_package(pkg_path)

        assert "Manifest integrity verification failed" in str(excinfo.value)
        # Verify staging directory was cleaned up
        assert not (tmp_path / ".staging" / "1.0.2").exists()
        assert not (tmp_path / "apps" / "1.0.2").exists()

    def test_rollback_activation_swaps_pointers(self, tmp_path: Path) -> None:
        mgr = AppUpdateManager(install_root=tmp_path)

        # Create dummy apps/1.0.0 and apps/1.0.1 entrypoint files
        (tmp_path / "apps" / "1.0.0").mkdir(parents=True, exist_ok=True)
        (tmp_path / "apps" / "1.0.0" / "SSBOM_App.py").write_text("v1.0.0", encoding="utf-8")

        (tmp_path / "apps" / "1.0.1").mkdir(parents=True, exist_ok=True)
        (tmp_path / "apps" / "1.0.1" / "SSBOM_App.py").write_text("v1.0.1", encoding="utf-8")

        # Setup current = 1.0.1, previous = 1.0.0
        (tmp_path / "current.json").write_text(
            json.dumps(
                {
                    "active_version": "1.0.1",
                    "entrypoint": "apps/1.0.1/SSBOM_App.py",
                    "manifest_sha256": "sha_101",
                }
            ),
            encoding="utf-8",
        )
        (tmp_path / "previous.json").write_text(
            json.dumps(
                {
                    "active_version": "1.0.0",
                    "entrypoint": "apps/1.0.0/SSBOM_App.py",
                    "manifest_sha256": "sha_100",
                }
            ),
            encoding="utf-8",
        )

        assert mgr.get_active_version() == "1.0.1"

        # Execute rollback
        success = mgr.rollback_activation()
        assert success is True
        assert mgr.get_active_version() == "1.0.0"

        # Verify previous now holds 1.0.1
        with open(tmp_path / "previous.json", "r", encoding="utf-8") as fp:
            prev_after = json.load(fp)
        assert prev_after["active_version"] == "1.0.1"
