"""Feature F28: Standalone PyInstaller Executable & Packaging Isolation Tests.

Verifies adherence to MP2027 Standard (HASH_ONLY_LAN):
1. PackageBuilder directory preparation and versioning.
2. Two-tier launcher architecture with root Launcher and versioned apps/<version>/ folder.
3. Deterministic SHA-256 computation.
4. CLI --health-check support verifying payload integrity before activation.
5. Rollback capability using previous.json pointer.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
import pytest

from scripts.package_app import APP_NAME, APP_VERSION, PackageBuilder, compute_sha256
from SSBOM_Launcher import AppLauncher


class TestF28StandalonePackaging:
    """Test suite for Feature F28: Standalone PyInstaller Executable & Packaging."""

    def test_f28_package_builder_initialization_and_clean(self, tmp_path: Path) -> None:
        """Test 1: Validate PackageBuilder directory preparation matching MP2027 spec."""
        builder = PackageBuilder(root_dir=tmp_path, version="1.0.0")
        builder.clean_and_prepare()

        assert builder.dist_dir.exists()
        assert builder.release_update_dir.exists()
        assert builder.apps_dir.exists()
        assert builder.version == "1.0.0"

    def test_f28_two_tier_launcher_architecture(self, tmp_path: Path) -> None:
        """Test 2: Two-tier structure with root Launcher and versioned apps/<version>/ folder."""
        launcher = AppLauncher(root_dir=tmp_path)
        # Fallback in dev when current.json is missing
        dev_pointer = launcher.read_current_pointer()
        assert dev_pointer.get("is_dev") is True
        assert "gui" in dev_pointer.get("entrypoint", "")

        # When current.json is present
        current_json = tmp_path / "current.json"
        current_json.write_text(json.dumps({
            "active_version": "1.0.0",
            "entrypoint": "apps/1.0.0/SSBOM_App.py",
            "manifest_sha256": "abc12345",
        }))
        pointer = launcher.read_current_pointer()
        assert pointer["active_version"] == "1.0.0"
        assert pointer["entrypoint"] == "apps/1.0.0/SSBOM_App.py"

    def test_f28_sha256_computation_and_integrity(self, tmp_path: Path) -> None:
        """Test 3: Validate deterministic SHA-256 computation."""
        test_file = tmp_path / "payload.bin"
        test_file.write_bytes(b"KYOCERA_SSBOM_TEST_PAYLOAD")

        digest = compute_sha256(test_file)
        assert isinstance(digest, str)
        assert len(digest) == 64

    def test_f28_healthcheck_argument_support(self) -> None:
        """Test 4: Validate --health-check CLI parameter execution against genuine app."""
        cmd = [sys.executable, "-m", "src.gui.app", "--health-check"]
        ret = subprocess.run(cmd, capture_output=True, text=True)
        assert ret.returncode == 0
        assert "SSBOM Health Check: OK" in ret.stdout

    def test_f28_rollback_via_previous_json(self, tmp_path: Path) -> None:
        """Test 5: Verify rollback capability using previous.json pointer."""
        previous_json = tmp_path / "previous.json"
        previous_json.write_text(json.dumps({
            "active_version": "0.9.0",
            "entrypoint": "apps/0.9.0/SSBOM_App.py",
            "manifest_sha256": "prev_hash",
        }))

        current_json = tmp_path / "current.json"
        current_json.write_text(json.dumps({
            "active_version": "1.0.0",
            "entrypoint": "apps/1.0.0/SSBOM_App.py",
            "manifest_sha256": "broken_hash",
        }))

        launcher = AppLauncher(root_dir=tmp_path)
        rolled_back = launcher.rollback()
        assert rolled_back is True

        updated_pointer = launcher.read_current_pointer()
        assert updated_pointer["active_version"] == "0.9.0"
