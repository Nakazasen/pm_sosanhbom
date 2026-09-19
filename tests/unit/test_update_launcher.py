"""Unit tests for Update Launcher Pointer Resolution and Integrity Verification.

Verifies:
- Launcher resolves active application entrypoint from current.json
- Cryptographic check matches manifest_sha256
- Launcher rejects tampered manifest.json
- Clean handling of missing pointers
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from scripts.update_launcher import LauncherResolver, compute_sha256


class TestLauncherResolver:
    """Tests for LauncherResolver."""

    def test_resolver_success(self, tmp_path: Path) -> None:
        # Create apps/1.0.0 directory with manifest.json and entrypoint
        app_dir = tmp_path / "apps" / "1.0.0"
        app_dir.mkdir(parents=True, exist_ok=True)

        entry_file = app_dir / "SSBOM_App.py"
        entry_file.write_text("print('App running')", encoding="utf-8")

        manifest_file = app_dir / "manifest.json"
        manifest_data = {
            "version": "1.0.0",
            "entrypoint": "SSBOM_App.py",
            "files": {},
        }
        manifest_file.write_text(json.dumps(manifest_data), encoding="utf-8")
        manifest_sha = compute_sha256(manifest_file)

        # Create current.json
        current_data = {
            "active_version": "1.0.0",
            "entrypoint": "apps/1.0.0/SSBOM_App.py",
            "manifest_sha256": manifest_sha,
        }
        (tmp_path / "current.json").write_text(json.dumps(current_data), encoding="utf-8")

        resolver = LauncherResolver(install_root=tmp_path)
        resolved_entry, active_ver = resolver.resolve()

        assert active_ver == "1.0.0"
        assert resolved_entry == entry_file

    def test_resolver_rejects_tampered_manifest(self, tmp_path: Path) -> None:
        app_dir = tmp_path / "apps" / "1.0.0"
        app_dir.mkdir(parents=True, exist_ok=True)

        entry_file = app_dir / "SSBOM_App.py"
        entry_file.write_text("print('App running')", encoding="utf-8")

        manifest_file = app_dir / "manifest.json"
        manifest_file.write_text(json.dumps({"version": "1.0.0", "entrypoint": "SSBOM_App.py"}), encoding="utf-8")

        # Set wrong expected manifest SHA in current.json
        current_data = {
            "active_version": "1.0.0",
            "entrypoint": "apps/1.0.0/SSBOM_App.py",
            "manifest_sha256": "bad0000000000000000000000000000000000000000000000000000000000bad",
        }
        (tmp_path / "current.json").write_text(json.dumps(current_data), encoding="utf-8")

        resolver = LauncherResolver(install_root=tmp_path)
        with pytest.raises(RuntimeError) as excinfo:
            resolver.resolve()

        assert "manifest SHA-256 mismatch" in str(excinfo.value)

    def test_resolver_missing_pointer(self, tmp_path: Path) -> None:
        resolver = LauncherResolver(install_root=tmp_path)
        with pytest.raises(RuntimeError) as excinfo:
            resolver.resolve()

        assert "Missing pointer file 'current.json'" in str(excinfo.value)
