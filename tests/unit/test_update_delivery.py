"""Unit tests for Update Delivery Service.

Verifies:
- Dotted / SemVer version comparison logic
- Multi-tier sources configuration hierarchy
- Catalog inspection and schema validation
- Atomic LAN package download via .part with SHA-256 verification
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from src.services.update_delivery import (
    UpdateCandidate,
    UpdateDeliveryError,
    UpdateDeliveryService,
    is_version_greater,
    parse_version_tuple,
)
from src.services.update_security import compute_file_sha256


class TestVersionComparison:
    """Tests for version parsing and comparison."""

    def test_parse_version_tuple(self) -> None:
        assert parse_version_tuple("1.0.0") == (1, 0, 0)
        assert parse_version_tuple("v0.1.7") == (0, 1, 7)
        assert parse_version_tuple("2.10.3") == (2, 10, 3)
        assert parse_version_tuple("invalid") == (0,)

    def test_is_version_greater(self) -> None:
        assert is_version_greater("1.0.1", "1.0.0") is True
        assert is_version_greater("2.0.0", "1.9.9") is True
        assert is_version_greater("1.1.0", "1.0.9") is True
        assert is_version_greater("1.0.0", "1.0.0") is False
        assert is_version_greater("0.9.9", "1.0.0") is False


class TestUpdateDeliveryService:
    """Tests for UpdateDeliveryService operations."""

    def test_load_sources_configuration_hierarchy(self, tmp_path: Path) -> None:
        # Default bundled
        default_file = tmp_path / "update_sources.default.json"
        default_file.write_text(
            json.dumps(
                {
                    "schema": 1,
                    "startup_check": True,
                    "sources": [{"type": "folder", "location": "C:/default", "enabled": True}],
                }
            ),
            encoding="utf-8",
        )

        service = UpdateDeliveryService(base_dir=tmp_path, current_version="1.0.0")
        cfg = service.load_sources_configuration()
        assert cfg["startup_check"] is True
        assert len(cfg["sources"]) == 1
        assert cfg["sources"][0]["location"] == "C:/default"

        # Local override
        override_file = tmp_path / "update_sources.json"
        override_file.write_text(
            json.dumps(
                {
                    "schema": 1,
                    "startup_check": False,
                    "sources": [{"type": "folder", "location": "D:/custom", "enabled": True}],
                }
            ),
            encoding="utf-8",
        )

        cfg_overridden = service.load_sources_configuration()
        assert cfg_overridden["startup_check"] is False
        assert cfg_overridden["sources"][0]["location"] == "D:/custom"

    def test_check_for_updates_discovers_newer_version(self, tmp_path: Path) -> None:
        lan_dir = tmp_path / "lan_share"
        lan_dir.mkdir(parents=True, exist_ok=True)

        pkg_file = lan_dir / "SSBOM_Manager-1.0.1.mpupdate"
        pkg_file.write_bytes(b"MOCK_ZIP_PACKAGE_CONTENT_V101")
        pkg_sha = compute_file_sha256(pkg_file)
        pkg_size = pkg_file.stat().st_size

        latest_json = lan_dir / "latest.json"
        catalog_data = {
            "schema": 1,
            "channel": "stable",
            "version": "1.0.1",
            "package": "SSBOM_Manager-1.0.1.mpupdate",
            "sha256": pkg_sha,
            "size": pkg_size,
            "notes": "- Fix critical UI bug\n- New feature",
        }
        latest_json.write_text(json.dumps(catalog_data), encoding="utf-8")

        # Configure service with LAN source
        sources_cfg = tmp_path / "update_sources.default.json"
        sources_cfg.write_text(
            json.dumps(
                {
                    "schema": 1,
                    "sources": [{"type": "folder", "location": str(lan_dir), "enabled": True}],
                }
            ),
            encoding="utf-8",
        )

        service = UpdateDeliveryService(base_dir=tmp_path, current_version="1.0.0")
        candidate = service.check_for_updates()

        assert candidate is not None
        assert candidate.version == "1.0.1"
        assert candidate.package_name == "SSBOM_Manager-1.0.1.mpupdate"
        assert candidate.package_sha256 == pkg_sha
        assert candidate.package_size == pkg_size
        assert "- Fix critical UI bug" in candidate.release_notes

    def test_check_for_updates_ignores_older_or_same_version(self, tmp_path: Path) -> None:
        lan_dir = tmp_path / "lan_share"
        lan_dir.mkdir(parents=True, exist_ok=True)

        pkg_file = lan_dir / "SSBOM_Manager-1.0.0.mpupdate"
        pkg_file.write_bytes(b"MOCK_CONTENT")
        pkg_sha = compute_file_sha256(pkg_file)

        latest_json = lan_dir / "latest.json"
        catalog_data = {
            "schema": 1,
            "version": "1.0.0",
            "package": "SSBOM_Manager-1.0.0.mpupdate",
            "sha256": pkg_sha,
            "size": pkg_file.stat().st_size,
            "notes": "Current version",
        }
        latest_json.write_text(json.dumps(catalog_data), encoding="utf-8")

        sources_cfg = tmp_path / "update_sources.default.json"
        sources_cfg.write_text(
            json.dumps(
                {
                    "schema": 1,
                    "sources": [{"type": "folder", "location": str(lan_dir), "enabled": True}],
                }
            ),
            encoding="utf-8",
        )

        service = UpdateDeliveryService(base_dir=tmp_path, current_version="1.0.0")
        candidate = service.check_for_updates()
        assert candidate is None

    def test_download_package_atomic_verified(self, tmp_path: Path) -> None:
        lan_dir = tmp_path / "lan_share"
        lan_dir.mkdir(parents=True, exist_ok=True)

        pkg_file = lan_dir / "SSBOM_Manager-1.0.1.mpupdate"
        content = b"GENUINE_UPDATE_BINARY_DATA"
        pkg_file.write_bytes(content)
        pkg_sha = compute_file_sha256(pkg_file)
        pkg_size = len(content)

        candidate = UpdateCandidate(
            version="1.0.1",
            package_name="SSBOM_Manager-1.0.1.mpupdate",
            package_sha256=pkg_sha,
            package_size=pkg_size,
            release_notes="Update",
            source_location=str(lan_dir),
        )

        service = UpdateDeliveryService(base_dir=tmp_path, current_version="1.0.0")
        downloaded = service.download_package(candidate)

        assert downloaded.is_file()
        assert downloaded.name == "SSBOM_Manager-1.0.1.mpupdate"
        assert compute_file_sha256(downloaded) == pkg_sha
        assert not (service.cache_dir / f"{candidate.package_name}.part").exists()

    def test_download_package_detects_hash_tamper(self, tmp_path: Path) -> None:
        lan_dir = tmp_path / "lan_share"
        lan_dir.mkdir(parents=True, exist_ok=True)

        pkg_file = lan_dir / "SSBOM_Manager-1.0.1.mpupdate"
        pkg_file.write_bytes(b"TAMPERED_DATA")

        # Fake expected sha256
        candidate = UpdateCandidate(
            version="1.0.1",
            package_name="SSBOM_Manager-1.0.1.mpupdate",
            package_sha256="ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff",
            package_size=len(b"TAMPERED_DATA"),
            release_notes="Tampered",
            source_location=str(lan_dir),
        )

        service = UpdateDeliveryService(base_dir=tmp_path, current_version="1.0.0")
        with pytest.raises(UpdateDeliveryError) as excinfo:
            service.download_package(candidate)

        assert "SHA-256 mismatch" in str(excinfo.value)
        # Verify temporary .part file was cleaned up
        assert not (service.cache_dir / f"{candidate.package_name}.part").exists()
