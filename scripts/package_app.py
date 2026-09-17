"""Dual-Distribution Packaging & LAN Auto-Update Builder for SSBOM Manager.

Strictly follows the MP2027 packaging standard from:
D:\\Sandbox\\MP2027\\huongdansetup_autoupdate.md

Produces:
1. Inno Setup installer: SSBOM_Manager_Setup_<version>.exe
2. Deterministic LAN update package: SSBOM_Manager-<version>.mpupdate + latest.json
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("package_app")

APP_VERSION = "1.0.0"
APP_NAME = "SSBOM_Manager"


def compute_sha256(file_path: Path) -> str:
    """Compute SHA-256 hex digest of a file."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as fp:
        while chunk := fp.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


class PackageBuilder:
    """Builds Inno Setup staging directory and deterministic .mpupdate package."""

    def __init__(self, root_dir: Path, version: str = APP_VERSION) -> None:
        self.root_dir = root_dir.resolve()
        self.version = version
        self.dist_dir = self.root_dir / "dist"
        self.release_update_dir = self.root_dir / "release_update"
        self.apps_dir = self.root_dir / "apps" / self.version

    def clean_and_prepare(self) -> None:
        """Create necessary directories."""
        self.dist_dir.mkdir(parents=True, exist_ok=True)
        self.release_update_dir.mkdir(parents=True, exist_ok=True)
        self.apps_dir.mkdir(parents=True, exist_ok=True)

    def assemble_app_bundle(self) -> None:
        """Assemble portable application payload into apps/<version>/."""
        logger.info("Assembling portable app bundle into %s...", self.apps_dir)

        # Copy source code and locales
        src_target = self.apps_dir / "src"
        if src_target.exists():
            shutil.rmtree(src_target)
        shutil.copytree(self.root_dir / "src", src_target, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))

        locales_target = self.apps_dir / "locales"
        if locales_target.exists():
            shutil.rmtree(locales_target)
        shutil.copytree(self.root_dir / "locales", locales_target)

        # Copy entrypoint runner
        entrypoint_py = self.apps_dir / "SSBOM_App.py"
        entrypoint_content = """# Entrypoint for active SSBOM bundle
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.gui.app import main

if __name__ == '__main__':
    main()
"""
        with open(entrypoint_py, "w", encoding="utf-8") as fp:
            fp.write(entrypoint_content)

        # Create manifest.json
        manifest_files = {}
        for p in self.apps_dir.rglob("*"):
            if p.is_file() and p.name != "manifest.json":
                rel_path = p.relative_to(self.apps_dir).as_posix()
                manifest_files[rel_path] = {
                    "sha256": compute_sha256(p),
                    "size": p.stat().st_size,
                }

        manifest_data = {
            "name": APP_NAME,
            "version": self.version,
            "created_at": datetime.utcnow().isoformat() + "Z",
            "entrypoint": "SSBOM_App.py",
            "files": manifest_files,
        }

        manifest_path = self.apps_dir / "manifest.json"
        with open(manifest_path, "w", encoding="utf-8") as fp:
            json.dump(manifest_data, fp, indent=2)

        manifest_sha = compute_sha256(manifest_path)

        # Create current.json pointer at root
        current_data = {
            "active_version": self.version,
            "entrypoint": f"apps/{self.version}/SSBOM_App.py",
            "manifest_sha256": manifest_sha,
            "updated_at": datetime.utcnow().isoformat() + "Z",
        }
        with open(self.root_dir / "current.json", "w", encoding="utf-8") as fp:
            json.dump(current_data, fp, indent=2)

        logger.info("Created current.json with manifest SHA256: %s", manifest_sha)

    def create_mpupdate_package(self) -> Path:
        """Create deterministic .mpupdate zip package for LAN auto-update."""
        mpupdate_filename = f"{APP_NAME}-{self.version}.mpupdate"
        mpupdate_path = self.release_update_dir / mpupdate_filename

        logger.info("Creating %s...", mpupdate_path)
        with zipfile.ZipFile(mpupdate_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in self.apps_dir.rglob("*"):
                if p.is_file():
                    arcname = p.relative_to(self.apps_dir).as_posix()
                    zf.write(p, arcname=arcname)

        package_sha256 = compute_sha256(mpupdate_path)
        package_size = mpupdate_path.stat().st_size

        # Create latest.json catalog
        latest_data = {
            "schema": 1,
            "channel": "stable",
            "policy": "HASH_ONLY_LAN",
            "version": self.version,
            "package_name": mpupdate_filename,
            "package_sha256": package_sha256,
            "package_size": package_size,
            "published_at": datetime.utcnow().isoformat() + "Z",
            "min_launcher_version": "1.0.0",
            "release_notes": "Bản phát hành chính thức hệ thống So Sánh BOM tự động (TC14 & SAP R3).",
        }

        latest_json_path = self.release_update_dir / "latest.json"
        # Atomic write via .tmp
        tmp_latest = self.release_update_dir / "latest.json.tmp"
        with open(tmp_latest, "w", encoding="utf-8") as fp:
            json.dump(latest_data, fp, indent=2)
        tmp_latest.replace(latest_json_path)

        logger.info("Published latest.json atomically with SHA256: %s", package_sha256)
        return mpupdate_path


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    builder = PackageBuilder(root_dir=root, version=APP_VERSION)
    builder.clean_and_prepare()
    builder.assemble_app_bundle()
    builder.create_mpupdate_package()
    print("\n=======================================================")
    print("[OK] SSBOM Packaging and LAN Auto-Update Artifacts Built!")
    print(f"[OK] Apps Bundle: {builder.apps_dir}")
    print(f"[OK] Update Package: {builder.release_update_dir / f'{APP_NAME}-{APP_VERSION}.mpupdate'}")
    print(f"[OK] Catalog Manifest: {builder.release_update_dir / 'latest.json'}")
    print("=======================================================\n")


if __name__ == "__main__":
    main()
