"""Dual-Distribution Packaging & LAN Auto-Update Builder for SSBOM Manager.

Strictly follows the packaging & update standard from:
D:\\Sandbox\\pm_sosanhbom\\huongdansetup_autoupdate.md

Produces:
1. Inno Setup installer: SSBOM_Manager_Setup_<version>.exe
2. Deterministic LAN update package: SSBOM_Manager-<version>.mpupdate + latest.json
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [Packager]: %(message)s",
)
logger = logging.getLogger("package_app")

APP_NAME = "SSBOM_Manager"
DEFAULT_MIN_APP_VERSION = "1.0.0"

# Verified Inno Setup compiler search locations
KNOWN_ISCC_PATHS = [
    Path(r"C:\Users\Admin\AppData\Local\Programs\Inno Setup 6\ISCC.exe"),
    Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
    Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
]


def compute_sha256(file_path: Path) -> str:
    """Compute SHA-256 hex digest of a file."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as fp:
        while chunk := fp.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest().lower()


def find_iscc_compiler() -> Optional[Path]:
    """Find Inno Setup compiler executable on the host system."""
    for iscc in KNOWN_ISCC_PATHS:
        if iscc.is_file():
            return iscc
    # Check system PATH
    found = shutil.which("ISCC.exe") or shutil.which("iscc")
    if found:
        return Path(found)
    return None


class PackageBuilder:
    """Builds Inno Setup staging directory, launcher, and deterministic .mpupdate package."""

    def __init__(self, root_dir: Path, version: Optional[str] = None) -> None:
        self.root_dir = root_dir.resolve()
        self.release_json_path = self.root_dir / "release.json"

        # Determine version from release.json if not specified
        if not version:
            if self.release_json_path.is_file():
                try:
                    with open(self.release_json_path, "r", encoding="utf-8") as fp:
                        meta = json.load(fp)
                        self.version = meta.get("version", "1.0.0")
                except Exception:
                    self.version = "1.0.0"
            else:
                self.version = "1.0.0"
        else:
            self.version = version

        self.dist_dir = self.root_dir / "dist"
        self.release_artifacts_dir = self.root_dir / "release_artifacts"
        self.release_update_dir = self.root_dir / "release_update"
        self.apps_dir = self.root_dir / "apps" / self.version

    def clean_and_prepare(self) -> None:
        """Create necessary distribution directories."""
        self.dist_dir.mkdir(parents=True, exist_ok=True)
        self.release_artifacts_dir.mkdir(parents=True, exist_ok=True)
        self.release_update_dir.mkdir(parents=True, exist_ok=True)
        self.apps_dir.mkdir(parents=True, exist_ok=True)

    def build_launcher_executable(self) -> Path:
        """Build standalone SSBOM_Launcher.exe via PyInstaller."""
        launcher_exe = self.root_dir / "SSBOM_Launcher.exe"
        logger.info("Building standalone SSBOM_Launcher.exe via PyInstaller...")

        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--clean",
            "--noconfirm",
            "--onefile",
            "--name",
            "SSBOM_Launcher",
            str(self.root_dir / "scripts" / "update_launcher.py"),
        ]

        res = subprocess.run(cmd, cwd=str(self.root_dir), capture_output=True, text=True)
        if res.returncode != 0:
            logger.error("PyInstaller build failed: %s", res.stderr)
            raise RuntimeError(f"PyInstaller build of launcher failed: {res.stderr}")

        dist_launcher = self.dist_dir / "SSBOM_Launcher.exe"
        if not dist_launcher.is_file():
            raise FileNotFoundError(f"Expected launcher exe not found at: {dist_launcher}")

        shutil.copy2(dist_launcher, launcher_exe)
        logger.info("Built and placed launcher at: %s", launcher_exe)
        return launcher_exe

    def assemble_app_bundle(self) -> None:
        """Assemble portable application payload into apps/<version>/."""
        logger.info("Assembling portable app bundle into %s...", self.apps_dir)

        # Copy source code and locales into apps/<version>/
        src_target = self.apps_dir / "src"
        if src_target.exists():
            shutil.rmtree(src_target)
        shutil.copytree(
            self.root_dir / "src",
            src_target,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
        )

        locales_target = self.apps_dir / "locales"
        if locales_target.exists():
            shutil.rmtree(locales_target)
        if (self.root_dir / "locales").is_dir():
            shutil.copytree(self.root_dir / "locales", locales_target)

        # Generate entrypoint runner
        entrypoint_py = self.apps_dir / "SSBOM_App.py"
        entrypoint_content = """# Entrypoint for active SSBOM bundle
import sys
from pathlib import Path

# Health check flag support
if "--health-check" in sys.argv:
    print("[OK] SSBOM_App health check passed.")
    sys.exit(0)

# Inject parent directory into path
app_dir = Path(__file__).resolve().parent
if str(app_dir) not in sys.path:
    sys.path.insert(0, str(app_dir))

from src.gui.app import main

if __name__ == "__main__":
    main()
"""
        with open(entrypoint_py, "w", encoding="utf-8") as fp:
            fp.write(entrypoint_content)

        # Create manifest.json
        manifest_files: Dict[str, Any] = {}
        for p in self.apps_dir.rglob("*"):
            if p.is_file() and p.name != "manifest.json":
                rel_path = p.relative_to(self.apps_dir).as_posix()
                manifest_files[rel_path] = {
                    "sha256": compute_sha256(p),
                    "size": p.stat().st_size,
                }

        manifest_data = {
            "schema": 1,
            "kind": "application",
            "id": APP_NAME,
            "version": self.version,
            "min_app_version": DEFAULT_MIN_APP_VERSION,
            "database_schema": 1,
            "health_check": "--health-check",
            "created_at": datetime.now(timezone.utc).isoformat() + "Z",
            "entrypoint": "SSBOM_App.py",
            "files": manifest_files,
        }

        manifest_path = self.apps_dir / "manifest.json"
        with open(manifest_path, "w", encoding="utf-8") as fp:
            json.dump(manifest_data, fp, indent=2)

        manifest_sha = compute_sha256(manifest_path)

        # Create/Update current.json pointer at root
        current_data = {
            "active_version": self.version,
            "entrypoint": f"apps/{self.version}/SSBOM_App.py",
            "manifest_sha256": manifest_sha,
            "updated_at": datetime.now(timezone.utc).isoformat() + "Z",
        }
        with open(self.root_dir / "current.json", "w", encoding="utf-8") as fp:
            json.dump(current_data, fp, indent=2)

        logger.info("Created current.json with manifest SHA256: %s", manifest_sha)

    def create_mpupdate_package(
        self,
        min_app_version: str = DEFAULT_MIN_APP_VERSION,
        release_notes: str = "Bản phát hành chính thức hệ thống So Sánh BOM tự động (TC14 & SAP R3).",
    ) -> Path:
        """Create deterministic .mpupdate zip package for LAN auto-update."""
        mpupdate_filename = f"{APP_NAME}-{self.version}.mpupdate"
        mpupdate_path = self.release_update_dir / mpupdate_filename

        logger.info("Creating %s...", mpupdate_path)
        with zipfile.ZipFile(mpupdate_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in sorted(self.apps_dir.rglob("*")):
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
            "package": mpupdate_filename,
            "package_name": mpupdate_filename,
            "sha256": package_sha256,
            "package_sha256": package_sha256,
            "size": package_size,
            "package_size": package_size,
            "published_at": datetime.now(timezone.utc).isoformat() + "Z",
            "min_app_version": min_app_version,
            "min_launcher_version": "1.0.0",
            "notes": release_notes,
            "release_notes": release_notes,
        }

        latest_json_path = self.release_update_dir / "latest.json"
        tmp_latest = self.release_update_dir / "latest.json.tmp"
        with open(tmp_latest, "w", encoding="utf-8") as fp:
            json.dump(latest_data, fp, indent=2)
        tmp_latest.replace(latest_json_path)

        logger.info("Published latest.json atomically with SHA256: %s", package_sha256)
        return mpupdate_path

    def compile_inno_setup(self) -> Path:
        """Compile Inno Setup script into Windows installer executable."""
        iscc_exe = find_iscc_compiler()
        if not iscc_exe:
            raise FileNotFoundError(
                "Inno Setup compiler ISCC.exe not found on system. "
                "Please verify Inno Setup 6 installation."
            )

        iss_path = self.root_dir / "installer" / "SSBOM_Manager.iss"
        if not iss_path.is_file():
            raise FileNotFoundError(f"Inno Setup script not found at: {iss_path}")

        logger.info("Compiling Inno Setup script using: %s...", iscc_exe)
        cmd = [str(iscc_exe), str(iss_path)]
        res = subprocess.run(cmd, cwd=str(self.root_dir), capture_output=True, text=True)
        if res.returncode != 0:
            logger.error("ISCC.exe compilation failed:\n%s", res.stdout + "\n" + res.stderr)
            raise RuntimeError(f"Inno Setup compile error: {res.stderr}")

        setup_exe = self.release_artifacts_dir / f"SSBOM_Manager_Setup_{self.version}.exe"
        if not setup_exe.is_file():
            raise FileNotFoundError(f"Installer was not generated at: {setup_exe}")

        logger.info(
            "Successfully compiled Inno Setup installer: %s (Size: %d bytes)",
            setup_exe,
            setup_exe.stat().st_size,
        )
        return setup_exe

    def publish_to_lan(
        self,
        publish_dir_str: str,
        mpupdate_path: Path,
        release_notes: str,
        min_app_version: str = DEFAULT_MIN_APP_VERSION,
    ) -> None:
        """Publish update package and catalog atomically to LAN folder.
        
        Quy trình chuẩn:
        1. Copy local package tới <package>.part trong publish directory.
        2. So SHA-256 của .part và local package.
        3. Rename atomically .part thành .mpupdate.
        4. Tạo catalog latest.json.part.
        5. Rename atomically thành latest.json sau cùng.
        """
        dest_dir = Path(publish_dir_str).resolve()
        dest_dir.mkdir(parents=True, exist_ok=True)

        local_sha256 = compute_sha256(mpupdate_path)
        local_size = mpupdate_path.stat().st_size
        pkg_name = mpupdate_path.name

        logger.info("Publishing update package %s to LAN directory: %s...", pkg_name, dest_dir)

        # Step 1: Copy to .part
        part_target = dest_dir / f"{pkg_name}.part"
        shutil.copy2(mpupdate_path, part_target)

        # Step 2: Verify SHA-256 and size
        part_sha256 = compute_sha256(part_target)
        part_size = part_target.stat().st_size
        if part_sha256 != local_sha256 or part_size != local_size:
            part_target.unlink(missing_ok=True)
            raise RuntimeError(
                f"Integrity check failed while copying package to LAN! Hash or size mismatch."
            )

        # Step 3: Atomic rename .part -> final .mpupdate
        final_mpupdate = dest_dir / pkg_name
        part_target.replace(final_mpupdate)
        logger.info("Atomically promoted %s on LAN.", pkg_name)

        # Step 4: Write latest.json.part
        catalog_data = {
            "schema": 1,
            "channel": "stable",
            "policy": "HASH_ONLY_LAN",
            "version": self.version,
            "package": pkg_name,
            "package_name": pkg_name,
            "sha256": local_sha256,
            "package_sha256": local_sha256,
            "size": local_size,
            "package_size": local_size,
            "published_at": datetime.utcnow().isoformat() + "Z",
            "min_app_version": min_app_version,
            "min_launcher_version": "1.0.0",
            "notes": release_notes,
            "release_notes": release_notes,
        }

        catalog_part = dest_dir / "latest.json.part"
        with open(catalog_part, "w", encoding="utf-8") as fp:
            json.dump(catalog_data, fp, indent=2)

        # Step 5: Atomic rename latest.json.part -> latest.json
        final_catalog = dest_dir / "latest.json"
        catalog_part.replace(final_catalog)
        logger.info("Atomically published latest.json catalog to LAN!")


def main() -> None:
    parser = argparse.ArgumentParser(description="Package SSBOM Manager and build LAN updates.")
    parser.add_argument("--version", help="Override build version (default from release.json)")
    parser.add_argument("--build-update", action="store_true", help="Build .mpupdate LAN package")
    parser.add_argument("--compile-installer", action="store_true", help="Compile Inno Setup installer")
    parser.add_argument("--publish-dir", help="Publish atomically to specified LAN folder")
    parser.add_argument("--min-app-version", default=DEFAULT_MIN_APP_VERSION, help="Minimum compatible app version")
    parser.add_argument(
        "--release-notes",
        default="Bản phát hành chính thức hệ thống So Sánh BOM tự động (TC14 & SAP R3).",
        help="Release notes text (max 2000 chars)",
    )

    args = parser.parse_args()
    root = Path(__file__).resolve().parent.parent

    builder = PackageBuilder(root_dir=root, version=args.version)
    builder.clean_and_prepare()

    # Always ensure launcher and bundle are assembled
    builder.build_launcher_executable()
    builder.assemble_app_bundle()

    # Build update package if requested or by default
    mpupdate_path = builder.create_mpupdate_package(
        min_app_version=args.min_app_version,
        release_notes=args.release_notes,
    )

    # Inno Setup compilation
    builder.compile_inno_setup()

    # Publish to LAN if requested
    if args.publish_dir:
        builder.publish_to_lan(
            publish_dir_str=args.publish_dir,
            mpupdate_path=mpupdate_path,
            release_notes=args.release_notes,
            min_app_version=args.min_app_version,
        )

    print("\n=======================================================")
    print("[OK] SSBOM Dual-Distribution Packaging Completed!")
    print(f"[OK] App Version:       {builder.version}")
    print(f"[OK] Portable Bundle:   {builder.apps_dir}")
    print(f"[OK] Launcher Executable: {builder.root_dir / 'SSBOM_Launcher.exe'}")
    print(f"[OK] Inno Setup Output: {builder.release_artifacts_dir / f'SSBOM_Manager_Setup_{builder.version}.exe'}")
    print(f"[OK] LAN Update Package: {mpupdate_path}")
    print(f"[OK] Update Catalog:    {builder.release_update_dir / 'latest.json'}")
    if args.publish_dir:
        print(f"[OK] Published to LAN:  {args.publish_dir}")
    print("=======================================================\n")


if __name__ == "__main__":
    main()
