"""Application Update Staging, Verification, Health-Check, and Activation Manager.

Strictly follows: huongdansetup_autoupdate.md
- Fail-closed lifecycle:
  1. Validate package & manifest
  2. Extract safely into .staging/<version>
  3. Verify all SHA-256 hashes against manifest
  4. Promote staging to apps/<version>
  5. Run isolated --health-check
  6. Backup runtime databases to backups/before-<version>/
  7. Atomically save current.json -> previous.json
  8. Atomically activate new current.json
- Rollback support: swap previous.json <-> current.json
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from src.services.update_delivery import is_version_greater, parse_version_tuple
from src.services.update_security import (
    SafeArchiveExtractor,
    UpdateSecurityError,
    compute_file_sha256,
    is_safe_relative_path,
    verify_manifest_integrity,
)

logger = logging.getLogger("app_updates")


class AppUpdateError(Exception):
    """Raised when an error occurs during update installation or activation."""
    pass


class AppUpdateManager:
    """Manages the installation, staging, health-checking, activation, and rollback of updates."""

    def __init__(self, install_root: Path) -> None:
        self.install_root = Path(install_root).resolve()
        self.staging_dir = self.install_root / ".staging"
        self.apps_dir = self.install_root / "apps"
        self.backups_dir = self.install_root / "backups"
        self.current_json_path = self.install_root / "current.json"
        self.previous_json_path = self.install_root / "previous.json"
        self.release_json_path = self.install_root / "release.json"

    def get_current_state(self) -> Dict[str, Any]:
        """Read current.json pointer if exists, else return default state."""
        if self.current_json_path.is_file():
            try:
                with open(self.current_json_path, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    if isinstance(data, dict):
                        return data
            except Exception as exc:
                logger.error("Failed to read current.json: %s", exc)

        # Fallback to release.json if current.json not present
        if self.release_json_path.is_file():
            try:
                with open(self.release_json_path, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    if isinstance(data, dict):
                        return {
                            "active_version": data.get("version", "1.0.0"),
                            "entrypoint": "src/gui/app.py",
                            "manifest_sha256": "",
                        }
            except Exception:
                pass

        return {"active_version": "1.0.0", "entrypoint": "src/gui/app.py", "manifest_sha256": ""}

    def get_active_version(self) -> str:
        """Get string representation of the currently active version."""
        return str(self.get_current_state().get("active_version", "1.0.0"))

    def install_update_package(
        self,
        package_path: Path,
        allow_same_version: bool = False,
    ) -> Dict[str, Any]:
        """Execute full fail-closed update lifecycle.
        
        Args:
            package_path: Path to verified .mpupdate file.
            allow_same_version: Only True in testing / re-installation scenarios.
            
        Returns:
            Dictionary with update summary including activated version and manifest hash.
        """
        pkg = Path(package_path).resolve()
        if not pkg.is_file():
            raise FileNotFoundError(f"Update package not found: {pkg}")

        current_ver = self.get_active_version()

        # Step 1: Open archive & read manifest.json
        try:
            with zipfile.ZipFile(pkg, "r") as zf:
                if "manifest.json" not in zf.namelist():
                    raise AppUpdateError("Package is missing 'manifest.json' at root.")
                with zf.open("manifest.json") as mf:
                    manifest_data = json.loads(mf.read().decode("utf-8"))
        except Exception as exc:
            raise AppUpdateError(f"Cannot read manifest from update archive: {exc}") from exc

        target_version = manifest_data.get("version")
        if not target_version:
            raise AppUpdateError("Manifest does not specify target version.")

        if not allow_same_version and not is_version_greater(target_version, current_ver):
            raise AppUpdateError(
                f"Target update version ({target_version}) is not newer than current ({current_ver})."
            )

        # Validate minimum app version if declared
        min_app_ver = manifest_data.get("min_app_version")
        if min_app_ver and parse_version_tuple(current_ver) < parse_version_tuple(min_app_ver):
            raise AppUpdateError(
                f"Current app version {current_ver} is lower than required minimum {min_app_ver}."
            )

        # Validate entrypoint safety
        entrypoint = manifest_data.get("entrypoint", "SSBOM_App.py")
        if not is_safe_relative_path(entrypoint):
            raise AppUpdateError(f"Unsafe entrypoint specified in manifest: '{entrypoint}'")

        # Step 2: Extract safely into .staging/<target_version>
        self.staging_dir.mkdir(parents=True, exist_ok=True)
        staged_target = self.staging_dir / target_version
        if staged_target.exists():
            shutil.rmtree(staged_target, ignore_errors=True)

        logger.info("Extracting package to staging area: %s", staged_target)
        extractor = SafeArchiveExtractor()
        try:
            extractor.extract_all(pkg, staged_target)
        except UpdateSecurityError as exc:
            shutil.rmtree(staged_target, ignore_errors=True)
            raise AppUpdateError(f"Security violation during package extraction: {exc}") from exc

        # Step 3: Verify manifest integrity against extracted files
        is_intact, reason = verify_manifest_integrity(manifest_data, staged_target)
        if not is_intact:
            shutil.rmtree(staged_target, ignore_errors=True)
            raise AppUpdateError(f"Manifest integrity verification failed: {reason}")

        manifest_file_on_disk = staged_target / "manifest.json"
        manifest_sha256 = compute_file_sha256(manifest_file_on_disk)

        # Step 4: Promote staging to apps/<target_version>
        self.apps_dir.mkdir(parents=True, exist_ok=True)
        final_app_dir = self.apps_dir / target_version
        if final_app_dir.exists():
            shutil.rmtree(final_app_dir, ignore_errors=True)

        try:
            shutil.move(str(staged_target), str(final_app_dir))
        except Exception as exc:
            shutil.rmtree(staged_target, ignore_errors=True)
            raise AppUpdateError(f"Failed to promote staged files to apps directory: {exc}") from exc

        # Step 5: Execute isolated health check
        health_passed = self._run_health_check(final_app_dir, entrypoint)
        if not health_passed:
            shutil.rmtree(final_app_dir, ignore_errors=True)
            raise AppUpdateError(
                f"Health check failed for version {target_version}. Aborting update."
            )

        # Step 6: Backup runtime databases if present
        self._backup_runtime_data(target_version)

        # Step 7 & 8: Atomic Pointer Update
        self._activate_version_atomically(
            target_version=target_version,
            entrypoint=f"apps/{target_version}/{entrypoint}",
            manifest_sha256=manifest_sha256,
        )

        logger.info("Update to version %s successfully installed and activated!", target_version)
        return {
            "version": target_version,
            "manifest_sha256": manifest_sha256,
            "installed_at": datetime.now(timezone.utc).isoformat() + "Z",
            "entrypoint": f"apps/{target_version}/{entrypoint}",
        }

    def _run_health_check(self, app_dir: Path, entrypoint: str) -> bool:
        """Run health check against the newly deployed app directory in an isolated profile."""
        entry_file = app_dir / entrypoint
        if not entry_file.is_file():
            logger.error("Health check target entrypoint not found: %s", entry_file)
            return False

        with tempfile.TemporaryDirectory() as temp_localappdata:
            env = os.environ.copy()
            env["LOCALAPPDATA"] = temp_localappdata
            env["PYTHONPATH"] = str(app_dir)

            cmd = [sys.executable, str(entry_file), "--health-check"]
            logger.info("Running health check command: %s", " ".join(cmd))

            try:
                res = subprocess.run(
                    cmd,
                    cwd=str(app_dir),
                    env=env,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=60,
                )
                if res.returncode == 0:
                    logger.info("Health check passed successfully: %s", res.stdout.strip())
                    return True
                else:
                    logger.error(
                        "Health check returned non-zero exit code %d. Stderr: %s",
                        res.returncode,
                        res.stderr,
                    )
                    return False
            except subprocess.TimeoutExpired:
                logger.error("Health check timed out after 60 seconds.")
                return False
            except Exception as exc:
                logger.error("Error executing health check process: %s", exc)
                return False

    def _backup_runtime_data(self, target_version: str) -> None:
        """Backup database and runtime files to backups/before-<target_version>/."""
        backup_folder = self.backups_dir / f"before-{target_version}"
        backup_folder.mkdir(parents=True, exist_ok=True)

        inventory: Dict[str, Any] = {
            "created_at": datetime.now(timezone.utc).isoformat() + "Z",
            "target_version": target_version,
            "files": {},
        }

        # Backup any .db, .sqlite, .sqlite3 files in install root
        for db_file in self.install_root.glob("*.db"):
            dest = backup_folder / db_file.name
            shutil.copy2(db_file, dest)
            inventory["files"][db_file.name] = {
                "sha256": compute_file_sha256(dest),
                "size": dest.stat().st_size,
            }

        for db_file in self.install_root.glob("*.sqlite*"):
            dest = backup_folder / db_file.name
            shutil.copy2(db_file, dest)
            inventory["files"][db_file.name] = {
                "sha256": compute_file_sha256(dest),
                "size": dest.stat().st_size,
            }

        with open(backup_folder / "backup.json", "w", encoding="utf-8") as fp:
            json.dump(inventory, fp, indent=2)

    def _activate_version_atomically(
        self,
        target_version: str,
        entrypoint: str,
        manifest_sha256: str,
    ) -> None:
        """Atomically promote target_version as the new current.json."""
        # 1. If current.json exists, atomically back it up to previous.json
        if self.current_json_path.is_file():
            tmp_prev = self.previous_json_path.with_suffix(".tmp")
            shutil.copy2(self.current_json_path, tmp_prev)
            tmp_prev.replace(self.previous_json_path)

        # 2. Write new current.json via temporary file
        new_state = {
            "active_version": target_version,
            "entrypoint": entrypoint,
            "manifest_sha256": manifest_sha256,
            "updated_at": datetime.now(timezone.utc).isoformat() + "Z",
        }

        tmp_current = self.current_json_path.with_suffix(".tmp")
        with open(tmp_current, "w", encoding="utf-8") as fp:
            json.dump(new_state, fp, indent=2)

        tmp_current.replace(self.current_json_path)
        logger.info("Successfully activated version %s in current.json", target_version)

    def rollback_activation(self) -> bool:
        """Roll back active version to previous.json if available and valid."""
        if not self.previous_json_path.is_file():
            logger.warning("No previous.json found. Rollback is not possible.")
            return False

        try:
            with open(self.previous_json_path, "r", encoding="utf-8") as fp:
                prev_data = json.load(fp)
        except Exception as exc:
            logger.error("Failed to read previous.json during rollback: %s", exc)
            return False

        prev_version = prev_data.get("active_version")
        prev_entry = prev_data.get("entrypoint")
        if not prev_version or not prev_entry:
            logger.error("Invalid previous.json format: missing version or entrypoint.")
            return False

        entry_path = self.install_root / prev_entry
        if not entry_path.is_file():
            logger.error("Rollback target entrypoint does not exist on disk: %s", entry_path)
            return False

        # Atomic swap: read current, write current = prev, write previous = old current
        old_current_data = self.get_current_state()

        tmp_curr = self.current_json_path.with_suffix(".tmp")
        with open(tmp_curr, "w", encoding="utf-8") as fp:
            json.dump(prev_data, fp, indent=2)
        tmp_curr.replace(self.current_json_path)

        tmp_prev = self.previous_json_path.with_suffix(".tmp")
        with open(tmp_prev, "w", encoding="utf-8") as fp:
            json.dump(old_current_data, fp, indent=2)
        tmp_prev.replace(self.previous_json_path)

        logger.info("Successfully rolled back active version to: %s", prev_version)
        return True
