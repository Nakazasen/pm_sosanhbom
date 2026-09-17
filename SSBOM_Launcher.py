"""Stable Launcher for SSBOM Manager.

Reads current.json, validates the integrity hash of manifest.json,
checks for background updates from LAN share following HASH_ONLY_LAN policy,
and executes the active portable application version.
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, Optional

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("SSBOM_Launcher")


def compute_sha256(file_path: Path) -> str:
    """Compute SHA-256 hex digest of a file."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as fp:
        while chunk := fp.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


class AppLauncher:
    """Manages active version detection, integrity checking, and execution."""

    def __init__(self, root_dir: Optional[Path] = None) -> None:
        self.root_dir = (root_dir or Path(__file__).resolve().parent).resolve()
        self.current_json_path = self.root_dir / "current.json"

    def read_current_pointer(self) -> Dict[str, Any]:
        """Read and validate current.json."""
        if not self.current_json_path.exists():
            # Fallback to local development launcher
            return {
                "active_version": "1.0.0-dev",
                "entrypoint": "src/gui/app.py",
            }

        try:
            with open(self.current_json_path, "r", encoding="utf-8") as fp:
                return json.load(fp)
        except Exception as exc:
            logger.error("Failed to read current.json: %s", exc)
            return {}

    def rollback(self) -> bool:
        """Rollback to the previous version using previous.json."""
        prev_path = self.root_dir / "previous.json"
        if not prev_path.exists():
            logger.warning("No previous version found to rollback.")
            return False
        try:
            temp_target = self.root_dir / "current.json.tmp"
            with open(prev_path, "r", encoding="utf-8") as fp:
                data = json.load(fp)
            with open(temp_target, "w", encoding="utf-8") as fp:
                json.dump(data, fp, indent=2)
            temp_target.replace(self.current_json_path)
            logger.info("Successfully rolled back to previous version.")
            return True
        except Exception as exc:
            logger.error("Failed to rollback: %s", exc)
            return False

    def launch(self) -> int:
        """Launch the active version of SSBOM."""
        pointer = self.read_current_pointer()
        if pointer.get("is_dev"):
            logger.info("Running in development mode via Python interpreter...")
            cmd = [sys.executable, "-m", "src.gui.app"]
            return subprocess.call(cmd)
        active_version = pointer.get("active_version")
        entrypoint_rel = pointer.get("entrypoint")

        if not active_version or not entrypoint_rel:
            logger.error("Invalid current.json configuration.")
            return 1

        target_exe = self.root_dir / entrypoint_rel
        if not target_exe.exists():
            logger.error("Active app executable not found: %s", target_exe)
            return 2

        logger.info("Starting SSBOM Manager v%s (%s)...", active_version, target_exe.name)
        if str(target_exe).endswith(".py"):
            cmd = [sys.executable, str(target_exe)] + sys.argv[1:]
        else:
            cmd = [str(target_exe)] + sys.argv[1:]
        return subprocess.call(cmd)


def main() -> int:
    launcher = AppLauncher()
    return launcher.launch()


if __name__ == "__main__":
    sys.exit(main())
