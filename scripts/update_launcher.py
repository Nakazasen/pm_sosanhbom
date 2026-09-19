"""Stable Update Launcher for SSBOM Manager.

Strictly follows huongdansetup_autoupdate.md:
- Reads current.json from install root.
- Resolves pointer to active version.
- Validates manifest_sha256 of manifest.json to prevent corrupted/tampered launches.
- Supports --health-check flag for install bundle verification.
- Launches the active app entrypoint forwarding all arguments.
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
    format="%(asctime)s [%(levelname)s] [Launcher]: %(message)s",
)
logger = logging.getLogger("SSBOM_Launcher")


def compute_sha256(file_path: Path) -> str:
    """Compute SHA-256 of a file in 64KB chunks."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as fp:
        while chunk := fp.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest().lower()


class LauncherResolver:
    """Resolves active application pointer from current.json and verifies integrity."""

    def __init__(self, install_root: Path) -> None:
        self.install_root = Path(install_root).resolve()
        self.current_json_path = self.install_root / "current.json"

    def resolve(self) -> tuple[Path, str]:
        """Resolve entrypoint path and active version after cryptographic integrity check.
        
        Returns:
            Tuple of (entrypoint_path, active_version)
            
        Raises:
            RuntimeError: If pointer or manifest verification fails.
        """
        if not self.current_json_path.is_file():
            raise RuntimeError(f"Missing pointer file 'current.json' in {self.install_root}")

        try:
            with open(self.current_json_path, "r", encoding="utf-8") as fp:
                pointer_data: Dict[str, Any] = json.load(fp)
        except Exception as exc:
            raise RuntimeError(f"Could not parse current.json: {exc}") from exc

        active_version = pointer_data.get("active_version")
        entrypoint_rel = pointer_data.get("entrypoint")
        expected_manifest_sha = (pointer_data.get("manifest_sha256") or "").lower()

        if not active_version or not entrypoint_rel:
            raise RuntimeError("current.json missing 'active_version' or 'entrypoint'")

        # Locate manifest.json for active version
        app_dir = self.install_root / "apps" / active_version
        manifest_path = app_dir / "manifest.json"

        if not manifest_path.is_file():
            raise RuntimeError(f"Active version manifest not found: {manifest_path}")

        # Cryptographic verification against current.json
        if expected_manifest_sha:
            actual_manifest_sha = compute_sha256(manifest_path)
            if actual_manifest_sha != expected_manifest_sha:
                raise RuntimeError(
                    f"Integrity check failed: manifest SHA-256 mismatch!\n"
                    f"Expected: {expected_manifest_sha}\n"
                    f"Actual:   {actual_manifest_sha}"
                )

        entry_path = (self.install_root / entrypoint_rel).resolve()
        if not entry_path.is_file():
            # Fallback check inside app_dir if relative to app_dir
            alt_entry = (app_dir / entrypoint_rel).resolve()
            if alt_entry.is_file():
                entry_path = alt_entry
            else:
                raise RuntimeError(f"Entrypoint file does not exist: {entry_path}")

        return entry_path, active_version


def get_python_executable() -> str:
    """Find python interpreter, avoiding frozen executable recursion."""
    if not getattr(sys, "frozen", False):
        return sys.executable
    import shutil
    py_path = shutil.which("py") or shutil.which("python")
    if py_path:
        return py_path
    return "py"


def main() -> int:
    """Launcher main entrypoint."""
    # Determine install root: directory of launcher
    if getattr(sys, "frozen", False):
        install_root = Path(sys.executable).resolve().parent
    else:
        install_root = Path(__file__).resolve().parent
        # If placed in scripts/, root is parent
        if install_root.name == "scripts":
            install_root = install_root.parent

    # Allow overriding root via environment variable if running tests
    if "SSBOM_INSTALL_ROOT" in os.environ:
        install_root = Path(os.environ["SSBOM_INSTALL_ROOT"]).resolve()

    resolver = LauncherResolver(install_root)

    try:
        entrypoint_path, active_ver = resolver.resolve()
    except Exception as exc:
        logger.error("Failed to resolve active application: %s", exc)
        print(f"[LAUNCHER ERROR] {exc}", file=sys.stderr)
        return 1

    forward_args = sys.argv[1:]
    python_bin = get_python_executable()

    # Handle health-check flag
    if "--health-check" in forward_args:
        logger.info("Health check requested for active version %s.", active_ver)
        if entrypoint_path.suffix.lower() == ".py":
            cmd = [python_bin, str(entrypoint_path), "--health-check"]
        else:
            cmd = [str(entrypoint_path), "--health-check"]

        res = subprocess.run(cmd, cwd=str(entrypoint_path.parent))
        if res.returncode == 0:
            print(f"[OK] SSBOM_Launcher: active version {active_ver} health check passed.")
            return 0
        else:
            print(f"[FAIL] SSBOM_Launcher: health check returned code {res.returncode}", file=sys.stderr)
            return res.returncode

    # Normal launch
    logger.info("Launching SSBOM Manager version %s (%s)...", active_ver, entrypoint_path.name)
    if entrypoint_path.suffix.lower() == ".py":
        cmd = [python_bin, str(entrypoint_path)] + forward_args
    else:
        cmd = [str(entrypoint_path)] + forward_args

    proc = subprocess.run(cmd, cwd=str(entrypoint_path.parent))
    return proc.returncode


if __name__ == "__main__":
    sys.exit(main())
