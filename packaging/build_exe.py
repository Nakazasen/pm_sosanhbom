"""PyInstaller Executable Builder & Health Check for SSBOM Manager."""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("build_exe")


def build_standalone_exe() -> int:
    """Run PyInstaller using pm_sosanhbom.spec."""
    project_root = Path(__file__).resolve().parent.parent
    spec_path = project_root / "pm_sosanhbom.spec"

    if not spec_path.exists():
        logger.error("Spec file not found: %s", spec_path)
        return 1

    logger.info("Starting PyInstaller build using spec: %s...", spec_path)
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        str(spec_path),
    ]

    try:
        ret = subprocess.call(cmd, cwd=str(project_root))
        if ret == 0:
            logger.info("PyInstaller build completed successfully.")
            dist_output = project_root / "dist" / "SSBOM_Portable"
            logger.info("Executable bundle located at: %s", dist_output)
        else:
            logger.error("PyInstaller build exited with error code: %d", ret)
        return ret
    except Exception as exc:
        logger.error("Failed to execute PyInstaller: %s", exc)
        return 2


if __name__ == "__main__":
    sys.exit(build_standalone_exe())

