"""Path and Environment Utility Module for SSBOM.

Provides centralized, robust path resolution and safe filesystem operations
to prevent unhandled OS errors (e.g. BitLocker drive locks, missing drives,
network SMB disconnects) across dev, portable, and installed environments.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
import socket
import sys
from typing import Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Global Runtime Patch for Python 3.13+ Pathlib BitLocker / Device Lock Handling
# ---------------------------------------------------------------------------
# Python 3.13 introduces pathlib._abc where Path.exists(), is_file(), is_dir()
# catch only errors listed in _IGNORED_WINERRORS. BitLocker volume locked
# (WinError -2144272384 / 0x80310000 / ERROR_FVE_LOCKED_VOLUME) and device not ready
# are not in the standard library tuple, causing unhandled OSError crashes.
# Patching this tuple guarantees that all Path.exists() calls across all libraries
# and codebase return False cleanly instead of crashing the process.
try:
    import pathlib._abc
    if hasattr(pathlib._abc, "_IGNORED_WINERRORS"):
        _extra_winerrors = (-2144272384, 0x80310000, 21, 433)
        _curr = tuple(pathlib._abc._IGNORED_WINERRORS)
        pathlib._abc._IGNORED_WINERRORS = _curr + tuple(w for w in _extra_winerrors if w not in _curr)
except Exception:
    pass


def safe_exists(path: Path | str | None) -> bool:
    """Safely check if a path exists, catching BitLocker and OS errors.

    Guards against:
    - Empty string or whitespace (which Path('') resolves to Path('.'), giving True)
    - OSError: [WinError -2144272384] Drive is locked by BitLocker Drive Encryption
    - OSError: [WinError 21] The device is not ready
    - PermissionError, FileNotFoundError, ValueError
    """
    if path is None:
        return False
    p_str = str(path).strip()
    if not p_str:
        return False
    try:
        p = Path(p_str)
        return p.exists()
    except (OSError, ValueError):
        return False


def safe_is_file(path: Path | str | None) -> bool:
    """Safely check if a path is an existing regular file."""
    if path is None:
        return False
    p_str = str(path).strip()
    if not p_str:
        return False
    try:
        p = Path(p_str)
        return p.is_file()
    except (OSError, ValueError):
        return False


def safe_is_dir(path: Path | str | None) -> bool:
    """Safely check if a path is an existing directory."""
    if path is None:
        return False
    p_str = str(path).strip()
    if not p_str:
        return False
    try:
        p = Path(p_str)
        return p.is_dir()
    except (OSError, ValueError):
        return False


import threading
import time

_unc_host_cache: dict[str, tuple[bool, float]] = {}


def is_unc_path_accessible(path: Path | str | None, timeout: float = 0.3) -> bool:
    """Fast pre-flight network reachability test for UNC paths.

    Prevents Windows SMB and DNS from blocking the main GUI thread for seconds
    when a network server is offline or unreachable. Uses a daemon thread to enforce
    the timeout against Windows DNS resolution hangs, and caches host status for 15 seconds.

    Args:
        path: Path to test (e.g. \\\\fstvn01\\Data\\... or //fstvn01/Data/...).
        timeout: Maximum duration to wait for host reachability (default 0.3s).

    Returns:
        True if the host is reachable or if path is a local drive; False otherwise.
    """
    if path is None:
        return False
    p_str = str(path).strip()
    if not p_str:
        return False
    p_norm = p_str.replace("/", "\\")
    if p_norm.startswith(r"\\"):
        parts = [part for part in p_norm.lstrip("\\").split("\\") if part]
        host = parts[0] if parts else ""
        if not host:
            return False

        now = time.time()
        cached = _unc_host_cache.get(host)
        if cached is not None and (now - cached[1] < 15.0):
            return cached[0]

        result = [False]

        def _check() -> None:
            try:
                with socket.create_connection((host, 445), timeout=timeout):
                    result[0] = True
            except Exception:
                pass

        t = threading.Thread(target=_check, daemon=True)
        t.start()
        t.join(timeout=timeout)

        is_avail = result[0]
        _unc_host_cache[host] = (is_avail, now)
        return is_avail

    return safe_exists(p_str)


def get_app_base_dir() -> Path:
    """Resolve the canonical base directory for the SSBOM application.

    Priority:
    1. SSBOM_BASE_DIR environment variable (if explicitly set and accessible).
    2. If frozen (PyInstaller executable):
       - If placed inside an apps/<version>/ folder (standard autoupdate layout):
         install_root is the great-grandparent (executable.parent.parent.parent).
       - Otherwise:
         install_root is executable.parent.
    3. If running from Python source:
       - Repo root from src/core/path_utils.py -> parents[2].
    4. Fallback to current working directory (Path.cwd()).

    CRITICAL: Never fall back to a hardcoded drive path like 'D:\\Sandbox\\...'
    which crashes with WinError on machines where drive D: is BitLocker locked or missing!
    """
    env_dir = os.environ.get("SSBOM_BASE_DIR")
    if env_dir:
        cand = Path(env_dir)
        if safe_exists(cand):
            return cand.resolve()

    if getattr(sys, "frozen", False):
        exe_path = Path(sys.executable).resolve()
        # Check standard autoupdate structure: <install_root>/apps/<version>/SSBOM_Portable.exe
        try:
            if exe_path.parent.parent.name.lower() == "apps":
                install_root = exe_path.parent.parent.parent
                if safe_exists(install_root):
                    return install_root
        except Exception:
            pass
        return exe_path.parent

    # Source code layout: <repo_root>/src/core/path_utils.py
    try:
        repo_root = Path(__file__).resolve().parents[2]
        if safe_exists(repo_root / "src") or safe_exists(repo_root / "config") or safe_exists(repo_root / "assets"):
            return repo_root
    except Exception:
        pass

    return Path.cwd().resolve()


def get_asset_path(filename: str, base_dir: Path | str | None = None) -> Path | None:
    """Locate an asset file safely across bundled PyInstaller and install directory.

    Guards against BitLocker locked drives and missing folders.

    Args:
        filename: Asset file name (e.g. 'icon.png', 'icon.ico').
        base_dir: Optional explicit base directory.

    Returns:
        Valid Path to the asset file if found, else None.
    """
    candidates: list[Path] = []

    # 1. PyInstaller _MEIPASS (onefile or bundled datas)
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        meipass = Path(getattr(sys, "_MEIPASS"))
        candidates.append(meipass / "assets" / filename)
        candidates.append(meipass / filename)

    # 2. Alongside executable (_internal/assets or assets)
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "_internal" / "assets" / filename)
        candidates.append(exe_dir / "assets" / filename)

    # 3. Base directory candidates
    b_dir = Path(base_dir) if base_dir else get_app_base_dir()
    candidates.append(b_dir / "assets" / filename)
    candidates.append(b_dir / "src" / "gui" / "assets" / filename)

    for cand in candidates:
        if safe_is_file(cand):
            return cand

    return None
