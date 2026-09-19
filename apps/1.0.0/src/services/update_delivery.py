"""Update Delivery Service for SSBOM Manager.

Handles discovering, checking, and safely caching update packages from LAN sources.
Strictly adheres to: huongdansetup_autoupdate.md (HASH_ONLY_LAN policy).
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from src.services.update_security import (
    CHUNK_SIZE_BYTES,
    UpdateSecurityError,
    compute_file_sha256,
    is_safe_relative_path,
)

logger = logging.getLogger("update_delivery")


class UpdateDeliveryError(Exception):
    """Raised when discovering or downloading updates fails."""
    pass


@dataclass
class UpdateCandidate:
    """Represents a validated update candidate discovered from a catalog."""
    version: str
    package_name: str
    package_sha256: str
    package_size: int
    release_notes: str
    source_location: str
    local_cached_package: Optional[Path] = None


def parse_version_tuple(v_str: str) -> Tuple[int, ...]:
    """Parse a semantic or dotted version string into a tuple of integers.
    
    Examples:
        '1.0.0' -> (1, 0, 0)
        '0.1.7' -> (0, 1, 7)
        '1.2.3.4' -> (1, 2, 3, 4)
    """
    clean = re.sub(r"[^\d.]", "", v_str.strip())
    if not clean:
        return (0,)
    parts = clean.split(".")
    try:
        return tuple(int(p) for p in parts if p)
    except ValueError:
        return (0,)


def is_version_greater(candidate_ver: str, current_ver: str) -> bool:
    """Return True if candidate_ver is strictly greater than current_ver."""
    return parse_version_tuple(candidate_ver) > parse_version_tuple(current_ver)


class UpdateDeliveryService:
    """Manages update sources configuration, latest.json catalog inspection,

    and safe atomic downloading into the local update cache.
    """

    def __init__(
        self,
        base_dir: Path,
        current_version: str = "1.0.0",
        programdata_dir: Optional[Path] = None,
    ) -> None:
        self.base_dir = Path(base_dir).resolve()
        self.current_version = current_version
        self.programdata_dir = (
            Path(programdata_dir).resolve()
            if programdata_dir
            else Path(os.environ.get("PROGRAMDATA", "C:/ProgramData")) / "SSBOM_Manager"
        )
        self.cache_dir = self.base_dir / ".updates" / "downloads"
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def load_sources_configuration(self) -> Dict[str, Any]:
        """Load sources config in ascending priority:

        1. Default bundled: update_sources.default.json
        2. Runtime override: <base_dir>/update_sources.json
        3. Company policy: %PROGRAMDATA%/SSBOM_Manager/update_sources.json
        """
        config: Dict[str, Any] = {
            "schema": 1,
            "startup_check": True,
            "sources": [],
        }

        # 1. Default bundled
        default_file = self.base_dir / "update_sources.default.json"
        if default_file.is_file():
            try:
                with open(default_file, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    if isinstance(data, dict):
                        config.update(data)
            except Exception as exc:
                logger.warning("Could not read bundled sources config: %s", exc)

        # 2. Local user override
        override_file = self.base_dir / "update_sources.json"
        if override_file.is_file():
            try:
                with open(override_file, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    if isinstance(data, dict):
                        config.update(data)
            except Exception as exc:
                logger.warning("Could not read local override sources config: %s", exc)

        # 3. Company policy
        policy_file = self.programdata_dir / "update_sources.json"
        if policy_file.is_file():
            try:
                with open(policy_file, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    if isinstance(data, dict):
                        config.update(data)
            except Exception as exc:
                logger.warning("Could not read company policy sources config: %s", exc)

        return config

    def check_for_updates(self) -> Optional[UpdateCandidate]:
        """Query enabled sources for the highest valid update candidate greater than current version."""
        config = self.load_sources_configuration()
        sources = config.get("sources", [])
        best_candidate: Optional[UpdateCandidate] = None

        for src in sources:
            if not isinstance(src, dict) or not src.get("enabled", True):
                continue

            src_type = src.get("type", "folder")
            location = src.get("location", "")
            if not location:
                continue

            try:
                candidate = self._inspect_source(src_type, location)
                if candidate and is_version_greater(candidate.version, self.current_version):
                    if (
                        best_candidate is None
                        or is_version_greater(candidate.version, best_candidate.version)
                    ):
                        best_candidate = candidate
            except Exception as exc:
                logger.warning("Failed inspecting update source '%s': %s", location, exc)

        return best_candidate

    def _inspect_source(self, src_type: str, location: str) -> Optional[UpdateCandidate]:
        """Inspect a single source for latest.json catalog."""
        if src_type == "folder":
            folder_path = Path(location)
            if not folder_path.exists():
                logger.debug("LAN source folder does not exist: %s", folder_path)
                return None

            latest_file = folder_path / "latest.json"
            if not latest_file.is_file():
                logger.debug("No latest.json found in: %s", folder_path)
                return None

            with open(latest_file, "r", encoding="utf-8") as fp:
                catalog_data = json.load(fp)

            return self._parse_catalog_data(catalog_data, source_location=location)

        elif src_type in ("http", "https"):
            url = location.rstrip("/") + "/latest.json"
            req = urllib.request.Request(url, headers={"User-Agent": "SSBOM_UpdateClient/1.0"})
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.status == 200:
                    raw_data = response.read().decode("utf-8")
                    catalog_data = json.loads(raw_data)
                    return self._parse_catalog_data(catalog_data, source_location=location)
            return None

        else:
            logger.warning("Unsupported source type: %s", src_type)
            return None

    def _parse_catalog_data(
        self, catalog_data: Dict[str, Any], source_location: str
    ) -> Optional[UpdateCandidate]:
        """Validate catalog data schema and return UpdateCandidate."""
        if not isinstance(catalog_data, dict):
            raise UpdateDeliveryError("Invalid catalog format: root must be a JSON object.")

        version = catalog_data.get("version")
        if not version or not isinstance(version, str):
            raise UpdateDeliveryError("Catalog missing valid 'version' string.")

        pkg_name = catalog_data.get("package") or catalog_data.get("package_name")
        if not pkg_name or not isinstance(pkg_name, str) or not is_safe_relative_path(pkg_name):
            raise UpdateDeliveryError(f"Catalog has invalid package filename: '{pkg_name}'")

        pkg_sha256 = (catalog_data.get("sha256") or catalog_data.get("package_sha256") or "").lower()
        if len(pkg_sha256) != 64 or not re.match(r"^[0-9a-f]{64}$", pkg_sha256):
            raise UpdateDeliveryError(f"Catalog has invalid 64-hex SHA-256: '{pkg_sha256}'")

        pkg_size = catalog_data.get("size") or catalog_data.get("package_size") or 0
        if not isinstance(pkg_size, int) or pkg_size <= 0:
            raise UpdateDeliveryError(f"Catalog has invalid package size: {pkg_size}")

        release_notes = str(catalog_data.get("notes") or catalog_data.get("release_notes") or "")

        return UpdateCandidate(
            version=version,
            package_name=pkg_name,
            package_sha256=pkg_sha256,
            package_size=pkg_size,
            release_notes=release_notes,
            source_location=source_location,
        )

    def download_package(
        self,
        candidate: UpdateCandidate,
        progress_callback: Optional[Any] = None,
    ) -> Path:
        """Download or copy package atomically into .updates/downloads/<pkg>.

        Steps:
        1. Target is <cache_dir>/<package_name>. If already exists and matches hash/size, return it.
        2. Stream copy to <cache_dir>/<package_name>.part.
        3. Verify size and SHA-256 of the downloaded .part.
        4. Atomic rename .part -> final filename.
        """
        final_path = self.cache_dir / candidate.package_name

        # Check existing cached package
        if final_path.is_file():
            if (
                final_path.stat().st_size == candidate.package_size
                and compute_file_sha256(final_path) == candidate.package_sha256
            ):
                logger.info("Found valid existing cached package: %s", final_path)
                candidate.local_cached_package = final_path
                return final_path
            else:
                # Corrupted or stale cache, delete it
                try:
                    final_path.unlink()
                except OSError:
                    pass

        part_path = self.cache_dir / f"{candidate.package_name}.part"
        if part_path.exists():
            try:
                part_path.unlink()
            except OSError:
                pass

        logger.info(
            "Downloading/copying update package %s from %s...",
            candidate.package_name,
            candidate.source_location,
        )

        source_loc = candidate.source_location
        if source_loc.startswith("http://") or source_loc.startswith("https://"):
            pkg_url = source_loc.rstrip("/") + "/" + candidate.package_name
            req = urllib.request.Request(pkg_url, headers={"User-Agent": "SSBOM_UpdateClient/1.0"})
            with urllib.request.urlopen(req, timeout=30) as resp, open(part_path, "wb") as out_fp:
                downloaded = 0
                while chunk := resp.read(CHUNK_SIZE_BYTES):
                    out_fp.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(downloaded, candidate.package_size)
        else:
            # File / LAN UNC path
            source_folder = Path(source_loc)
            source_file = source_folder / candidate.package_name
            if not source_file.is_file():
                raise FileNotFoundError(f"Package file not found at source: {source_file}")

            # Safe chunked copy with progress callback
            with open(source_file, "rb") as in_fp, open(part_path, "wb") as out_fp:
                copied = 0
                while chunk := in_fp.read(CHUNK_SIZE_BYTES):
                    out_fp.write(chunk)
                    copied += len(chunk)
                    if progress_callback:
                        progress_callback(copied, candidate.package_size)

        # Verify size
        actual_size = part_path.stat().st_size
        if actual_size != candidate.package_size:
            part_path.unlink(missing_ok=True)
            raise UpdateDeliveryError(
                f"Downloaded package size mismatch: expected {candidate.package_size} bytes, "
                f"got {actual_size} bytes."
            )

        # Verify SHA-256
        actual_sha256 = compute_file_sha256(part_path)
        if actual_sha256 != candidate.package_sha256:
            part_path.unlink(missing_ok=True)
            raise UpdateDeliveryError(
                f"Downloaded package SHA-256 mismatch: expected {candidate.package_sha256}, "
                f"got {actual_sha256}."
            )

        # Atomic rename
        part_path.replace(final_path)
        logger.info("Successfully verified and cached update package to: %s", final_path)
        candidate.local_cached_package = final_path
        return final_path
