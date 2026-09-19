"""Security and Safe Archive Extraction for SSBOM LAN Updates.

Policy: HASH_ONLY_LAN
Strictly adheres to: huongdansetup_autoupdate.md
- No digital signatures or PKI required.
- All integrity verified via SHA-256 digests, size checks, safe Zip slip prevention,
  and total extraction boundaries (512 MB maximum).
"""

from __future__ import annotations

import hashlib
import os
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Tuple


class UpdateSecurityError(Exception):
    """Raised when an update package violates safety or integrity rules."""
    pass


# Security constraints per huongdansetup_autoupdate.md
MAX_EXTRACTED_SIZE_BYTES = 512 * 1024 * 1024  # 512 MB
MAX_EXTRACTED_FILES = 20_000
CHUNK_SIZE_BYTES = 65536  # 64 KB chunks for streaming hash


def compute_file_sha256(file_path: Path) -> str:
    """Compute SHA-256 hex digest of a file in 64KB chunks.
    
    Args:
        file_path: Absolute or relative path to file.
        
    Returns:
        64-character lowercase hexadecimal SHA-256 string.
        
    Raises:
        FileNotFoundError: If the file does not exist.
        UpdateSecurityError: If file cannot be read.
    """
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"File not found for hashing: {file_path}")
    
    hasher = hashlib.sha256()
    try:
        with open(path, "rb") as fp:
            while chunk := fp.read(CHUNK_SIZE_BYTES):
                hasher.update(chunk)
    except Exception as exc:
        raise UpdateSecurityError(f"Failed to read file during hash computation: {exc}") from exc
    return hasher.hexdigest().lower()


def is_safe_relative_path(path_str: str) -> bool:
    """Check if a relative path string is strictly safe (anti-Zip Slip).
    
    Rejects:
    - Absolute paths (leading '/', '\\', or drive letter 'C:')
    - Path traversal segments ('..')
    - Null bytes or empty paths
    """
    if not path_str or not path_str.strip():
        return False
    
    if "\0" in path_str:
        return False

    # Normalize slashes
    normalized = path_str.replace("\\", "/")
    
    # Reject drive letters (e.g. C:)
    if len(normalized) > 1 and normalized[1] == ":":
        return False

    # Reject absolute paths
    if normalized.startswith("/"):
        return False

    # Check each segment
    parts = normalized.split("/")
    for part in parts:
        if part == "..":
            return False
        if part == "" and parts.index(part) != len(parts) - 1:
            # Empty segment inside (e.g. foo//bar)
            continue

    return True


class SafeArchiveExtractor:
    """Safely extracts .mpupdate zip archives without allowing Zip Slip or Zip Bombs."""

    def __init__(
        self,
        max_size_bytes: int = MAX_EXTRACTED_SIZE_BYTES,
        max_files: int = MAX_EXTRACTED_FILES,
    ) -> None:
        self.max_size_bytes = max_size_bytes
        self.max_files = max_files

    def validate_archive_members(self, zf: zipfile.ZipFile) -> Tuple[int, int]:
        """Inspect all ZIP members before extracting any byte.
        
        Returns:
            Tuple of (total_files_count, total_uncompressed_bytes)
            
        Raises:
            UpdateSecurityError: If any member violates path safety or size limits.
        """
        infolist = zf.infolist()
        if len(infolist) > self.max_files:
            raise UpdateSecurityError(
                f"Archive contains {len(infolist)} entries, exceeding maximum allowed of {self.max_files}."
            )

        total_bytes = 0
        file_count = 0

        for info in infolist:
            # Check for path safety
            if not is_safe_relative_path(info.filename):
                raise UpdateSecurityError(
                    f"Insecure path detected in archive entry: '{info.filename}'. "
                    f"Path traversal and absolute paths are strictly forbidden."
                )

            # Check for symlinks or weird attributes
            if (info.external_attr >> 16) & 0o120000 == 0o120000:
                raise UpdateSecurityError(
                    f"Symlink detected in archive entry: '{info.filename}'. Symlinks are forbidden."
                )

            total_bytes += info.file_size
            if not info.is_dir():
                file_count += 1

            if total_bytes > self.max_size_bytes:
                raise UpdateSecurityError(
                    f"Archive uncompressed size exceeds limit of {self.max_size_bytes} bytes "
                    f"(current accumulator: {total_bytes} bytes)."
                )

        return file_count, total_bytes

    def extract_all(self, archive_path: Path, target_dir: Path) -> List[Path]:
        """Safely extract all files from archive_path into target_dir.
        
        Args:
            archive_path: Path to the .mpupdate or .zip file.
            target_dir: Destination folder to extract into.
            
        Returns:
            List of Path objects for all extracted files.
            
        Raises:
            UpdateSecurityError: If extraction fails or limits are breached.
        """
        archive = Path(archive_path).resolve()
        target = Path(target_dir).resolve()
        target.mkdir(parents=True, exist_ok=True)

        if not archive.is_file():
            raise FileNotFoundError(f"Archive file does not exist: {archive}")

        extracted_files: List[Path] = []

        try:
            with zipfile.ZipFile(archive, "r") as zf:
                # First pass: preflight validation
                self.validate_archive_members(zf)

                # Second pass: safe extraction
                bytes_extracted = 0
                for member in zf.infolist():
                    rel_clean = os.path.normpath(member.filename.replace("\\", "/"))
                    if not is_safe_relative_path(rel_clean):
                        raise UpdateSecurityError(f"Rejected path during extraction: '{rel_clean}'")

                    dest_path = (target / rel_clean).resolve()
                    # Ensure resolved destination is strictly within target_dir
                    if not str(dest_path).startswith(str(target)):
                        raise UpdateSecurityError(
                            f"Extraction path escapes target directory: {dest_path} not in {target}"
                        )

                    if member.is_dir():
                        dest_path.mkdir(parents=True, exist_ok=True)
                        continue

                    dest_path.parent.mkdir(parents=True, exist_ok=True)

                    # Extract file with byte-level counting
                    with zf.open(member, "r") as src_f, open(dest_path, "wb") as dst_f:
                        while chunk := src_f.read(CHUNK_SIZE_BYTES):
                            bytes_extracted += len(chunk)
                            if bytes_extracted > self.max_size_bytes:
                                raise UpdateSecurityError(
                                    f"Total extracted bytes exceeded limit of {self.max_size_bytes}."
                                )
                            dst_f.write(chunk)

                    extracted_files.append(dest_path)

        except zipfile.BadZipFile as exc:
            raise UpdateSecurityError(f"Corrupted or invalid zip archive: {exc}") from exc
        except Exception as exc:
            if isinstance(exc, UpdateSecurityError):
                raise
            raise UpdateSecurityError(f"Error during archive extraction: {exc}") from exc

        return extracted_files


def verify_manifest_integrity(
    manifest_data: Dict[str, Any],
    extracted_dir: Path,
) -> Tuple[bool, str]:
    """Verify that all files in extracted_dir match manifest_data exactly.
    
    Checks:
    1. Schema validity & required keys.
    2. Every file in manifest exists in extracted_dir with matching size and SHA-256.
    3. No undeclared extra files exist in extracted_dir (except manifest.json itself).
    
    Returns:
        (True, "OK") if intact, or (False, error_message) if mismatched.
    """
    target = Path(extracted_dir).resolve()
    if not target.is_dir():
        return False, f"Target directory does not exist: {target}"

    # Verify manifest fields
    if not isinstance(manifest_data, dict):
        return False, "Manifest data is not a valid JSON dictionary."

    version = manifest_data.get("version")
    if not version:
        return False, "Manifest missing 'version' field."

    entrypoint = manifest_data.get("entrypoint")
    if not entrypoint or not is_safe_relative_path(entrypoint):
        return False, f"Manifest missing or has invalid 'entrypoint': {entrypoint}"

    files_map = manifest_data.get("files")
    if not isinstance(files_map, dict):
        return False, "Manifest 'files' field must be a dictionary of path -> metadata."

    # 1. Verify declared files
    for rel_path, meta in files_map.items():
        if not is_safe_relative_path(rel_path):
            return False, f"Manifest declared unsafe file path: '{rel_path}'"

        file_on_disk = target / rel_path
        if not file_on_disk.is_file():
            return False, f"Missing file declared in manifest: '{rel_path}'"

        expected_size = meta.get("size")
        actual_size = file_on_disk.stat().st_size
        if expected_size is not None and actual_size != expected_size:
            return False, (
                f"File size mismatch for '{rel_path}': "
                f"manifest expected {expected_size} bytes, got {actual_size} bytes."
            )

        expected_sha = (meta.get("sha256") or "").lower()
        if not expected_sha:
            return False, f"Manifest missing 'sha256' for file: '{rel_path}'"

        actual_sha = compute_file_sha256(file_on_disk)
        if actual_sha != expected_sha:
            return False, (
                f"SHA-256 checksum mismatch for '{rel_path}': "
                f"expected {expected_sha}, got {actual_sha}."
            )

    # 2. Check for undeclared extra files
    allowed_rel_paths = set(files_map.keys())
    allowed_rel_paths.add("manifest.json")

    for disk_file in target.rglob("*"):
        if disk_file.is_file():
            rel_disk = disk_file.relative_to(target).as_posix()
            if rel_disk not in allowed_rel_paths:
                return False, f"Undeclared foreign file found in extracted directory: '{rel_disk}'"

    # 3. Verify entrypoint file exists
    entry_path = target / entrypoint
    if not entry_path.is_file():
        return False, f"Entrypoint declared in manifest does not exist: '{entrypoint}'"

    return True, "OK"
