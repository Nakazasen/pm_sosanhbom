"""Unit tests for Update Security and Safe Extraction Service.

Strictly verifies HASH_ONLY_LAN security constraints:
- Safe relative path validation (Zip Slip defense)
- Chunked SHA-256 calculation
- Archive size and member limit enforcement
- Manifest integrity and anti-tamper verification
"""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

import pytest

from src.services.update_security import (
    MAX_EXTRACTED_SIZE_BYTES,
    SafeArchiveExtractor,
    UpdateSecurityError,
    compute_file_sha256,
    is_safe_relative_path,
    verify_manifest_integrity,
)


class TestUpdateSecurityPathValidation:
    """Tests for relative path sanitization."""

    def test_safe_paths_allowed(self) -> None:
        assert is_safe_relative_path("app.py") is True
        assert is_safe_relative_path("src/gui/app.py") is True
        assert is_safe_relative_path("locales/vi.json") is True
        assert is_safe_relative_path("sub/folder/file.txt") is True

    def test_unsafe_paths_rejected(self) -> None:
        assert is_safe_relative_path("") is False
        assert is_safe_relative_path("   ") is False
        assert is_safe_relative_path("../secret.txt") is False
        assert is_safe_relative_path("foo/../../secret.txt") is False
        assert is_safe_relative_path("/etc/passwd") is False
        assert is_safe_relative_path("\\Windows\\System32") is False
        assert is_safe_relative_path("C:\\Program Files") is False
        assert is_safe_relative_path("D:/Sandbox") is False
        assert is_safe_relative_path("file\0null.txt") is False


class TestComputeFileSha256:
    """Tests for SHA-256 calculation."""

    def test_sha256_known_content(self, tmp_path: Path) -> None:
        test_file = tmp_path / "hello.txt"
        content = b"Kyocera Document Solutions Vietnam"
        test_file.write_bytes(content)

        expected_sha = hashlib.sha256(content).hexdigest().lower()
        actual_sha = compute_file_sha256(test_file)

        assert actual_sha == expected_sha
        assert len(actual_sha) == 64

    def test_sha256_missing_file_raises(self, tmp_path: Path) -> None:
        missing = tmp_path / "non_existent.txt"
        with pytest.raises(FileNotFoundError):
            compute_file_sha256(missing)


class TestSafeArchiveExtractor:
    """Tests for safe archive extraction and zip-slip prevention."""

    def test_extract_valid_archive(self, tmp_path: Path) -> None:
        zip_path = tmp_path / "valid.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("file1.txt", "Hello World 1")
            zf.writestr("sub/file2.txt", "Hello World 2")

        target_dir = tmp_path / "extracted"
        extractor = SafeArchiveExtractor()
        extracted = extractor.extract_all(zip_path, target_dir)

        assert len(extracted) == 2
        assert (target_dir / "file1.txt").read_text(encoding="utf-8") == "Hello World 1"
        assert (target_dir / "sub" / "file2.txt").read_text(encoding="utf-8") == "Hello World 2"

    def test_extract_rejects_zip_slip(self, tmp_path: Path) -> None:
        zip_path = tmp_path / "malicious_slip.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("../escaped.txt", "Pwned!")

        target_dir = tmp_path / "extracted"
        extractor = SafeArchiveExtractor()
        with pytest.raises(UpdateSecurityError) as excinfo:
            extractor.extract_all(zip_path, target_dir)
        assert "Insecure path detected" in str(excinfo.value)

    def test_extract_rejects_size_limit(self, tmp_path: Path) -> None:
        zip_path = tmp_path / "bomb.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            # Write 200 bytes of data
            zf.writestr("big.txt", "A" * 200)

        target_dir = tmp_path / "extracted"
        # Set artificially small max_size_bytes to test limit enforcement
        extractor = SafeArchiveExtractor(max_size_bytes=100)
        with pytest.raises(UpdateSecurityError) as excinfo:
            extractor.extract_all(zip_path, target_dir)
        assert "exceeds limit" in str(excinfo.value)


class TestVerifyManifestIntegrity:
    """Tests for manifest and extracted file integrity checking."""

    def test_manifest_verification_success(self, tmp_path: Path) -> None:
        f1 = tmp_path / "SSBOM_App.py"
        f1.write_text("print('hello')", encoding="utf-8")

        f2 = tmp_path / "config.json"
        f2.write_text("{}", encoding="utf-8")

        manifest = {
            "schema": 1,
            "version": "1.0.1",
            "entrypoint": "SSBOM_App.py",
            "files": {
                "SSBOM_App.py": {
                    "sha256": compute_file_sha256(f1),
                    "size": f1.stat().st_size,
                },
                "config.json": {
                    "sha256": compute_file_sha256(f2),
                    "size": f2.stat().st_size,
                },
            },
        }

        # Write manifest.json
        manifest_file = tmp_path / "manifest.json"
        manifest_file.write_text(json.dumps(manifest), encoding="utf-8")

        is_valid, msg = verify_manifest_integrity(manifest, tmp_path)
        assert is_valid is True
        assert msg == "OK"

    def test_manifest_verification_hash_mismatch(self, tmp_path: Path) -> None:
        f1 = tmp_path / "SSBOM_App.py"
        f1.write_text("original", encoding="utf-8")

        manifest = {
            "schema": 1,
            "version": "1.0.1",
            "entrypoint": "SSBOM_App.py",
            "files": {
                "SSBOM_App.py": {
                    "sha256": "0000000000000000000000000000000000000000000000000000000000000000",
                    "size": f1.stat().st_size,
                },
            },
        }

        is_valid, msg = verify_manifest_integrity(manifest, tmp_path)
        assert is_valid is False
        assert "SHA-256 checksum mismatch" in msg

    def test_manifest_verification_foreign_file_detected(self, tmp_path: Path) -> None:
        f1 = tmp_path / "SSBOM_App.py"
        f1.write_text("ok", encoding="utf-8")

        # Extra file not in manifest
        f_bad = tmp_path / "malicious.py"
        f_bad.write_text("malicious", encoding="utf-8")

        manifest = {
            "schema": 1,
            "version": "1.0.1",
            "entrypoint": "SSBOM_App.py",
            "files": {
                "SSBOM_App.py": {
                    "sha256": compute_file_sha256(f1),
                    "size": f1.stat().st_size,
                },
            },
        }

        is_valid, msg = verify_manifest_integrity(manifest, tmp_path)
        assert is_valid is False
        assert "Undeclared foreign file found" in msg
