"""Windows Data Protection API (DPAPI) native wrapper.

Provides secure cryptographic encryption and decryption tied to the current
Windows user's logon credentials (SID) without requiring any external keys.
Used as the primary storage mechanism for TC2412 automation credentials.
"""

from __future__ import annotations

import ctypes
import logging
import os
import re
import sys
from ctypes import wintypes
from pathlib import Path

logger = logging.getLogger(__name__)

# Win32 DPAPI Constants
CRYPTPROTECT_UI_FORBIDDEN = 0x1
CRYPTPROTECT_LOCAL_MACHINE = 0x4


class SecurityError(Exception):
    """Base exception for all security and credential subsystem errors."""


class KeyringAccessError(SecurityError):
    """Raised when secure storage (DPAPI or Keyring) is blocked or inaccessible."""


class DPAPIError(KeyringAccessError):
    """Raised when native Windows DPAPI protect or unprotect operations fail."""


class DATA_BLOB(ctypes.Structure):
    """Win32 CryptoAPI data blob structure."""

    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


# Detect Windows and ctypes.windll availability
IS_WINDOWS = sys.platform == "win32" and hasattr(ctypes, "windll")

if IS_WINDOWS:
    try:
        _crypt32 = ctypes.windll.crypt32
        _kernel32 = ctypes.windll.kernel32

        _CryptProtectData = _crypt32.CryptProtectData
        _CryptProtectData.argtypes = [
            ctypes.POINTER(DATA_BLOB),
            wintypes.LPCWSTR,
            ctypes.POINTER(DATA_BLOB),
            wintypes.LPVOID,
            wintypes.LPVOID,
            wintypes.DWORD,
            ctypes.POINTER(DATA_BLOB),
        ]
        _CryptProtectData.restype = wintypes.BOOL

        _CryptUnprotectData = _crypt32.CryptUnprotectData
        _CryptUnprotectData.argtypes = [
            ctypes.POINTER(DATA_BLOB),
            ctypes.POINTER(wintypes.LPWSTR),
            ctypes.POINTER(DATA_BLOB),
            wintypes.LPVOID,
            wintypes.LPVOID,
            wintypes.DWORD,
            ctypes.POINTER(DATA_BLOB),
        ]
        _CryptUnprotectData.restype = wintypes.BOOL

        _LocalFree = _kernel32.LocalFree
        _LocalFree.argtypes = [wintypes.HLOCAL]
        _LocalFree.restype = wintypes.HLOCAL
    except (OSError, AttributeError) as exc:  # pragma: no cover
        logger.warning("Failed to initialize Windows Crypt32 library: %s", exc)
        IS_WINDOWS = False


def dpapi_encrypt(
    plaintext: bytes,
    description: str = "",
    entropy: bytes | None = None,
    flags: int = CRYPTPROTECT_UI_FORBIDDEN,
) -> bytes:
    """Encrypts plaintext bytes using Windows DPAPI (CryptProtectData).

    Args:
        plaintext: The raw bytes to encrypt.
        description: Optional descriptive string associated with the ciphertext.
        entropy: Optional additional secret bytes required for decryption.
        flags: Win32 flags (default: CRYPTPROTECT_UI_FORBIDDEN for silent execution).

    Returns:
        bytes: The encrypted ciphertext blob.

    Raises:
        DPAPIError: If DPAPI is not supported or the Win32 call fails.
    """
    if not IS_WINDOWS:
        raise DPAPIError("Windows DPAPI is only supported on Windows systems.")

    blob_in = DATA_BLOB(
        len(plaintext),
        ctypes.cast(ctypes.create_string_buffer(plaintext), ctypes.POINTER(ctypes.c_byte)),
    )
    blob_out = DATA_BLOB()

    p_entropy = None
    if entropy is not None:
        blob_entropy = DATA_BLOB(
            len(entropy),
            ctypes.cast(ctypes.create_string_buffer(entropy), ctypes.POINTER(ctypes.c_byte)),
        )
        p_entropy = ctypes.byref(blob_entropy)

    success = _CryptProtectData(
        ctypes.byref(blob_in),
        description,
        p_entropy,
        None,
        None,
        flags,
        ctypes.byref(blob_out),
    )

    if not success:
        err_code = ctypes.GetLastError()
        err_msg = ctypes.FormatError(err_code) if hasattr(ctypes, "FormatError") else f"Code {err_code}"
        raise DPAPIError(f"CryptProtectData failed with error {err_code}: {err_msg}")

    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _LocalFree(blob_out.pbData)


def dpapi_decrypt(
    ciphertext: bytes,
    entropy: bytes | None = None,
    flags: int = CRYPTPROTECT_UI_FORBIDDEN,
) -> bytes:
    """Decrypts ciphertext bytes previously encrypted via Windows DPAPI (CryptUnprotectData).

    Args:
        ciphertext: The encrypted ciphertext blob.
        entropy: Optional entropy bytes that were supplied during encryption.
        flags: Win32 flags (default: CRYPTPROTECT_UI_FORBIDDEN).

    Returns:
        bytes: The decrypted plaintext bytes.

    Raises:
        DPAPIError: If DPAPI is not supported or the ciphertext is invalid/corrupted.
    """
    if not IS_WINDOWS:
        raise DPAPIError("Windows DPAPI is only supported on Windows systems.")

    if not ciphertext:
        raise DPAPIError("Ciphertext blob cannot be empty.")

    blob_in = DATA_BLOB(
        len(ciphertext),
        ctypes.cast(ctypes.create_string_buffer(ciphertext), ctypes.POINTER(ctypes.c_byte)),
    )
    blob_out = DATA_BLOB()

    p_entropy = None
    if entropy is not None:
        blob_entropy = DATA_BLOB(
            len(entropy),
            ctypes.cast(ctypes.create_string_buffer(entropy), ctypes.POINTER(ctypes.c_byte)),
        )
        p_entropy = ctypes.byref(blob_entropy)

    success = _CryptUnprotectData(
        ctypes.byref(blob_in),
        None,
        p_entropy,
        None,
        None,
        flags,
        ctypes.byref(blob_out),
    )

    if not success:
        err_code = ctypes.GetLastError()
        err_msg = ctypes.FormatError(err_code) if hasattr(ctypes, "FormatError") else f"Code {err_code}"
        raise DPAPIError(f"CryptUnprotectData failed with error {err_code}: {err_msg}")

    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _LocalFree(blob_out.pbData)


def get_default_storage_dir() -> Path:
    """Returns the default directory for DPAPI credential persistence."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data and local_app_data.strip():
        return Path(local_app_data) / "pm_sosanhbom" / "credentials"
    return Path.home() / ".pm_sosanhbom" / "credentials"


def sanitize_service_name(service: str) -> str:
    """Sanitizes service identifier for filesystem-safe storage filenames."""
    cleaned = re.sub(r"[^a-zA-Z0-9_\-]", "_", service)
    return cleaned if cleaned else "default_service"
