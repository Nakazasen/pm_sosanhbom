"""Security and credential management package for TC2412 automation."""

from __future__ import annotations

from typing import Any

__all__ = [
    "DATA_BLOB",
    "BaseCredentialStorage",
    "CredentialManager",
    "Credentials",
    "DPAPIError",
    "KeyringAccessError",
    "KeyringStorage",
    "MockCredentialStorage",
    "SecurityError",
    "WindowsDPAPIStorage",
    "create_parser",
    "dpapi_decrypt",
    "dpapi_encrypt",
    "main",
]


def __getattr__(name: str) -> Any:
    if name in (
        "BaseCredentialStorage",
        "CredentialManager",
        "Credentials",
        "KeyringStorage",
        "MockCredentialStorage",
        "WindowsDPAPIStorage",
        "create_parser",
        "main",
    ):
        from . import credentials
        return getattr(credentials, name)
    if name in (
        "DATA_BLOB",
        "DPAPIError",
        "KeyringAccessError",
        "SecurityError",
        "dpapi_decrypt",
        "dpapi_encrypt",
    ):
        from . import dpapi
        return getattr(dpapi, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __dir__() -> list[str]:
    return sorted(list(globals().keys()) + __all__)
