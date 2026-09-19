"""Credential storage and security management for TC2412 automation.

Conforms to Section 4.3 of SPEC_PLM_AUTO_DOWNLOAD.md and PROJECT.md interface contracts.
Provides native Windows DPAPI storage as the primary mechanism, with seamless fallback
to the OS keyring and mockable providers for testing and CI environments.
Ensures passwords are never stored in cleartext, printed to consoles, or leaked into logs.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import threading
import time
from abc import ABC, abstractmethod
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Self


from .dpapi import (
    IS_WINDOWS,
    DPAPIError,
    KeyringAccessError,
    dpapi_decrypt,
    dpapi_encrypt,
    get_default_storage_dir,
    sanitize_service_name,
)

logger = logging.getLogger("src.security")


class Credentials(tuple):
    """Immutable representation of stored credentials masking sensitive fields.

    Conforms to `tuple[str, str]` contract: can be directly unpacked as
    `(username, password) = creds` or indexed as `creds[0]`, `creds[1]`.
    All string representations mask the raw password as '***' to prevent
    accidental exposure in logs, traces, or terminal outputs.
    """

    def __new__(
        cls,
        username: str,
        password: str,
        service: str = "PM_SOSANHBOM_TC2412",
    ) -> Self:
        if not isinstance(username, str) or not isinstance(password, str):
            raise TypeError("Username and password must be strings")
        obj = super().__new__(cls, (username, password))
        obj._username = username
        obj._password = password
        obj._service = service
        return obj

    @property
    def username(self) -> str:
        """Returns the username."""
        return self._username

    @property
    def password(self) -> str:
        """Returns the cleartext password (for internal authenticated sessions only)."""
        return self._password

    @property
    def service(self) -> str:
        """Returns the associated service identifier."""
        return self._service

    def __repr__(self) -> str:
        return (
            f"Credentials(username={self._username!r}, "
            f"password='***', "
            f"service={self._service!r})"
        )

    def __str__(self) -> str:
        return (
            f"Credentials(username={self._username}, "
            f"password=***, "
            f"service={self._service})"
        )


class BaseCredentialStorage(ABC):
    """Abstract interface for credential storage providers."""

    @abstractmethod
    def is_available(self) -> bool:
        """Returns True if this storage mechanism is functional in the current environment."""

    @abstractmethod
    def save(self, service: str, username: str, password: str) -> bool:
        """Encodes and securely stores credentials."""

    @abstractmethod
    def get(self, service: str) -> Credentials | None:
        """Retrieves and decrypts credentials for the given service."""

    @abstractmethod
    def delete(self, service: str) -> bool:
        """Deletes credentials for the given service."""


class WindowsDPAPIStorage(BaseCredentialStorage):
    """Primary credential storage using native Windows DPAPI (CryptProtectData)."""

    def __init__(self, storage_dir: Path | None = None) -> None:
        self._storage_dir = storage_dir

    @property
    def storage_dir(self) -> Path:
        if self._storage_dir is not None:
            return self._storage_dir
        return get_default_storage_dir()

    def is_available(self) -> bool:
        return IS_WINDOWS

    def _get_file_path(self, service: str) -> Path:
        safe_name = sanitize_service_name(service)
        return self.storage_dir / f"{safe_name}.dpapi"

    def save(self, service: str, username: str, password: str) -> bool:
        if not self.is_available():
            raise DPAPIError("Windows DPAPI is not available on this platform.")

        payload = {
            "version": 1,
            "service": service,
            "username": username,
            "password": password,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }
        raw_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        encrypted_bytes = dpapi_encrypt(
            raw_bytes,
            description=f"PM_SOSANHBOM DPAPI Credential [{service}]",
        )

        target_file = self._get_file_path(service)
        target_file.parent.mkdir(parents=True, exist_ok=True)

        # Atomic file write with unique nonce and retry loop for Windows file-locking
        temp_file = target_file.with_name(
            f"{target_file.stem}_{os.getpid()}_{threading.get_ident()}_{time.time_ns()}.tmp"
        )
        try:
            temp_file.write_bytes(encrypted_bytes)
            max_retries = 20
            for attempt in range(max_retries):
                try:
                    temp_file.replace(target_file)
                    break
                except OSError as replace_err:
                    if attempt == max_retries - 1:
                        raise replace_err
                    time.sleep(0.005 * (attempt + 1))
            logger.info("Successfully encrypted and stored credentials via DPAPI for service '%s'", service)
            return True
        except (OSError, DPAPIError) as exc:
            if temp_file.exists():
                try:
                    temp_file.unlink()
                except OSError as cleanup_err:
                    logger.debug("Failed cleaning temp file: %s", cleanup_err)
            logger.error("Failed to write encrypted DPAPI file: %s", exc)
            raise DPAPIError(f"Failed to persist DPAPI encrypted file: {exc}") from exc

    def get(self, service: str) -> Credentials | None:
        if not self.is_available():
            return None

        target_file = self._get_file_path(service)
        if not target_file.is_file():
            return None

        try:
            encrypted_bytes = None
            for attempt in range(15):
                try:
                    encrypted_bytes = target_file.read_bytes()
                    break
                except OSError:
                    if attempt == 14:
                        return None
                    time.sleep(0.002)

            if not encrypted_bytes:
                return None


            decrypted_bytes = dpapi_decrypt(encrypted_bytes)
            data = json.loads(decrypted_bytes.decode("utf-8"))
            username = data.get("username")
            password = data.get("password")

            if username is not None and password is not None:
                return Credentials(username, password, service)
            return None
        except DPAPIError as exc:
            logger.warning("Corrupted or undecryptable DPAPI credentials for service '%s': %s", service, exc)
            return None
        except (OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
            logger.warning("Failed to load DPAPI credentials for service '%s': %s", service, exc)
            return None

    def delete(self, service: str) -> bool:
        if not self.is_available():
            return True

        target_file = self._get_file_path(service)
        try:
            if target_file.is_file():
                target_file.unlink()
                logger.info("Deleted DPAPI credentials file for service '%s'", service)
            return True
        except OSError as exc:
            logger.error("Failed to delete DPAPI credentials file: %s", exc)
            return False


class KeyringStorage(BaseCredentialStorage):
    """Fallback credential storage using the cross-platform keyring package."""

    def __init__(self, custom_backend: Any = None) -> None:
        self._custom_backend = custom_backend

    def _get_keyring(self) -> Any:
        if self._custom_backend is not None:
            return self._custom_backend
        try:
            import keyring
            return keyring.get_keyring()
        except ImportError:
            return None

    def is_available(self) -> bool:
        try:
            kr = self._get_keyring()
            if kr is None:
                return False
            from keyring.backends import null
            if isinstance(kr, null.Keyring):
                return False
            return getattr(kr, "priority", 1) > 0
        except (ImportError, AttributeError):
            return False

    def save(self, service: str, username: str, password: str) -> bool:
        try:
            kr = self._get_keyring()
            if kr is None:
                raise KeyringAccessError("Keyring module not available")

            kr.set_password(service, username, password)
            # Store username reference so get_credentials(service) can locate the user
            kr.set_password(service, "__current_user__", username)
            logger.info("Successfully stored credentials via Keyring for service '%s'", service)
            return True
        except Exception as exc:
            logger.error("Failed to save credentials in keyring: %s", exc)
            raise KeyringAccessError(f"Keyring save failed: {exc}") from exc

    def get(self, service: str) -> Credentials | None:
        try:
            kr = self._get_keyring()
            if kr is None:
                return None

            # 1. Attempt get_credential directly
            try:
                cred = kr.get_credential(service, None)
                if cred and getattr(cred, "username", None) and getattr(cred, "password", None):
                    if cred.username == "__current_user__":
                        # Windows Vault returned the __current_user__ record where password is username
                        real_user = cred.password
                        pwd = kr.get_password(service, real_user)
                        if pwd is not None:
                            return Credentials(real_user, pwd, service)
                    else:
                        return Credentials(cred.username, cred.password, service)
            except Exception as direct_err:  # noqa: BLE001
                logger.debug("Direct keyring get_credential not supported: %s", direct_err)

            # 2. Attempt lookup via stored __current_user__ reference
            username = kr.get_password(service, "__current_user__")
            if username:
                pwd = kr.get_password(service, username)
                if pwd is not None:
                    return Credentials(username, pwd, service)


            return None
        except Exception as exc:  # noqa: BLE001
            logger.warning("Keyring lookup failed for service '%s': %s", service, exc)
            return None

    def delete(self, service: str) -> bool:
        try:
            kr = self._get_keyring()
            if kr is None:
                return True

            username = None
            try:
                username = kr.get_password(service, "__current_user__")
            except Exception as user_err:  # noqa: BLE001
                logger.debug("Error reading current user marker from keyring: %s", user_err)

            if username:
                try:
                    kr.delete_password(service, username)
                except Exception as del_err:  # noqa: BLE001
                    logger.debug("Error deleting password from keyring: %s", del_err)

            try:
                kr.delete_password(service, "__current_user__")
            except Exception as del_marker_err:  # noqa: BLE001
                logger.debug("Error deleting user marker from keyring: %s", del_marker_err)

            logger.info("Deleted Keyring credentials for service '%s'", service)
            return True
        except Exception as exc:  # noqa: BLE001
            logger.warning("Error deleting credentials from keyring: %s", exc)
            return False


class MockCredentialStorage(BaseCredentialStorage):
    """Mock in-memory credential storage for unit testing and CI pipelines."""

    def __init__(self, available: bool = True) -> None:
        self._store: dict[str, tuple[str, str]] = {}
        self._available = available

    def is_available(self) -> bool:
        return self._available

    def set_available(self, available: bool) -> None:
        self._available = available

    def save(self, service: str, username: str, password: str) -> bool:
        if not self._available:
            raise KeyringAccessError("Mock credential storage is unavailable")
        self._store[service] = (username, password)
        return True

    def get(self, service: str) -> Credentials | None:
        if not self._available:
            raise KeyringAccessError("Mock credential storage is unavailable")
        pair = self._store.get(service)
        if pair:
            return Credentials(pair[0], pair[1], service)
        return None

    def delete(self, service: str) -> bool:
        self._store.pop(service, None)
        return True


class CredentialManager:
    """Quản lý thông tin xác thực an toàn bằng Windows DPAPI và OS Keyring.

    Conforms to Section 4.3 of SPEC_PLM_AUTO_DOWNLOAD.md and PROJECT.md:65-71.
    Primary storage: Native Windows DPAPI.
    Fallback storage: OS Keyring.
    """

    SERVICE_NAME: str = "PM_SOSANHBOM_TC2412"

    _primary_storage: BaseCredentialStorage = WindowsDPAPIStorage()
    _fallback_storage: BaseCredentialStorage = KeyringStorage()

    @classmethod
    def set_storage(
        cls,
        primary: BaseCredentialStorage | None = None,
        fallback: BaseCredentialStorage | None = None,
    ) -> None:
        """Injects custom storage backends (primarily for isolated unit tests)."""
        if primary is not None:
            cls._primary_storage = primary
        if fallback is not None:
            cls._fallback_storage = fallback

    @classmethod
    def set_storage_dir(cls, storage_dir: Path | None) -> None:
        """Sets the directory for primary DPAPI storage."""
        cls._primary_storage = WindowsDPAPIStorage(storage_dir=storage_dir)

    @classmethod
    def reset_storage(cls) -> None:
        """Resets storage backends to default implementations."""
        cls._primary_storage = WindowsDPAPIStorage()
        cls._fallback_storage = KeyringStorage()

    @classmethod
    def get_credentials(
        cls,
        service: str = "PM_SOSANHBOM_TC2412",
    ) -> tuple[str, str] | None:
        """Truy xuất cặp (username, password) từ Windows Credential Manager / DPAPI.

        Args:
            service: Tên định danh dịch vụ bảo mật (mặc định: 'PM_SOSANHBOM_TC2412')

        Returns:
            tuple[str, str] | None: Cặp (username, password) nếu tồn tại, None nếu chưa được lưu
        """
        # 1. Try Primary DPAPI Storage
        if cls._primary_storage.is_available():
            try:
                creds = cls._primary_storage.get(service)
                if creds is not None:
                    return creds
            except Exception as exc:  # noqa: BLE001
                logger.warning("Primary DPAPI retrieval failed for '%s': %s", service, exc)

        # 2. Fallback to Keyring Storage
        if cls._fallback_storage.is_available():
            try:
                creds = cls._fallback_storage.get(service)
                if creds is not None:
                    return creds
            except Exception as exc:  # noqa: BLE001
                logger.warning("Fallback Keyring retrieval failed for '%s': %s", service, exc)

        return None

    @classmethod
    def save_credentials(
        cls,
        username: str,
        password: str,
        service: str = "PM_SOSANHBOM_TC2412",
    ) -> bool:
        """Mã hóa và lưu trữ an toàn username và password vào Windows DPAPI.

        Args:
            username: Tên tài khoản TC2412 (ví dụ: 'vn_pe03')
            password: Mật khẩu plaintext (chỉ tồn tại trong RAM trong thời gian mã hóa)
            service: Tên định danh dịch vụ bảo mật (mặc định: 'PM_SOSANHBOM_TC2412')

        Returns:
            bool: True nếu mã hóa và lưu thành công, False nếu thất bại
        """
        if not username or not password:
            raise ValueError("Username and password must not be empty strings")

        # 1. Try Primary DPAPI Storage
        if cls._primary_storage.is_available():
            try:
                if cls._primary_storage.save(service, username, password):
                    return True
            except Exception as exc:  # noqa: BLE001
                logger.warning("Primary DPAPI save failed for '%s': %s. Falling back to keyring.", service, exc)

        # 2. Fallback to Keyring Storage
        if cls._fallback_storage.is_available():
            try:
                return cls._fallback_storage.save(service, username, password)
            except Exception as exc:
                logger.error("Fallback Keyring save failed for '%s': %s", service, exc)
                raise KeyringAccessError(f"Failed to persist credentials to both DPAPI and Keyring: {exc}") from exc

        raise KeyringAccessError(
            "No functional credential storage provider available (Windows DPAPI and Keyring both unavailable)."
        )

    @classmethod
    def delete_credentials(
        cls,
        service: str = "PM_SOSANHBOM_TC2412",
    ) -> bool:
        """Thu hồi và xóa vĩnh viễn thông tin tài khoản đã lưu khỏi Windows DPAPI.

        Args:
            service: Tên định danh dịch vụ bảo mật (mặc định: 'PM_SOSANHBOM_TC2412')

        Returns:
            bool: True nếu xóa thành công hoặc dịch vụ chưa từng tồn tại, False nếu gặp lỗi
        """
        dpapi_ok = True
        if cls._primary_storage.is_available():
            try:
                dpapi_ok = cls._primary_storage.delete(service)
            except Exception as exc:  # noqa: BLE001
                logger.warning("Error deleting DPAPI credentials for '%s': %s", service, exc)
                dpapi_ok = False

        keyring_ok = True
        if cls._fallback_storage.is_available():
            try:
                keyring_ok = cls._fallback_storage.delete(service)
            except Exception as exc:  # noqa: BLE001
                logger.warning("Error deleting Keyring credentials for '%s': %s", service, exc)
                keyring_ok = False

        return dpapi_ok and keyring_ok

    @classmethod
    def check_credentials(
        cls,
        service: str = "PM_SOSANHBOM_TC2412",
    ) -> bool:
        """Kiểm tra sự tồn tại của thông tin tài khoản mà không để lộ mật khẩu."""
        creds = cls.get_credentials(service)
        return creds is not None


def create_parser() -> argparse.ArgumentParser:
    """Constructs the CLI argument parser."""
    parser = argparse.ArgumentParser(
        description="PM_SOSANHBOM TC2412 Security & DPAPI Credential Manager CLI",
        prog="python -m src.security.credentials",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--save-credentials",
        nargs=2,
        metavar=("USERNAME", "PASSWORD"),
        help="Encrypt and securely save credentials for the given service",
    )
    group.add_argument(
        "--update-credentials",
        nargs=2,
        metavar=("USERNAME", "PASSWORD"),
        help="Update credentials (alias for --save-credentials)",
    )
    group.add_argument(
        "--get-credentials",
        action="store_true",
        help="Retrieve credentials for the service (password is safely masked in output)",
    )
    group.add_argument(
        "--delete-credentials",
        action="store_true",
        help="Delete stored credentials for the service",
    )
    group.add_argument(
        "--check-credentials",
        action="store_true",
        help="Check whether credentials exist for the service without retrieving password",
    )
    parser.add_argument(
        "--service",
        default=CredentialManager.SERVICE_NAME,
        help=f"Target service name (default: {CredentialManager.SERVICE_NAME})",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI execution entrypoint for credential management.

    Returns:
        int: 0 on success, non-zero on failure or missing credentials.
    """
    parser = create_parser()
    args = parser.parse_args(argv)
    service = args.service

    if args.save_credentials or args.update_credentials:
        creds_pair = args.save_credentials or args.update_credentials
        username, password = creds_pair
        try:
            success = CredentialManager.save_credentials(username, password, service=service)
            if success:
                print(f"[SUCCESS] Saved credentials for user '{username}' (service: '{service}').")
                return 0
            print(f"[ERROR] Failed to save credentials for service '{service}'.", file=sys.stderr)
            return 1
        except Exception as exc:  # noqa: BLE001
            print(f"[ERROR] Failed to save credentials: {exc}", file=sys.stderr)
            return 1

    elif args.get_credentials:
        try:
            creds = CredentialManager.get_credentials(service=service)
            if creds is not None:
                # Password must NEVER be logged or printed in cleartext!
                # creds string representation masks password as ***
                print(f"[FOUND] {creds}")
                return 0
            print(f"[NOT FOUND] No credentials found for service '{service}'.", file=sys.stderr)
            return 1
        except Exception as exc:  # noqa: BLE001
            print(f"[ERROR] Failed to get credentials: {exc}", file=sys.stderr)
            return 1

    elif args.check_credentials:
        try:
            exists = CredentialManager.check_credentials(service=service)
            if exists:
                creds = CredentialManager.get_credentials(service=service)
                user_label = creds.username if creds else "configured"
                print(f"[EXISTS] Credentials exist for service '{service}' (user: '{user_label}').")
                return 0
            print(f"[NOT FOUND] No credentials configured for service '{service}'.", file=sys.stderr)
            return 1
        except Exception as exc:  # noqa: BLE001
            print(f"[ERROR] Failed to check credentials: {exc}", file=sys.stderr)
            return 1

    elif args.delete_credentials:
        try:
            success = CredentialManager.delete_credentials(service=service)
            if success:
                print(f"[DELETED] Credentials deleted for service '{service}'.")
                return 0
            print(f"[ERROR] Failed to delete credentials for service '{service}'.", file=sys.stderr)
            return 1
        except Exception as exc:  # noqa: BLE001
            print(f"[ERROR] Failed to delete credentials: {exc}", file=sys.stderr)
            return 1

    else:
        parser.print_help()
        return 0


if __name__ == "__main__":
    sys.exit(main())
