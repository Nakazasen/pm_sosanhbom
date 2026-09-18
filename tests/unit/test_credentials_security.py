"""Comprehensive unit tests for CredentialManager and DPAPI security subsystem.

Covers:
- Credentials representation, tuple unpacking, and strict password masking (AC-AUTH-05).
- Native Windows DPAPI encryption and decryption roundtrip (CryptProtectData / CryptUnprotectData).
- Per-user DPAPI encrypted file storage, integrity, and atomic writes.
- Keyring fallback coordination when DPAPI is unavailable or encounters errors.
- Graceful degradation on corrupted data and missing credentials.
- Full CLI command interface (--save-credentials, --get-credentials, --check-credentials, --delete-credentials).
- Absolute verification that cleartext passwords are never leaked in logs, files, or CLI output.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.security.credentials import (
    CredentialManager,
    Credentials,
    KeyringStorage,
    MockCredentialStorage,
    WindowsDPAPIStorage,
    main,
)
from src.security.dpapi import (
    IS_WINDOWS,
    DPAPIError,
    KeyringAccessError,
    SecurityError,
    dpapi_decrypt,
    dpapi_encrypt,
    get_default_storage_dir,
    sanitize_service_name,
)


@pytest.fixture(autouse=True)
def isolated_storage(tmp_path: Path):
    """Isolates CredentialManager storage to a temporary directory for each test."""
    test_storage_dir = tmp_path / "test_credentials"
    test_storage_dir.mkdir(parents=True, exist_ok=True)
    CredentialManager.set_storage_dir(test_storage_dir)
    yield test_storage_dir
    CredentialManager.reset_storage()


class TestCredentialsClass:
    """Tests for the Credentials data structure and password masking."""

    def test_credentials_tuple_contract(self):
        creds = Credentials("vn_pe03", "secret_pass_123", service="PM_TEST")
        # Direct unpacking into 2 values
        username, password = creds
        assert username == "vn_pe03"
        assert password == "secret_pass_123"

        # Indexing
        assert creds[0] == "vn_pe03"
        assert creds[1] == "secret_pass_123"
        assert len(creds) == 2
        assert isinstance(creds, tuple)

    def test_credentials_properties(self):
        creds = Credentials("vn_pe03", "secret_pass_123", service="PM_TEST")
        assert creds.username == "vn_pe03"
        assert creds.password == "secret_pass_123"
        assert creds.service == "PM_TEST"

    def test_credentials_masking_in_repr(self):
        raw_secret = "UltraSecret_999!"
        creds = Credentials("vn_pe03", raw_secret, service="PM_TEST")
        repr_str = repr(creds)

        assert raw_secret not in repr_str
        assert "password='***'" in repr_str
        assert "username='vn_pe03'" in repr_str

    def test_credentials_masking_in_str(self):
        raw_secret = "UltraSecret_999!"
        creds = Credentials("vn_pe03", raw_secret, service="PM_TEST")
        str_repr = str(creds)

        assert raw_secret not in str_repr
        assert "password=***" in str_repr
        assert "username=vn_pe03" in str_repr

    def test_credentials_invalid_types_raise(self):
        with pytest.raises(TypeError):
            Credentials(123, "password")  # type: ignore

        with pytest.raises(TypeError):
            Credentials("user", None)  # type: ignore

    def test_credentials_immutability(self):
        creds = Credentials("vn_pe03", "secret")
        with pytest.raises((AttributeError, TypeError)):
            creds[0] = "other"  # type: ignore


@pytest.mark.skipif(not IS_WINDOWS, reason="Native DPAPI requires Windows environment")
class TestWindowsDPAPINative:
    """Tests native Windows DPAPI encryption and decryption via ctypes."""

    def test_dpapi_encrypt_decrypt_roundtrip_ascii(self):
        data = b"my_secure_automation_password_2026"
        encrypted = dpapi_encrypt(data, description="Test ASCII")
        assert encrypted != data
        assert len(encrypted) > 0

        decrypted = dpapi_decrypt(encrypted)
        assert decrypted == data

    def test_dpapi_encrypt_decrypt_roundtrip_unicode(self):
        text = "Mật khẩu tiếng Việt & 日本語のパスワード & 🔑🔒"
        data = text.encode("utf-8")
        encrypted = dpapi_encrypt(data, description="Test Unicode")
        decrypted = dpapi_decrypt(encrypted)
        assert decrypted.decode("utf-8") == text

    def test_dpapi_encrypt_decrypt_empty_data(self):
        data = b""
        encrypted = dpapi_encrypt(data)
        decrypted = dpapi_decrypt(encrypted)
        assert decrypted == b""

    def test_dpapi_entropy_roundtrip(self):
        data = b"sensitive_with_entropy"
        entropy = b"extra_entropy_salt_123"

        encrypted = dpapi_encrypt(data, entropy=entropy)
        decrypted = dpapi_decrypt(encrypted, entropy=entropy)
        assert decrypted == data

    def test_dpapi_entropy_mismatch_raises(self):
        data = b"sensitive_with_entropy"
        entropy = b"correct_entropy"
        wrong_entropy = b"wrong_entropy"

        encrypted = dpapi_encrypt(data, entropy=entropy)
        with pytest.raises(DPAPIError):
            dpapi_decrypt(encrypted, entropy=wrong_entropy)

    def test_dpapi_empty_ciphertext_raises(self):
        with pytest.raises(DPAPIError):
            dpapi_decrypt(b"")

    def test_dpapi_corrupted_ciphertext_raises(self):
        corrupted = b"this_is_not_a_valid_dpapi_ciphertext_blob_1234567890"
        with pytest.raises(DPAPIError):
            dpapi_decrypt(corrupted)


@pytest.mark.skipif(not IS_WINDOWS, reason="Native DPAPI requires Windows environment")
class TestWindowsDPAPIStorage:
    """Tests WindowsDPAPIStorage file-based persistence."""

    def test_storage_save_get_delete_cycle(self, isolated_storage: Path):
        storage = WindowsDPAPIStorage(storage_dir=isolated_storage)
        service = "TEST_DPAPI_SERVICE"
        username = "vn_pe03"
        password = "super_secret_password"

        # Save
        assert storage.save(service, username, password) is True

        # Check file was created
        expected_file = isolated_storage / f"{service}.dpapi"
        assert expected_file.is_file()

        # Check file content is NOT cleartext
        raw_bytes = expected_file.read_bytes()
        assert password.encode() not in raw_bytes

        # Get
        creds = storage.get(service)
        assert creds is not None
        assert creds.username == username
        assert creds.password == password
        assert creds.service == service

        # Delete
        assert storage.delete(service) is True
        assert not expected_file.is_file()

        # Subsequent get returns None
        assert storage.get(service) is None

    def test_storage_corrupted_file_returns_none_gracefully(self, isolated_storage: Path):
        storage = WindowsDPAPIStorage(storage_dir=isolated_storage)
        service = "TEST_CORRUPTED_SERVICE"

        target_file = isolated_storage / f"{service}.dpapi"
        target_file.write_bytes(b"corrupted_garbage_bytes_in_file")

        # Must return None gracefully and not raise unhandled crash
        assert storage.get(service) is None

    def test_storage_empty_file_returns_none(self, isolated_storage: Path):
        storage = WindowsDPAPIStorage(storage_dir=isolated_storage)
        service = "TEST_EMPTY_FILE_SERVICE"

        target_file = isolated_storage / f"{service}.dpapi"
        target_file.write_bytes(b"")

        assert storage.get(service) is None

    def test_storage_nonexistent_returns_none(self, isolated_storage: Path):
        storage = WindowsDPAPIStorage(storage_dir=isolated_storage)
        assert storage.get("NON_EXISTENT_SERVICE") is None

    def test_storage_delete_nonexistent_returns_true(self, isolated_storage: Path):
        storage = WindowsDPAPIStorage(storage_dir=isolated_storage)
        assert storage.delete("NON_EXISTENT_SERVICE") is True

    def test_storage_save_when_unavailable_raises(self):
        storage = WindowsDPAPIStorage()
        with patch.object(storage, "is_available", return_value=False), pytest.raises(DPAPIError):
            storage.save("SVC", "u", "p")


class TestKeyringStorage:
    """Tests KeyringStorage fallback mechanism."""

    def test_keyring_save_get_delete_with_mock_backend(self):
        class FakeKeyringBackend:
            def __init__(self):
                self.vault = {}
                self.priority = 10

            def set_password(self, service, username, password):
                self.vault[(service, username)] = password

            def get_password(self, service, username):
                return self.vault.get((service, username))

            def get_credential(self, service, username):
                # Simulate Windows Vault returning __current_user__ entry when username=None
                if (service, "__current_user__") in self.vault:
                    mock_cred = MagicMock()
                    mock_cred.username = "__current_user__"
                    mock_cred.password = self.vault[(service, "__current_user__")]
                    return mock_cred
                return None

            def delete_password(self, service, username):
                self.vault.pop((service, username), None)

        fake_backend = FakeKeyringBackend()
        storage = KeyringStorage(custom_backend=fake_backend)
        assert storage.is_available() is True

        # Save
        assert storage.save("KEYRING_SVC", "user1", "pass1") is True
        assert fake_backend.vault[("KEYRING_SVC", "user1")] == "pass1"
        assert fake_backend.vault[("KEYRING_SVC", "__current_user__")] == "user1"

        # Get - must return real user and password, NOT __current_user__
        creds = storage.get("KEYRING_SVC")
        assert creds is not None
        assert creds.username == "user1"
        assert creds.password == "pass1"

        # Verify that if only __current_user__ marker exists without user password, it returns None
        fake_backend.vault.pop(("KEYRING_SVC", "user1"))
        assert storage.get("KEYRING_SVC") is None

        # Re-save
        assert storage.save("KEYRING_SVC", "user1", "pass1") is True

        # Delete
        assert storage.delete("KEYRING_SVC") is True
        assert storage.get("KEYRING_SVC") is None
        assert ("KEYRING_SVC", "user1") not in fake_backend.vault
        assert ("KEYRING_SVC", "__current_user__") not in fake_backend.vault

    def test_keyring_get_nonexistent_returns_none(self):
        storage = KeyringStorage()
        assert storage.get("NONEXISTENT_KEYRING_SVC_XYZ") is None


class TestCredentialManager:
    """Tests CredentialManager coordination, fallbacks, and validations."""

    def test_manager_save_and_get_primary_flow(self, isolated_storage: Path):
        service = "PM_TEST_MANAGER_PRIMARY"
        username = "vn_pe03"
        password = "primary_password_123"

        assert CredentialManager.save_credentials(username, password, service=service) is True
        creds = CredentialManager.get_credentials(service=service)

        assert creds is not None
        assert creds[0] == username
        assert creds[1] == password
        assert creds.username == username
        assert creds.password == password

    def test_manager_delete_credentials(self, isolated_storage: Path):
        service = "PM_TEST_MANAGER_DELETE"
        CredentialManager.save_credentials("user", "pass", service=service)
        assert CredentialManager.check_credentials(service=service) is True

        assert CredentialManager.delete_credentials(service=service) is True
        assert CredentialManager.check_credentials(service=service) is False
        assert CredentialManager.get_credentials(service=service) is None

    def test_manager_fallback_to_keyring_when_dpapi_fails_on_save(self):
        mock_dpapi = MockCredentialStorage(available=True)
        mock_keyring = MockCredentialStorage(available=True)

        # Make DPAPI throw on save
        mock_dpapi.save = MagicMock(side_effect=DPAPIError("DPAPI drive full or OS locked"))

        CredentialManager.set_storage(primary=mock_dpapi, fallback=mock_keyring)

        # Save should succeed via keyring fallback
        assert CredentialManager.save_credentials("user_fallback", "pass_fallback", service="FALLBACK_SVC") is True

        # Keyring storage should hold credentials
        assert mock_keyring.get("FALLBACK_SVC") is not None
        assert mock_keyring.get("FALLBACK_SVC").username == "user_fallback"

    def test_manager_fallback_to_keyring_when_dpapi_unavailable(self):
        mock_dpapi = MockCredentialStorage(available=False)
        mock_keyring = MockCredentialStorage(available=True)

        CredentialManager.set_storage(primary=mock_dpapi, fallback=mock_keyring)

        assert CredentialManager.save_credentials("user_non_win", "pass_non_win", service="NON_WIN_SVC") is True
        creds = CredentialManager.get_credentials(service="NON_WIN_SVC")

        assert creds is not None
        assert creds.username == "user_non_win"

    def test_manager_fallback_to_keyring_when_dpapi_returns_none(self):
        mock_dpapi = MockCredentialStorage(available=True)
        mock_keyring = MockCredentialStorage(available=True)

        # Save credentials in keyring only
        mock_keyring.save("KEYRING_ONLY_SVC", "user_kr", "pass_kr")

        CredentialManager.set_storage(primary=mock_dpapi, fallback=mock_keyring)

        # DPAPI has nothing, so it should fetch from Keyring
        creds = CredentialManager.get_credentials("KEYRING_ONLY_SVC")
        assert creds is not None
        assert creds.username == "user_kr"
        assert creds.password == "pass_kr"

    def test_manager_empty_credentials_raise_value_error(self):
        with pytest.raises(ValueError, match="must not be empty"):
            CredentialManager.save_credentials("", "password")

        with pytest.raises(ValueError, match="must not be empty"):
            CredentialManager.save_credentials("user", "")

    def test_manager_no_provider_available_raises_keyring_access_error(self):
        mock_dpapi = MockCredentialStorage(available=False)
        mock_keyring = MockCredentialStorage(available=False)

        CredentialManager.set_storage(primary=mock_dpapi, fallback=mock_keyring)

        with pytest.raises(KeyringAccessError, match="No functional credential storage provider available"):
            CredentialManager.save_credentials("user", "pass", "UNAVAILABLE_SVC")


class TestCLICommands:
    """Tests the CLI command interface and output safety."""

    def test_cli_save_and_get_cycle(self, isolated_storage: Path, capsys):
        service = "PM_TEST_CLI_CYCLE"
        username = "cli_user"
        password = "cli_secret_password"

        # 1. Save
        exit_code = main(["--save-credentials", username, password, "--service", service])
        assert exit_code == 0
        captured = capsys.readouterr()
        assert "[SUCCESS]" in captured.out
        assert username in captured.out
        # Password must NEVER be printed
        assert password not in captured.out
        assert password not in captured.err

        # 2. Get
        exit_code = main(["--get-credentials", "--service", service])
        assert exit_code == 0
        captured = capsys.readouterr()
        assert "[FOUND]" in captured.out
        assert username in captured.out
        assert "password=***" in captured.out
        # Password must NEVER be printed
        assert password not in captured.out

    def test_cli_check_credentials(self, isolated_storage: Path, capsys):
        service = "PM_TEST_CLI_CHECK"

        # Before saving
        exit_code = main(["--check-credentials", "--service", service])
        assert exit_code == 1
        capsys.readouterr()

        # Save
        main(["--save-credentials", "check_user", "check_pass", "--service", service])
        capsys.readouterr()

        # After saving
        exit_code = main(["--check-credentials", "--service", service])
        assert exit_code == 0
        captured = capsys.readouterr()
        assert "[EXISTS]" in captured.out
        assert "check_user" in captured.out
        assert "check_pass" not in captured.out

    def test_cli_update_credentials(self, isolated_storage: Path, capsys):
        service = "PM_TEST_CLI_UPDATE"

        main(["--save-credentials", "update_user", "old_password", "--service", service])
        capsys.readouterr()

        # Update via --update-credentials
        exit_code = main(["--update-credentials", "update_user", "new_password", "--service", service])
        assert exit_code == 0
        captured = capsys.readouterr()
        assert "[SUCCESS]" in captured.out

        # Verify new password is stored
        creds = CredentialManager.get_credentials(service)
        assert creds is not None
        assert creds.password == "new_password"

    def test_cli_delete_credentials(self, isolated_storage: Path, capsys):
        service = "PM_TEST_CLI_DELETE"

        main(["--save-credentials", "del_user", "del_pass", "--service", service])
        capsys.readouterr()

        # Delete
        exit_code = main(["--delete-credentials", "--service", service])
        assert exit_code == 0
        captured = capsys.readouterr()
        assert "[DELETED]" in captured.out

        # Check not found
        exit_code = main(["--get-credentials", "--service", service])
        assert exit_code == 1
        captured = capsys.readouterr()
        assert "[NOT FOUND]" in captured.err

    def test_cli_no_arguments_prints_help(self, capsys):
        exit_code = main([])
        assert exit_code == 0
        captured = capsys.readouterr()
        assert "PM_SOSANHBOM TC2412 Security & DPAPI Credential Manager CLI" in captured.out


class TestSecurityAndAuditing:
    """Adversarial and security assertions."""

    def test_exception_hierarchy(self):
        assert issubclass(KeyringAccessError, SecurityError)
        assert issubclass(DPAPIError, KeyringAccessError)
        assert issubclass(DPAPIError, SecurityError)

    def test_sanitize_service_name(self):
        dirty = 'service/with\\illegal:chars*and?"<>|spaces'
        clean = sanitize_service_name(dirty)
        assert "/" not in clean
        assert "\\" not in clean
        assert ":" not in clean
        assert "*" not in clean
        assert "?" not in clean
        assert '"' not in clean
        assert "<" not in clean
        assert ">" not in clean
        assert "|" not in clean
        assert " " not in clean

    def test_default_storage_dir_resolution(self):
        path = get_default_storage_dir()
        assert isinstance(path, Path)
        assert "pm_sosanhbom" in str(path)

    def test_package_exports_and_lazy_getattr(self):
        import src.security as sec

        # Test attribute access through __getattr__
        assert sec.CredentialManager is not None
        assert sec.Credentials is not None
        assert sec.WindowsDPAPIStorage is not None
        assert sec.KeyringStorage is not None
        assert sec.MockCredentialStorage is not None
        assert sec.DPAPIError is not None
        assert sec.KeyringAccessError is not None
        assert sec.SecurityError is not None
        assert callable(sec.dpapi_encrypt)
        assert callable(sec.dpapi_decrypt)

        # Non-existent attribute raises AttributeError
        with pytest.raises(AttributeError):
            _ = sec.NonExistentClass

        # dir() includes __all__
        dir_list = dir(sec)
        assert "CredentialManager" in dir_list
        assert "dpapi_encrypt" in dir_list

    def test_dpapi_functions_raise_on_non_windows(self):
        with patch("src.security.dpapi.IS_WINDOWS", False):
            with pytest.raises(DPAPIError, match="only supported on Windows"):
                dpapi_encrypt(b"test")
            with pytest.raises(DPAPIError, match="only supported on Windows"):
                dpapi_decrypt(b"test")

    def test_windows_storage_default_dir_property(self):
        storage = WindowsDPAPIStorage()
        assert storage.storage_dir == get_default_storage_dir()

    def test_windows_storage_write_failure_cleans_tmp(self, tmp_path: Path):
        storage = WindowsDPAPIStorage(storage_dir=tmp_path)
        # Mock replace to raise an exception
        with patch.object(Path, "replace", side_effect=OSError("Disk write error")), pytest.raises(
            DPAPIError, match="Failed to persist DPAPI encrypted file"
        ):
            storage.save("SVC_ERR", "user", "pass")

    def test_cli_save_exception_handling(self, capsys):
        with patch.object(CredentialManager, "save_credentials", side_effect=Exception("Storage offline")):
            code = main(["--save-credentials", "u", "p", "--service", "SVC_ERR"])
            assert code == 1
            captured = capsys.readouterr()
            assert "[ERROR] Failed to save credentials: Storage offline" in captured.err

    def test_cli_save_returns_false_handling(self, capsys):
        with patch.object(CredentialManager, "save_credentials", return_value=False):
            code = main(["--save-credentials", "u", "p", "--service", "SVC_ERR"])
            assert code == 1
            captured = capsys.readouterr()
            assert "[ERROR] Failed to save credentials for service 'SVC_ERR'." in captured.err

    def test_cli_get_exception_handling(self, capsys):
        with patch.object(CredentialManager, "get_credentials", side_effect=Exception("DPAPI locked")):
            code = main(["--get-credentials", "--service", "SVC_ERR"])
            assert code == 1
            captured = capsys.readouterr()
            assert "[ERROR] Failed to get credentials: DPAPI locked" in captured.err

    def test_cli_check_exception_handling(self, capsys):
        with patch.object(CredentialManager, "check_credentials", side_effect=Exception("Check failed")):
            code = main(["--check-credentials", "--service", "SVC_ERR"])
            assert code == 1
            captured = capsys.readouterr()
            assert "[ERROR] Failed to check credentials: Check failed" in captured.err

    def test_cli_delete_failure_and_exception_handling(self, capsys):
        # 1. Delete returns False
        with patch.object(CredentialManager, "delete_credentials", return_value=False):
            code = main(["--delete-credentials", "--service", "SVC_ERR"])
            assert code == 1
            captured = capsys.readouterr()
            assert "[ERROR] Failed to delete credentials for service 'SVC_ERR'." in captured.err

        # 2. Delete raises Exception
        with patch.object(CredentialManager, "delete_credentials", side_effect=Exception("Delete locked")):
            code = main(["--delete-credentials", "--service", "SVC_ERR"])
            assert code == 1
            captured = capsys.readouterr()
            assert "[ERROR] Failed to delete credentials: Delete locked" in captured.err

    def test_keyring_storage_null_keyring_unavailable(self):
        from keyring.backends import null

        storage = KeyringStorage(custom_backend=null.Keyring())
        assert storage.is_available() is False

    def test_keyring_storage_get_credential_direct(self):
        class DirectCredBackend:
            def __init__(self):
                self.priority = 5

            def get_credential(self, service, username):
                mock_cred = MagicMock()
                mock_cred.username = "direct_user"
                mock_cred.password = "direct_pass"
                return mock_cred

        storage = KeyringStorage(custom_backend=DirectCredBackend())
        creds = storage.get("SOME_SVC")
        assert creds is not None
        assert creds.username == "direct_user"
        assert creds.password == "direct_pass"
