"""Adversarial Verification of CLI Interface and Fallback Behavior for Milestone M2.

This test suite executes empirical subprocess invocations against the CLI interface:
`python -m src.security.credentials`

Verification Scope:
1. Subprocess CLI command lifecycle:
   - `--save-credentials <USER> <PASS>`
   - `--get-credentials`
   - `--check-credentials`
   - `--update-credentials <USER> <NEW_PASS>`
   - `--delete-credentials`
2. Strict security audit:
   - ZERO cleartext password leakage in stdout or stderr across all command variations,
     including ASCII, unicode, shell meta-characters, and extreme length strings.
3. DPAPI Failure Simulation & Transparent Keyring Fallback:
   - Simulated DPAPI protect failure (CryptProtectData / dpapi_encrypt raises DPAPIError).
   - Simulated DPAPI unprotect/decryption failure.
   - Simulated non-Windows / DPAPI disabled environment (IS_WINDOWS=False).
   - Verifies transparent fallback to KeyringStorage with zero cleartext leaks.
4. CLI Argument Robustness & Fail-Closed Guardrails:
   - Mutually exclusive flag enforcement.
   - Missing argument validation.
   - Corrupted storage file handling without unhandled crash or trace dumps.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Any

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent.parent


def run_cli_subprocess(
    args: list[str],
    *,
    env_overrides: dict[str, str] | None = None,
    simulate_dpapi_failure: bool = False,
    simulate_non_windows: bool = False,
    simulate_dpapi_decrypt_failure: bool = False,
    timeout: float = 15.0,
) -> subprocess.CompletedProcess[str]:
    """Invokes the CLI via subprocess.run, optionally simulating failures."""
    env = dict(os.environ)
    env["PYTHONPATH"] = str(REPO_ROOT)
    if env_overrides:
        env.update(env_overrides)

    if not simulate_dpapi_failure and not simulate_non_windows and not simulate_dpapi_decrypt_failure:
        cmd = [sys.executable, "-m", "src.security.credentials", *args]
    else:
        # Wrapper script that applies monkeypatches in the subprocess before running main()
        wrapper_lines = [
            "import sys",
            "from unittest.mock import patch",
            "import src.security.dpapi as d",
            "import src.security.credentials as c",
        ]
        if simulate_non_windows:
            wrapper_lines.append("patch.object(d, 'IS_WINDOWS', False).start()")
        if simulate_dpapi_failure:
            wrapper_lines.append(
                "patch('src.security.credentials.dpapi_encrypt', "
                "side_effect=d.DPAPIError('Simulated Win32 CryptProtectData failure (Access Denied 0x80070005)')).start()"
            )
        if simulate_dpapi_decrypt_failure:
            wrapper_lines.append(
                "patch('src.security.credentials.dpapi_decrypt', "
                "side_effect=d.DPAPIError('Simulated Win32 CryptUnprotectData failure (Corrupted Ciphertext)')).start()"
            )
        wrapper_lines.append("sys.exit(c.main(sys.argv[1:]))")
        wrapper_code = "\n".join(wrapper_lines)

        cmd = [sys.executable, "-c", wrapper_code, *args]

    return subprocess.run(
        cmd,
        cwd=str(REPO_ROOT),
        env=env,
        capture_output=True,
        text=True,
        timeout=timeout,
        check=False,
    )


def assert_no_cleartext_password(
    result: subprocess.CompletedProcess[str],
    password: str,
    context: str = "",
) -> None:
    """Strict oracle: ensures password is never present in stdout or stderr."""
    if not password:
        return
    assert password not in result.stdout, (
        f"CRITICAL SECURITY VIOLATION ({context}): Cleartext password found in STDOUT!\n"
        f"Password: {password}\nSTDOUT: {result.stdout}"
    )
    assert password not in result.stderr, (
        f"CRITICAL SECURITY VIOLATION ({context}): Cleartext password found in STDERR!\n"
        f"Password: {password}\nSTDERR: {result.stderr}"
    )


class TestCLISubprocessLifecycle:
    """Adversarial verification of CLI commands in full subprocess execution."""

    @pytest.fixture
    def isolated_env(self, tmp_path: Path):
        """Creates a completely isolated LOCALAPPDATA directory for each test."""
        app_data = tmp_path / "appdata"
        app_data.mkdir(parents=True, exist_ok=True)
        return {"LOCALAPPDATA": str(app_data)}

    @pytest.fixture
    def test_service(self):
        """Generates a unique service name to avoid collisions."""
        return f"ADV_CLI_TEST_{uuid.uuid4().hex[:8]}"

    def test_subprocess_full_lifecycle_success(
        self,
        isolated_env: dict[str, str],
        test_service: str,
    ):
        """Tests save -> check -> get -> update -> delete -> check cycle via subprocess."""
        username = "adv_user_01"
        password = "Secr3t!P@ssw0rd#2026"

        # 1. Check before save -> Should exit 1 (Not Found)
        r_pre = run_cli_subprocess(
            ["--check-credentials", "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_pre.returncode == 1
        assert "[NOT FOUND]" in r_pre.stderr
        assert_no_cleartext_password(r_pre, password, "pre-check")

        # 2. Save credentials -> Should exit 0
        r_save = run_cli_subprocess(
            ["--save-credentials", username, password, "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_save.returncode == 0
        assert "[SUCCESS]" in r_save.stdout
        assert username in r_save.stdout
        assert test_service in r_save.stdout
        assert_no_cleartext_password(r_save, password, "save")

        # 3. Check credentials -> Should exit 0
        r_check = run_cli_subprocess(
            ["--check-credentials", "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_check.returncode == 0
        assert "[EXISTS]" in r_check.stdout
        assert username in r_check.stdout
        assert_no_cleartext_password(r_check, password, "check")

        # 4. Get credentials -> Should exit 0, output masked password
        r_get = run_cli_subprocess(
            ["--get-credentials", "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_get.returncode == 0
        assert "[FOUND]" in r_get.stdout
        assert username in r_get.stdout
        assert "password=***" in r_get.stdout
        assert_no_cleartext_password(r_get, password, "get")

        # 5. Update credentials -> Should exit 0
        new_password = "Updated_N3w_P@ssw0rd$!"
        r_update = run_cli_subprocess(
            ["--update-credentials", username, new_password, "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_update.returncode == 0
        assert "[SUCCESS]" in r_update.stdout
        assert_no_cleartext_password(r_update, new_password, "update")
        assert_no_cleartext_password(r_update, password, "update-old")

        # 6. Delete credentials -> Should exit 0
        r_del = run_cli_subprocess(
            ["--delete-credentials", "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_del.returncode == 0
        assert "[DELETED]" in r_del.stdout
        assert test_service in r_del.stdout
        assert_no_cleartext_password(r_del, new_password, "delete")

        # 7. Post-delete check -> Should exit 1
        r_post = run_cli_subprocess(
            ["--check-credentials", "--service", test_service],
            env_overrides=isolated_env,
        )
        assert r_post.returncode == 1
        assert "[NOT FOUND]" in r_post.stderr


class TestCleartextLeakageAdversarial:
    """Stress-tests password masking and leak prevention against edge-case passwords."""

    @pytest.fixture
    def isolated_env(self, tmp_path: Path):
        app_data = tmp_path / "appdata"
        app_data.mkdir(parents=True, exist_ok=True)
        return {"LOCALAPPDATA": str(app_data)}

    @pytest.mark.parametrize(
        "password_desc, password",
        [
            ("special_characters", "P@$$w0rd!#%^&*()_+{}:\"<>?~`-=[]\\;',./"),
            ("unicode_japanese_vietnamese", "MậtKhẩu_日本語_2026!🔑🔒"),
            ("whitespace_and_tabs", "  spaced   password   with \t tabs  "),
            ("json_injection_payload", '{"admin": true, "password": "hacked"}'),
            ("sql_injection_payload", "' OR '1'='1' -- ; DROP TABLE users;"),
            ("extremely_long_password", "A" * 4096),
        ],
    )
    def test_no_leakage_across_extreme_passwords(
        self,
        isolated_env: dict[str, str],
        password_desc: str,
        password: str,
    ):
        """Verifies that no password variant ever leaks into stdout or stderr."""
        service = f"ADV_LEAK_{uuid.uuid4().hex[:8]}"
        username = f"user_{password_desc[:10]}"

        # 1. Save
        r_save = run_cli_subprocess(
            ["--save-credentials", username, password, "--service", service],
            env_overrides=isolated_env,
        )
        assert r_save.returncode == 0
        assert_no_cleartext_password(r_save, password, f"save-{password_desc}")

        # 2. Get
        r_get = run_cli_subprocess(
            ["--get-credentials", "--service", service],
            env_overrides=isolated_env,
        )
        assert r_get.returncode == 0
        assert "[FOUND]" in r_get.stdout
        assert "password=***" in r_get.stdout
        assert_no_cleartext_password(r_get, password, f"get-{password_desc}")

        # 3. Clean up
        r_del = run_cli_subprocess(
            ["--delete-credentials", "--service", service],
            env_overrides=isolated_env,
        )
        assert r_del.returncode == 0
        assert_no_cleartext_password(r_del, password, f"del-{password_desc}")


class TestDPAPIFailureFallbackToKeyring:
    """Adversarially simulates DPAPI failures and verifies transparent keyring fallback."""

    @pytest.fixture
    def isolated_env(self, tmp_path: Path):
        app_data = tmp_path / "appdata"
        app_data.mkdir(parents=True, exist_ok=True)
        return {"LOCALAPPDATA": str(app_data)}

    def test_subprocess_fallback_when_dpapi_protect_fails(
        self,
        isolated_env: dict[str, str],
    ):
        """Simulates native CryptProtectData failure during save.

        Expected:
        - Primary DPAPI save raises DPAPIError.
        - CredentialManager catches it, logs warning, falls back to KeyringStorage.
        - KeyringStorage persists credentials into OS Keyring.
        - Subprocess returns exit code 0 and prints [SUCCESS].
        - STDOUT/STDERR contain ZERO cleartext password.
        """
        service = f"ADV_FALLBACK_SAVE_{uuid.uuid4().hex[:8]}"
        username = "fallback_hero"
        password = "FallbackSecret#2026!"

        # 1. Save with DPAPI protect failure simulation
        r_save = run_cli_subprocess(
            ["--save-credentials", username, password, "--service", service],
            env_overrides=isolated_env,
            simulate_dpapi_failure=True,
        )
        assert r_save.returncode == 0, f"Failed: {r_save.stderr}"
        assert "[SUCCESS]" in r_save.stdout
        assert_no_cleartext_password(r_save, password, "dpapi-failure-save")

        # Verify that no .dpapi file was created on disk due to the failure
        expected_dpapi_file = Path(isolated_env["LOCALAPPDATA"]) / "pm_sosanhbom" / "credentials" / f"{service}.dpapi"
        assert not expected_dpapi_file.exists(), "DPAPI file should not exist after DPAPI save failure"

        # 2. Check credentials via subprocess (with DPAPI still simulated as failed)
        r_check = run_cli_subprocess(
            ["--check-credentials", "--service", service],
            env_overrides=isolated_env,
            simulate_dpapi_failure=True,
        )
        assert r_check.returncode == 0
        assert "[EXISTS]" in r_check.stdout
        assert username in r_check.stdout
        assert_no_cleartext_password(r_check, password, "dpapi-failure-check")

        # 3. Get credentials via subprocess
        r_get = run_cli_subprocess(
            ["--get-credentials", "--service", service],
            env_overrides=isolated_env,
            simulate_dpapi_failure=True,
        )
        assert r_get.returncode == 0
        assert "[FOUND]" in r_get.stdout
        assert "password=***" in r_get.stdout
        assert_no_cleartext_password(r_get, password, "dpapi-failure-get")

        # 4. Delete credentials
        r_del = run_cli_subprocess(
            ["--delete-credentials", "--service", service],
            env_overrides=isolated_env,
            simulate_dpapi_failure=True,
        )
        assert r_del.returncode == 0
        assert "[DELETED]" in r_del.stdout
        assert_no_cleartext_password(r_del, password, "dpapi-failure-delete")

        # 5. Confirm deleted
        r_post = run_cli_subprocess(
            ["--get-credentials", "--service", service],
            env_overrides=isolated_env,
            simulate_dpapi_failure=True,
        )
        assert r_post.returncode == 1
        assert "[NOT FOUND]" in r_post.stderr

    def test_subprocess_fallback_when_platform_non_windows(
        self,
        isolated_env: dict[str, str],
    ):
        """Simulates IS_WINDOWS=False where DPAPI is completely disabled.

        Verifies that CredentialManager transparently utilizes KeyringStorage end-to-end.
        """
        service = f"ADV_NON_WIN_{uuid.uuid4().hex[:8]}"
        username = "linux_or_container_user"
        password = "CrossPlatformPass#999"

        # 1. Save
        r_save = run_cli_subprocess(
            ["--save-credentials", username, password, "--service", service],
            env_overrides=isolated_env,
            simulate_non_windows=True,
        )
        assert r_save.returncode == 0
        assert "[SUCCESS]" in r_save.stdout
        assert_no_cleartext_password(r_save, password, "non-windows-save")

        # 2. Get
        r_get = run_cli_subprocess(
            ["--get-credentials", "--service", service],
            env_overrides=isolated_env,
            simulate_non_windows=True,
        )
        assert r_get.returncode == 0
        assert "[FOUND]" in r_get.stdout
        assert username in r_get.stdout
        assert "password=***" in r_get.stdout
        assert_no_cleartext_password(r_get, password, "non-windows-get")

        # 3. Delete
        r_del = run_cli_subprocess(
            ["--delete-credentials", "--service", service],
            env_overrides=isolated_env,
            simulate_non_windows=True,
        )
        assert r_del.returncode == 0
        assert "[DELETED]" in r_del.stdout

    def test_subprocess_fallback_when_dpapi_file_corrupted(
        self,
        isolated_env: dict[str, str],
    ):
        """Simulates scenario where DPAPI file was corrupted, but credentials exist in Keyring."""
        service = f"ADV_CORRUPT_DPAPI_{uuid.uuid4().hex[:8]}"
        username = "dual_stored_user"
        password = "DualStoredPassword123"

        # First save to Keyring by running with simulate_non_windows=True
        r_save_kr = run_cli_subprocess(
            ["--save-credentials", username, password, "--service", service],
            env_overrides=isolated_env,
            simulate_non_windows=True,
        )
        assert r_save_kr.returncode == 0

        # Now deliberately place a corrupted .dpapi file in the DPAPI storage dir
        dpapi_dir = Path(isolated_env["LOCALAPPDATA"]) / "pm_sosanhbom" / "credentials"
        dpapi_dir.mkdir(parents=True, exist_ok=True)
        corrupted_file = dpapi_dir / f"{service}.dpapi"
        corrupted_file.write_bytes(b"MALFORMED_NON_DPAPI_GARBAGE_BYTES_XYZ_123456789")

        # Retrieve in normal mode (DPAPI is active but file is corrupted)
        # Primary DPAPI will fail to decrypt, return None, and fall back to Keyring!
        r_get = run_cli_subprocess(
            ["--get-credentials", "--service", service],
            env_overrides=isolated_env,
        )
        assert r_get.returncode == 0
        assert "[FOUND]" in r_get.stdout
        assert username in r_get.stdout
        assert "password=***" in r_get.stdout
        assert_no_cleartext_password(r_get, password, "corrupted-dpapi-fallback")

        # Clean up
        run_cli_subprocess(
            ["--delete-credentials", "--service", service],
            env_overrides=isolated_env,
        )


class TestCLIArgumentRobustness:
    """Verifies fail-closed behavior on malformed CLI invocations."""

    def test_cli_mutually_exclusive_flags_exit_code_2(self):
        """Passing multiple action flags must fail with exit code 2 (argparse error)."""
        r = run_cli_subprocess(["--save-credentials", "u", "p", "--get-credentials"])
        assert r.returncode == 2
        assert "not allowed with argument" in r.stderr

    def test_cli_missing_password_argument_exit_code_2(self):
        """--save-credentials requires 2 arguments; passing 1 must fail with exit code 2."""
        r = run_cli_subprocess(["--save-credentials", "only_user"])
        assert r.returncode == 2
        assert "expected 2 arguments" in r.stderr

    def test_cli_empty_credentials_exit_code_1(self, tmp_path: Path):
        """Passing empty strings as username or password should fail with code 1."""
        env = {"LOCALAPPDATA": str(tmp_path)}
        r = run_cli_subprocess(
            ["--save-credentials", "", "", "--service", "EMPTY_TEST"],
            env_overrides=env,
        )
        assert r.returncode == 1
        assert "[ERROR]" in r.stderr
        assert "must not be empty" in r.stderr

    def test_cli_no_args_prints_help_exit_code_0(self):
        """Calling CLI with no arguments should display help and exit code 0."""
        r = run_cli_subprocess([])
        assert r.returncode == 0
        assert "usage: python -m src.security.credentials" in r.stdout


class TestKeyringStorageRealWorldDefect:
    """Direct empirical probe on KeyringStorage retrieval bug under Windows WinVaultKeyring."""

    def test_keyring_storage_get_returns_real_user_and_password(self):
        """Probes that KeyringStorage.get returns actual username and password, not __current_user__."""
        from src.security.credentials import KeyringStorage

        storage = KeyringStorage()
        if not storage.is_available():
            pytest.skip("Keyring not available on this host")

        svc = f"ADV_PROBE_{uuid.uuid4().hex[:8]}"
        expected_user = "admin_technician"
        expected_pass = "ActualSuperSecretPass123"

        try:
            assert storage.save(svc, expected_user, expected_pass) is True
            creds = storage.get(svc)
            assert creds is not None, "Failed to retrieve saved credentials from Keyring"
            assert creds.username == expected_user, (
                f"BUG: KeyringStorage.get returned username={creds.username!r} instead of {expected_user!r}! "
                f"Raw credentials object: {creds!r}"
            )
            assert creds.password == expected_pass, (
                f"BUG: KeyringStorage.get returned password={creds.password!r} instead of {expected_pass!r}!"
            )
        finally:
            storage.delete(svc)


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v", "-s"]))
