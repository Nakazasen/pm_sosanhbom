"""
SAP R3 GUI Connection & Session Management via COM Interoperability.
Handles SAP Logon process discovery, ROT registration, connection acquisition,
multi-logon conflict resolution, and authenticated session establishment.
"""

import logging
import os
import subprocess
import time
from typing import Any, Callable, Optional

import psutil

from src.automation.sap.models import (
    SAPConnectionError,
    SAPCredentials,
)

logger = logging.getLogger(__name__)

# Fallback import for win32com
try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # type: ignore


class SAPConnectionManager:
    """Manages lifecycle of SAP GUI 770 COM scripting sessions."""

    def __init__(
        self,
        credentials: Optional[SAPCredentials] = None,
        com_provider: Optional[Any] = None,
        process_checker: Optional[Callable[[], bool]] = None,
        process_launcher: Optional[Callable[[str], None]] = None,
        rot_getter: Optional[Callable[[], Any]] = None,
    ):
        """
        Initialize SAPConnectionManager.

        Args:
            credentials: SAP credentials and target system configuration.
            com_provider: Optional COM factory/mock providing `GetObject(prog_id)`.
            process_checker: Optional custom process checking callable for unit tests.
            process_launcher: Optional custom process launcher callable for unit tests.
            rot_getter: Optional custom ROT acquisition callable for unit tests.
        """
        self.creds = credentials or SAPCredentials()
        self.com_provider = com_provider or win32com.client if win32com else None
        self.process_checker = process_checker
        self.process_launcher = process_launcher
        self.rot_getter = rot_getter

        self.app = None
        self.connection = None
        self.session = None
        self._connected: bool = False

    @property
    def _session(self) -> Any:
        return self.session

    @_session.setter
    def _session(self, val: Any) -> None:
        self.session = val

    def is_saplogon_process_running(self) -> bool:
        """Check if saplogon.exe or sapgui.exe is running on the host system."""
        if self.process_checker is not None:
            return self.process_checker()

        target_names = {"saplogon.exe", "sapgui.exe", "saplogon", "sapgui"}
        for proc in psutil.process_iter(["name"]):
            try:
                proc_name = (proc.info.get("name") or "").lower()
                if proc_name in target_names:
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        return False

    def launch_saplogon(self) -> None:
        """Launch the saplogon.exe executable if not currently running."""
        if self.process_launcher is not None:
            self.process_launcher(self.creds.saplogon_path)
            return

        if not os.path.exists(self.creds.saplogon_path):
            raise SAPConnectionError(
                f"SAP Logon executable not found at specified path: '{self.creds.saplogon_path}'. "
                "Verify SAP GUI 770 installation or configure credentials.saplogon_path."
            )

        logger.info(f"Starting SAP Logon from {self.creds.saplogon_path}...")
        subprocess.Popen([self.creds.saplogon_path])

    def ensure_saplogon_running(self, timeout_sec: float = 15.0) -> None:
        """
        Ensure saplogon.exe is running and the SAPGUI COM scripting engine
        is registered in the Running Object Table (ROT).
        """
        if not self.is_saplogon_process_running():
            self.launch_saplogon()

        if self.rot_getter is not None:
            self.app = self.rot_getter()
            if self.app is not None:
                return
            raise SAPConnectionError(
                f"Timed out waiting for SAPGUI scripting engine registration in ROT after {timeout_sec}s."
            )

        if self.com_provider is None:
            raise SAPConnectionError(
                "win32com.client is unavailable. Ensure pywin32 is installed on Windows."
            )

        start_time = time.time()
        last_error = None

        while time.time() - start_time < timeout_sec:
            try:
                sap_gui_auto = self.com_provider.GetObject("SAPGUI")
                if sap_gui_auto is not None:
                    # SAP GUI Scripting engine hook
                    self.app = getattr(sap_gui_auto, "GetScriptingEngine", None)
                    if callable(self.app):
                        self.app = self.app()
                    elif self.app is None:
                        # Direct engine object fallback
                        self.app = getattr(sap_gui_auto, "ScriptingEngine", sap_gui_auto)

                    if self.app is not None:
                        logger.info("Successfully acquired SAPGUI ScriptingEngine.")
                        return
            except Exception as ex:
                last_error = ex
                time.sleep(0.5)

        raise SAPConnectionError(
            f"Timed out waiting for SAPGUI scripting engine registration in ROT after {timeout_sec}s. "
            f"Last COM error: {last_error}"
        )

    def get_or_create_session(self, timeout_sec: float = 30.0) -> Any:
        """
        Acquire an authenticated SAP GUI session:
        1. Checks for an existing active connection to `creds.system` to reuse.
        2. If none, opens a new connection.
        3. Resolves multi-logon conflicts and authenticates credentials.
        """
        self.ensure_saplogon_running()

        if self.app is None:
            raise SAPConnectionError("Scripting engine is uninitialized.")

        # 1. Attempt to reuse existing connection
        try:
            connections = getattr(self.app, "Connections", None)
            count = getattr(connections, "Count", 0) if connections else 0
            for i in range(count):
                conn = connections(i) if callable(connections) else connections[i]
                desc = str(getattr(conn, "Description", ""))
                sys_id = str(getattr(conn, "Id", ""))

                if self.creds.system in desc or self.creds.system in sys_id:
                    children = getattr(conn, "Children", None)
                    child_count = getattr(children, "Count", 0) if children else 0
                    if child_count > 0:
                        sess = children(0) if callable(children) else children[0]
                        self.connection = conn
                        self.session = sess

                        if self.is_logged_in(self.session):
                            logger.info(f"Reusing existing logged-in session for {self.creds.system}.")
                            return self.session
                        else:
                            logger.info("Reusing existing connection, completing login flow.")
                            self.handle_multi_logon_and_login(self.session)
                            return self.session
        except Exception as ex:
            logger.debug(f"Connection reuse check encountered non-fatal error: {ex}")

        # 2. Open new connection
        logger.info(f"Opening new SAP connection to '{self.creds.system}'...")
        try:
            self.connection = self.app.OpenConnection(self.creds.system, True)
        except Exception as ex:
            raise SAPConnectionError(
                f"Failed to open connection to SAP system '{self.creds.system}': {ex}"
            ) from ex

        self.wait_ready(self.connection, timeout_sec=timeout_sec)

        # 3. Retrieve default session (Child 0)
        try:
            children = getattr(self.connection, "Children", None)
            self.session = children(0) if callable(children) else children[0]
        except Exception as ex:
            raise SAPConnectionError(f"Failed to acquire initial session from connection: {ex}") from ex

        # 4. Resolve multi-logon & login
        self.handle_multi_logon_and_login(self.session)
        return self.session

    def handle_multi_logon_and_login(self, session: Optional[Any] = None) -> None:
        """
        Handle multi-logon resolution dialog (`radMULTI_LOGON_OPT2`) and
        execute credentials entry on the initial SAP login window.
        """
        sess = session or self.session
        if sess is None:
            raise SAPConnectionError("No active SAP session to authenticate.")

        self.wait_ready(sess)

        # 1. Multi-logon popup resolution:
        # Option 2 terminates existing logons and enters this one
        multi_logon_resolved = False
        try:
            rad_opt2 = sess.findById("wnd[1]/usr/radMULTI_LOGON_OPT2")
            if rad_opt2 is not None:
                logger.warning(
                    f"Multi-logon conflict detected for {self.creds.username}. "
                    "Selecting Option 2 (terminate older sessions)."
                )
                if hasattr(rad_opt2, "Select"):
                    rad_opt2.Select()
                if hasattr(rad_opt2, "SetFocus"):
                    rad_opt2.SetFocus()

                # Press Enter / Confirm in wnd[1]
                btn_enter = sess.findById("wnd[1]/tbar[0]/btn[0]")
                if btn_enter is not None:
                    btn_enter.press()
                self.wait_ready(sess)
                multi_logon_resolved = True
        except Exception:
            # No multi-logon dialog present
            pass

        # 2. Check if login screen fields are present
        try:
            bname_field = sess.findById("wnd[0]/usr/txtRSYST-BNAME")
            if bname_field is not None:
                logger.info(f"Authenticating user '{self.creds.username}' in SAP...")
                bname_field.Text = self.creds.username

                bcode_field = sess.findById("wnd[0]/usr/pwdRSYST-BCODE")
                if bcode_field is not None:
                    bcode_field.Text = self.creds.password

                langu_field = sess.findById("wnd[0]/usr/txtRSYST-LANGU")
                if langu_field is not None and self.creds.language:
                    langu_field.Text = self.creds.language

                wnd0 = sess.findById("wnd[0]")
                if wnd0 is not None:
                    wnd0.sendVKey(0)  # VKey 0 is Enter

                self.wait_ready(sess)
        except Exception as ex:
            logger.debug(f"Login field entry bypassed or non-existent: {ex}")

        # 3. Handle any modal information popups post-login (system notes, license notices)
        if not multi_logon_resolved:
            self.dismiss_post_login_popups(sess)

        # 4. Final verification
        if not self.is_logged_in(sess):
            # Check if an error was posted to sbar
            sbar_text = ""
            try:
                sbar = sess.findById("wnd[0]/sbar")
                if sbar and getattr(sbar, "MessageType", "") == "E":
                    sbar_text = getattr(sbar, "Text", "")
            except Exception:
                pass

            msg = f"Login to SAP system '{self.creds.system}' failed."
            if sbar_text:
                msg += f" Status bar: {sbar_text}"
            raise SAPConnectionError(msg)

        logger.info(f"Successfully authenticated to SAP system '{self.creds.system}'.")

    def dismiss_post_login_popups(self, session: Any) -> None:
        """Dismiss informational modal popups that may appear after logon."""
        for _ in range(3):
            try:
                # If a modal dialog exists at wnd[1]
                wnd1 = session.findById("wnd[1]")
                if wnd1 is not None:
                    btn0 = session.findById("wnd[1]/tbar[0]/btn[0]")
                    if btn0 is not None:
                        btn0.press()
                        self.wait_ready(session)
                    else:
                        break
                else:
                    break
            except Exception:
                break

    def is_logged_in(self, session: Optional[Any] = None) -> bool:
        """
        Verify whether the session is currently authenticated and ready for transactions
        by checking for the command line transaction code box ('okcd').
        """
        sess = session or self.session
        if sess is None:
            return False

        try:
            # okcd indicates standard SAP main menu / transaction availability
            okcd = sess.findById("wnd[0]/tbar[0]/okcd")
            return okcd is not None
        except Exception:
            return False

    def wait_ready(self, target_obj: Optional[Any] = None, timeout_sec: float = 30.0) -> None:
        """Wait until session / connection is not Busy."""
        obj = target_obj or self.session
        if obj is None:
            return

        start_time = time.time()
        while getattr(obj, "Busy", False) is True:
            time.sleep(0.1)
            if time.time() - start_time > timeout_sec:
                raise SAPConnectionError(f"SAP GUI session busy timeout after {timeout_sec}s.")

    def connect(self, timeout_sec: float = 0.5) -> Any:
        """Acquire connection to SAP GUI Scripting engine and return active session."""
        if self.com_provider is not None:
            try:
                sap_gui_auto = self.com_provider.GetObject("SAPGUI")
                if sap_gui_auto is not None:
                    self.app = getattr(sap_gui_auto, "GetScriptingEngine", None)
                    if callable(self.app):
                        self.app = self.app()
            except Exception as ex:
                raise SAPConnectionError(f"Failed to acquire SAPGUI COM engine: {ex}") from ex

        self.ensure_saplogon_running(timeout_sec=timeout_sec)
        sess = self.get_or_create_session(timeout_sec=timeout_sec)
        self._connected = True
        return sess

    def is_connected(self) -> bool:
        """Check if an active connection or session exists and is responsive."""
        if not self._connected:
            return False
        if self.session is None and self.connection is None:
            return False
        if self.session is not None:
            try:
                return bool(self.is_logged_in(self.session))
            except Exception:
                return False
        if self.connection is not None:
            try:
                _ = getattr(self.connection, "Children", None)
                return True
            except Exception:
                return False
        return False

    def get_session(self) -> Any:
        """Return existing healthy session or automatically reconnect if severed/stale."""
        if self.session is not None:
            try:
                if self.is_logged_in(self.session):
                    return self.session
                logger.warning("Existing SAP session is not logged in or unresponsive; re-acquiring...")
            except Exception as ex:
                logger.warning("Existing SAP session threw error on health-check (%s); reconnecting...", ex)
            self.session = None

        return self.get_or_create_session()

    def disconnect(self) -> None:
        """Release session and connection references cleanly."""
        for target in (self.session, getattr(self, "_session", None), self.connection):
            if target is not None and hasattr(target, "close"):
                try:
                    target.close()
                except Exception:
                    pass
        self.session = None
        self.connection = None
        self.app = None
        self._connected = False
        logger.info("SAP connection released.")
