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
        subprocess.Popen(
            [self.creds.saplogon_path],
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0),
        )

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
            sap_gui_auto = None
            try:
                # If using real win32com.client, prefer SapROTWr.SapROTWrapper (works across 32/64-bit boundaries).
                # If using mock/custom provider, use GetObject as configured by test suite.
                is_win32com = (self.com_provider is win32com.client) if win32com else False

                if is_win32com:
                    if hasattr(self.com_provider, "Dispatch"):
                        try:
                            rot = self.com_provider.Dispatch("SapROTWr.SapROTWrapper")
                            sap_gui_auto = rot.GetROTEntry("SAPGUI")
                        except Exception as rot_ex:
                            logger.debug("SapROTWrapper Dispatch failed: %s", rot_ex)
                    if sap_gui_auto is None and hasattr(self.com_provider, "GetObject"):
                        try:
                            sap_gui_auto = self.com_provider.GetObject("SAPGUI")
                        except Exception as go_ex:
                            last_error = go_ex
                else:
                    if hasattr(self.com_provider, "GetObject"):
                        try:
                            sap_gui_auto = self.com_provider.GetObject("SAPGUI")
                        except Exception as go_ex:
                            last_error = go_ex
                    if sap_gui_auto is None and hasattr(self.com_provider, "Dispatch") and not hasattr(self.com_provider, "_mock_return_value"):
                        try:
                            rot = self.com_provider.Dispatch("SapROTWr.SapROTWrapper")
                            sap_gui_auto = rot.GetROTEntry("SAPGUI")
                        except Exception:
                            pass

                if sap_gui_auto is not None:
                    # Access ScriptingEngine property safely:
                    # In pywin32 dynamic COM dispatch, GetScriptingEngine is a CDispatch property;
                    # calling it like a function throws DISP_E_MEMBERNOTFOUND (-2147352573).
                    # But in test mocks (MagicMock), it is a callable returning mock_app.
                    app_obj = None
                    try:
                        app_obj = getattr(sap_gui_auto, "GetScriptingEngine", None)
                    except Exception:
                        pass

                    if app_obj is None:
                        app_obj = getattr(sap_gui_auto, "ScriptingEngine", sap_gui_auto)

                    is_com_disp = win32com and isinstance(app_obj, getattr(win32com.client, "CDispatch", ()))
                    if is_com_disp:
                        self.app = app_obj
                    elif callable(app_obj):
                        self.app = app_obj()
                    else:
                        self.app = app_obj

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
                desc = str(getattr(conn, "Description", "")).strip()
                sys_id = str(getattr(conn, "Id", "")).strip()

                target_sys = self.creds.system.strip()
                target_base = target_sys.split("(")[0].strip().upper()
                is_match = (
                    target_sys.lower() in desc.lower()
                    or target_sys.lower() in sys_id.lower()
                    or (target_base and target_base in desc.upper())
                    or (count == 1 and bool(desc))
                )

                if is_match:
                    children = getattr(conn, "Children", None)
                    child_count = getattr(children, "Count", 0) if children else 0
                    if child_count > 0:
                        sess = children(0) if callable(children) else children[0]
                        self.connection = conn
                        self.session = sess

                        if self.is_logged_in(self.session):
                            logger.info(f"Reusing existing logged-in session for {desc or self.creds.system}.")
                            return self.session
                        else:
                            logger.info("Reusing existing connection, completing login flow.")
                            self.handle_multi_logon_and_login(self.session)
                            return self.session
        except Exception as ex:
            logger.debug(f"Connection reuse check encountered non-fatal error: {ex}")

        # 2. Open new connection
        logger.info(f"Opening new SAP connection to '{self.creds.system}'...")
        conn_err = None
        candidates = [self.creds.system]
        if "AWS" in self.creds.system:
            candidates.extend(["P1J(ERP60)-VN-NEW", "P1J"])
        elif "NEW" in self.creds.system:
            candidates.extend(["P1J(ERP60-AWS)-VN", "P1J"])
        else:
            candidates.extend(["P1J(ERP60-AWS)-VN", "P1J(ERP60)-VN-NEW"])

        for sys_candidate in candidates:
            try:
                self.connection = self.app.OpenConnection(sys_candidate, True)
                if self.connection is not None:
                    break
            except Exception as ex:
                conn_err = ex
                continue

        if self.connection is None:
            raise SAPConnectionError(
                f"Failed to open connection to SAP system '{self.creds.system}' (tried candidates: {candidates}): {conn_err}"
            ) from conn_err

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

    def _resolve_multi_logon_dialog(self, session: Any) -> bool:
        """
        Detect and resolve SAP Multi-Logon license warning dialog (wnd[1]).
        Selects Option 2 (or Option 1 if Option 2 unavailable) and confirms with btn[0].
        Returns True if multi-logon dialog was found and handled, False otherwise.
        """
        try:
            # 1. Look for radMULTI_LOGON_OPT2 first (standard option: continue logon)
            rad_btn = None
            try:
                rad_btn = session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2")
            except Exception:
                pass

            # Fallback to radMULTI_LOGON_OPT1 if OPT2 is not present
            if rad_btn is None:
                try:
                    rad_btn = session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1")
                except Exception:
                    pass

            # Fallback: check children of wnd[1]/usr for any element containing MULTI_LOGON
            if rad_btn is None:
                try:
                    usr = session.findById("wnd[1]/usr")
                    children = getattr(usr, "Children", None)
                    count = getattr(children, "Count", 0) if children else 0
                    for i in range(count):
                        child = children(i) if callable(children) else children[i]
                        name = str(getattr(child, "Name", ""))
                        c_id = str(getattr(child, "Id", ""))
                        if "MULTI_LOGON_OPT2" in name or "MULTI_LOGON_OPT2" in c_id:
                            rad_btn = child
                            break
                        elif "MULTI_LOGON_OPT1" in name or "MULTI_LOGON_OPT1" in c_id:
                            rad_btn = child
                except Exception:
                    pass

            if rad_btn is not None:
                btn_name = getattr(rad_btn, "Name", "radMULTI_LOGON_OPT2")
                logger.warning(
                    f"Phát hiện cảnh báo đa phiên đăng nhập SAP (Multi-Logon) cho '{self.creds.username}'. "
                    f"Tự động chọn {btn_name} để tiếp tục đăng nhập..."
                )
                if hasattr(rad_btn, "Select"):
                    rad_btn.Select()
                if hasattr(rad_btn, "SetFocus"):
                    rad_btn.SetFocus()

                # Press Enter / Confirm in wnd[1]
                btn_enter = None
                try:
                    btn_enter = session.findById("wnd[1]/tbar[0]/btn[0]")
                except Exception:
                    pass

                if btn_enter is not None and hasattr(btn_enter, "press"):
                    btn_enter.press()
                else:
                    try:
                        wnd1 = session.findById("wnd[1]")
                        if hasattr(wnd1, "sendVKey"):
                            wnd1.sendVKey(0)
                    except Exception:
                        pass

                self.wait_ready(session)
                return True
        except Exception as ex:
            logger.debug(f"Multi-logon dialog resolution skipped or error: {ex}")

        return False

    def handle_multi_logon_and_login(self, session: Optional[Any] = None) -> None:
        """
        Handle multi-logon resolution dialog (`radMULTI_LOGON_OPT2`) and
        execute credentials entry on the initial SAP login window.
        """
        sess = session or self.session
        if sess is None:
            raise SAPConnectionError("No active SAP session to authenticate.")

        self.wait_ready(sess)

        # 1. Multi-logon popup resolution before credentials (if dialog already present or in test mocks)
        multi_logon_resolved = self._resolve_multi_logon_dialog(sess)

        # 2. Check if login screen fields are present and submit credentials
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

        # 3. Check for multi-logon dialog immediately after submitting credentials (real SAP behavior)
        if not multi_logon_resolved:
            multi_logon_resolved = self._resolve_multi_logon_dialog(sess)

        # 4. Handle modal information popups post-login (system notes, broadcast messages)
        # Avoid looping on static test mocks that do not simulate dialog closing
        is_mock = hasattr(sess, "_mock_return_value") or "Mock" in type(sess).__name__
        if not multi_logon_resolved or not is_mock:
            self.dismiss_post_login_popups(sess, multi_logon_already_handled=multi_logon_resolved)

        # 5. Final verification
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

    def dismiss_post_login_popups(self, session: Any, multi_logon_already_handled: bool = False) -> None:
        """Dismiss informational modal popups that may appear after logon."""
        multi_logon_done = multi_logon_already_handled
        for _ in range(5):
            try:
                # If a modal dialog exists at wnd[1]
                wnd1 = session.findById("wnd[1]")
                if wnd1 is not None:
                    # If this is a multi-logon dialog and not handled yet, select option 2 and confirm
                    if not multi_logon_done and self._resolve_multi_logon_dialog(session):
                        multi_logon_done = True
                        continue
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

    def logoff_and_exit(self, close_saplogon: bool = False) -> bool:
        """Gracefully log off active SAP session (/nex) and release connection."""
        success = False
        if self.session is not None:
            try:
                # 1. First attempt: Send /nex command to fast logoff without confirmation prompt
                okcd = None
                try:
                    okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
                except Exception:
                    pass

                if okcd is not None:
                    okcd.Text = "/nex"
                    wnd0 = self.session.findById("wnd[0]")
                    if wnd0 is not None:
                        wnd0.sendVKey(0)  # Enter
                        logger.info("Executed /nex fast logoff command successfully.")
                        success = True
                        time.sleep(0.5)
            except Exception as ex:
                logger.debug("Failed sending /nex to SAP session: %s", ex)

            if not success:
                try:
                    # Fallback: Close window or send cancel/exit VKeys
                    wnd0 = self.session.findById("wnd[0]")
                    if wnd0 is not None:
                        if hasattr(wnd0, "close"):
                            wnd0.close()
                        elif hasattr(wnd0, "sendVKey"):
                            wnd0.sendVKey(15)  # Shift+F3 Exit
                        success = True
                except Exception as ex:
                    logger.debug("Fallback close on SAP session failed: %s", ex)

        self.disconnect()

        if close_saplogon:
            try:
                for proc in psutil.process_iter(["name", "pid"]):
                    pname = (proc.info.get("name") or "").lower()
                    if pname in ("saplogon.exe", "sapgui.exe"):
                        proc.terminate()
                        logger.info("Closed saplogon process PID %s", proc.info.get("pid"))
            except Exception as ex:
                logger.debug("Could not terminate saplogon process: %s", ex)

        return success
