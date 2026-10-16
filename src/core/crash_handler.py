"""SSBOM Application Crash Handler & Diagnostic Logging Subsystem.

Provides enterprise-grade fault protection and diagnostic logging:
- faulthandler integration capturing C-level faults (SIGSEGV, access violations 0xC0000005)
  into a persistent logs/crash_fault.log.
- sys.excepthook intercepting unhandled Python exceptions on the main thread, generating
  a detailed diagnostic report (logs/crash_YYYYMMDD_HHMMSS.log), flushing logs, and alerting
  the user via a modal GUI dialog rather than silently disappearing.
- threading.excepthook (Python 3.8+) preventing silent background worker thread death.
- Qt message interception via qInstallMessageHandler routing Qt warnings, errors, and fatal
  events into the Python logging pipeline.
- Automatic RotatingFileHandler (logs/app.log) ensuring all activities and errors are
  persisted to disk.
- Built-in utilities to open log directory and retrieve crash dumps for rapid support.
"""

from __future__ import annotations

import datetime
import faulthandler
import logging
import logging.handlers
import os
from pathlib import Path
import platform
import subprocess
import sys
import threading
import traceback
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)

_fault_file_handle: Optional[Any] = None
_crash_logs_dir: Optional[Path] = None
_in_exception_hook = False
_original_excepthook = sys.excepthook
_original_threading_excepthook = getattr(threading, "excepthook", None)


def resolve_logs_directory(base_dir: Optional[Path] = None) -> Path:
    """Resolve and ensure the application log directory exists with safe fallbacks."""
    global _crash_logs_dir
    if base_dir:
        cand = Path(base_dir) / "logs"
        try:
            cand.mkdir(parents=True, exist_ok=True)
            test_file = cand / ".write_test"
            test_file.write_text("ok", encoding="utf-8")
            test_file.unlink(missing_ok=True)
            _crash_logs_dir = cand.resolve()
            return _crash_logs_dir
        except Exception:
            pass

    if _crash_logs_dir and _crash_logs_dir.is_dir():
        return _crash_logs_dir

    candidates: list[Path] = []
    try:
        from src.core.path_utils import get_app_base_dir
        app_base = get_app_base_dir()
        candidates.append(app_base / "logs")
    except Exception:
        pass

    candidates.append(Path.cwd() / "logs")
    candidates.append(Path.home() / ".ssbom" / "logs")
    
    import tempfile
    candidates.append(Path(tempfile.gettempdir()) / "ssbom_logs")

    for cand in candidates:
        try:
            cand.mkdir(parents=True, exist_ok=True)
            # Test write access
            test_file = cand / ".write_test"
            test_file.write_text("ok", encoding="utf-8")
            test_file.unlink(missing_ok=True)
            _crash_logs_dir = cand.resolve()
            return _crash_logs_dir
        except Exception:
            continue

    fallback = Path.cwd()
    _crash_logs_dir = fallback
    return fallback


def get_crash_logs_dir() -> Path:
    """Return current logs directory."""
    return resolve_logs_directory()


def open_logs_folder() -> bool:
    """Open the crash logs directory in Windows Explorer or default file manager."""
    log_dir = get_crash_logs_dir()
    try:
        if sys.platform == "win32":
            os.startfile(str(log_dir))  # type: ignore[attr-defined]
            return True
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(log_dir)])
            return True
        else:
            subprocess.Popen(["xdg-open", str(log_dir)])
            return True
    except Exception as exc:
        logger.error("Could not open logs folder '%s': %s", log_dir, exc)
        return False


def _collect_system_diagnostics() -> dict[str, Any]:
    """Collect diagnostic information for crash reports."""
    diagnostics: dict[str, Any] = {
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "platform": platform.platform(),
        "python_version": sys.version,
        "python_executable": sys.executable,
        "working_directory": str(Path.cwd()),
        "active_threads": [t.name for t in threading.enumerate()],
    }

    try:
        from src.services.app_updates import AppUpdateManager
        diagnostics["app_version"] = AppUpdateManager().get_active_version()
    except Exception:
        diagnostics["app_version"] = "unknown"

    # Memory info if available
    try:
        import ctypes
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong),
                ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("sllAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        stat = MEMORYSTATUSEX()
        stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):  # type: ignore[attr-defined]
            diagnostics["memory_load_percent"] = stat.dwMemoryLoad
            diagnostics["avail_phys_mb"] = round(stat.ullAvailPhys / (1024 * 1024), 1)
            diagnostics["total_phys_mb"] = round(stat.ullTotalPhys / (1024 * 1024), 1)
    except Exception:
        pass

    return diagnostics


def _write_crash_report(
    error_type: str,
    exc_type: type,
    exc_value: BaseException,
    tb: Any,
    thread_name: Optional[str] = None,
) -> Path:
    """Format and write a complete crash diagnostic report to disk."""
    log_dir = get_crash_logs_dir()
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = f"crash_{thread_name}_" if thread_name else "crash_"
    report_path = log_dir / f"{prefix}{ts}.log"

    diag = _collect_system_diagnostics()
    tb_formatted = "".join(traceback.format_exception(exc_type, exc_value, tb))

    content = [
        "=" * 80,
        f"SSBOM MANAGER CRASH DIAGNOSTIC REPORT - {error_type}",
        "=" * 80,
        f"Timestamp          : {diag.get('timestamp')}",
        f"Application Version: {diag.get('app_version')}",
        f"Thread             : {thread_name or 'MainThread'}",
        f"Python Version     : {diag.get('python_version')}",
        f"Platform           : {diag.get('platform')}",
        f"Working Directory  : {diag.get('working_directory')}",
        f"Active Threads     : {', '.join(diag.get('active_threads', []))}",
    ]

    if "memory_load_percent" in diag:
        content.append(
            f"Memory Status      : {diag['memory_load_percent']}% used "
            f"({diag.get('avail_phys_mb')} MB free / {diag.get('total_phys_mb')} MB total)"
        )

    content.extend([
        "-" * 80,
        "EXCEPTION DETAILS:",
        f"Exception Type     : {exc_type.__module__}.{exc_type.__name__}",
        f"Exception Value    : {exc_value}",
        "-" * 80,
        "STACK TRACE:",
        tb_formatted,
        "=" * 80,
    ])

    report_text = "\n".join(content)
    try:
        report_path.write_text(report_text, encoding="utf-8")
    except Exception as write_err:
        print(f"[CRASH_HANDLER ERROR] Could not write report to {report_path}: {write_err}", file=sys.stderr)

    return report_path


def _show_gui_crash_dialog(report_file: Path, exc_value: BaseException) -> None:
    """Display user-facing crash warning dialog if QApplication is active and on the main thread."""
    if threading.current_thread() != threading.main_thread():
        # Avoid blocking or deadlocking Qt by creating widgets from background threads
        return

    try:
        from PyQt6.QtWidgets import QApplication, QMessageBox, QPushButton
        app = QApplication.instance()
        if not app:
            return

        box = QMessageBox()
        box.setWindowTitle("🚨 Sự Cố Ứng Dụng (Application Crash)")
        box.setIcon(QMessageBox.Icon.Critical)
        
        msg_text = (
            "<h3>Đã xảy ra sự cố không mong muốn (Crash)!</h3>"
            "<p>Chương trình gặp phải lỗi chưa được xử lý và cần ghi nhận chẩn đoán.</p>"
            f"<p><b>Chi tiết lỗi:</b> <code style='color:#DC2626;'>{exc_value}</code></p>"
            f"<p>📄 <b>Tệp nhật ký đã được lưu tại:</b><br/>"
            f"<a href='file:///{report_file.as_posix()}'>{report_file.name}</a></p>"
            "<p>Vui lòng gửi tệp này cho quản trị viên/kỹ sư phần mềm để được hỗ trợ khắc phục.</p>"
        )
        box.setText(msg_text)

        btn_open_folder = box.addButton("📂 Mở Thư Mục Logs", QMessageBox.ButtonRole.ActionRole)
        btn_copy = box.addButton("📋 Sao Chép Đường Dẫn", QMessageBox.ButtonRole.ActionRole)
        box.addButton("Đóng", QMessageBox.ButtonRole.AcceptRole)

        box.exec()

        if box.clickedButton() == btn_open_folder:
            open_logs_folder()
        elif box.clickedButton() == btn_copy:
            clipboard = app.clipboard()
            if clipboard:
                clipboard.setText(str(report_file))
    except Exception as ex:
        print(f"[CRASH_HANDLER] Could not show GUI crash dialog: {ex}", file=sys.stderr)


def ssbom_excepthook(
    exc_type: type[BaseException],
    exc_value: BaseException,
    tb: Any,
) -> None:
    """Global sys.excepthook catching unhandled Python exceptions on the main thread."""
    global _in_exception_hook
    if _in_exception_hook:
        # Prevent recursion if crash happens inside handler
        _original_excepthook(exc_type, exc_value, tb)
        return

    _in_exception_hook = True
    try:
        if issubclass(exc_type, KeyboardInterrupt):
            _original_excepthook(exc_type, exc_value, tb)
            return

        report_file = _write_crash_report("MAIN_THREAD_EXCEPTION", exc_type, exc_value, tb)
        logger.critical(
            "FATAL: Unhandled main thread exception! Diagnostic report generated: %s",
            report_file,
            exc_info=(exc_type, exc_value, tb),
        )

        # Flush all handlers to disk
        for h in logging.getLogger().handlers:
            try:
                h.flush()
            except Exception:
                pass

        _show_gui_crash_dialog(report_file, exc_value)
    except Exception as hook_err:
        print(f"[CRASH_HANDLER] Error inside ssbom_excepthook: {hook_err}", file=sys.stderr)
    finally:
        _in_exception_hook = False
        _original_excepthook(exc_type, exc_value, tb)


def ssbom_threading_excepthook(args: Any) -> None:
    """Global threading.excepthook catching unhandled exceptions in background threads."""
    if issubclass(args.exc_type, KeyboardInterrupt):
        return

    thread_name = getattr(args.thread, "name", "WorkerThread")
    try:
        report_file = _write_crash_report(
            "BACKGROUND_THREAD_EXCEPTION",
            args.exc_type,
            args.exc_value,
            args.exc_traceback,
            thread_name=thread_name,
        )
        logger.error(
            "Unhandled exception in background thread '%s'! Crash report: %s",
            thread_name,
            report_file,
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )
        for h in logging.getLogger().handlers:
            try:
                h.flush()
            except Exception:
                pass
    except Exception as th_err:
        print(f"[CRASH_HANDLER] Error inside threading excepthook: {th_err}", file=sys.stderr)


def install_qt_message_handler() -> None:
    """Install custom Qt message handler to capture Qt C++ warnings/fatals into logger."""
    try:
        from PyQt6.QtCore import QtMsgType, qInstallMessageHandler

        def _qt_handler(msg_type: QtMsgType, context: Any, message: str) -> None:
            if msg_type == QtMsgType.QtFatalMsg:
                logger.critical("[Qt Fatal] %s (file: %s, line: %s)", message, getattr(context, "file", ""), getattr(context, "line", ""))
                log_dir = get_crash_logs_dir()
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                fat_file = log_dir / f"crash_qt_fatal_{ts}.log"
                fat_file.write_text(f"QT FATAL MESSAGE:\n{message}\nFile: {getattr(context, 'file', '')}:{getattr(context, 'line', '')}", encoding="utf-8")
            elif msg_type == QtMsgType.QtCriticalMsg:
                logger.error("[Qt Critical] %s", message)
            elif msg_type == QtMsgType.QtWarningMsg:
                # Filter out spammy Qt font or platform theme warnings
                if "QWindowsWindow" not in message and "DirectWrite" not in message:
                    logger.warning("[Qt Warning] %s", message)
            else:
                logger.debug("[Qt Info] %s", message)

        qInstallMessageHandler(_qt_handler)
    except Exception as ex:
        logger.debug("Could not install Qt message handler: %s", ex)


def setup_crash_protection(base_dir: Optional[Path] = None) -> Path:
    """Initialize full crash protection and diagnostic logging.

    Returns:
        Resolved logs directory path.
    """
    global _fault_file_handle
    log_dir = resolve_logs_directory(base_dir)

    # 1. Enable faulthandler for low-level OS crashes
    fault_log_path = log_dir / "crash_fault.log"
    try:
        if _fault_file_handle is None or _fault_file_handle.closed:
            _fault_file_handle = open(fault_log_path, "a", encoding="utf-8", buffering=1)
            faulthandler.enable(file=_fault_file_handle, all_threads=True)
    except Exception as fh_exc:
        logger.warning("Could not enable faulthandler to file: %s", fh_exc)
        faulthandler.enable(all_threads=True)

    # 2. Attach RotatingFileHandler to root logger if not already attached
    root_logger = logging.getLogger()
    has_file_handler = any(isinstance(h, logging.handlers.RotatingFileHandler) for h in root_logger.handlers)
    if not has_file_handler:
        try:
            app_log_path = log_dir / "app.log"
            rf_handler = logging.handlers.RotatingFileHandler(
                app_log_path,
                maxBytes=10 * 1024 * 1024,  # 10 MB
                backupCount=5,
                encoding="utf-8",
            )
            rf_handler.setFormatter(
                logging.Formatter("[%(asctime)s] [%(levelname)s] [%(name)s]: %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
            )
            rf_handler.setLevel(logging.INFO)
            root_logger.addHandler(rf_handler)
            logger.info("Attached persistent file logger: %s", app_log_path)
        except Exception as rf_exc:
            print(f"[CRASH_HANDLER] Could not attach file logger: {rf_exc}", file=sys.stderr)

    # 3. Register global exception hooks
    sys.excepthook = ssbom_excepthook
    if hasattr(threading, "excepthook"):
        threading.excepthook = ssbom_threading_excepthook

    # 4. Register Qt message handler if Qt is imported or available
    install_qt_message_handler()

    return log_dir
