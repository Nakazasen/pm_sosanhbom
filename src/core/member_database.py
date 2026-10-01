"""Centralized SQLite Member Management Database Module.

Provides shared database persistence and offline caching for Kyocera SSBOM:
- Remote primary DB (UNC): \\\\fstvn01\\Data\\00_KDTVN Common(KDTVN共通)\\⑤Production Engineering(製造技術)\\Hang muc can luu\\Vinh\\Pm_sosanhBOM\\ssbom_master.db
- Local offline cache fallback: <base_dir>/data/ssbom_master_cache.db
- Full CRUD operations for team personnel (engineers, departments, default sub-units).
- Automatic database migration and seeding with the 38 canonical engineers.
"""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import asdict, dataclass
import datetime
import json
import logging
import os
from pathlib import Path
import random
import re
import shutil
import sqlite3
import time
from typing import Any

import openpyxl

from src.core.path_utils import (
    get_app_base_dir,
    is_unc_path_accessible,
    safe_exists,
    safe_is_file,
)

logger = logging.getLogger(__name__)

# Default UNC Network path requested by User
DEFAULT_SHARED_DB_PATH = (
    r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)"
    r"\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db"
)

DEFAULT_MEMBER_EXCEL_PATH = (
    r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)"
    r"\Hang muc can luu\Vinh\Pm_sosanhBOM\Danhsachthanhvien.xlsx"
)

# Canonical 38 engineers from Sheet tenphong_pt
SEED_ROSTER_MECHA_1: list[str] = [
    "Son_mecha1", "Duy_mecha1", "KhiemA_mecha1", "KhiemE_mecha1",
    "Nghia_mecha1", "Quyen_mecha1", "Ly_mecha1", "Tuong_mecha1",
    "Q.Anh_mecha1", "Ngoc_mecha1", "Hung_mecha1", "Long_mecha1",
]

SEED_ROSTER_MECHA_2: list[str] = [
    "Loc_mecha2", "Hai_mecha2", "Dat_mecha2", "Nhung_mecha2",
    "Hung_mecha2", "Quynh_mecha2", "HuuTu_mecha2", "Tuan_mecha2",
    "XuanSon_mecha2", "Thuy_mecha2", "Canh_mecha2",
]

SEED_ROSTER_MECHA_3: list[str] = [
    "LVThuan_mecha3", "V.Thuong_mecha3", "Vinh_mecha3", "Viet_mecha3",
    "Nguyet_mecha3", "Thao_mecha3", "Trong_mecha3", "Minh_mecha3",
    "Khai_mecha3", "V.Thanh_mecha3", "Do_mecha3", "NThuan_mecha3",
    "Khang_mecha3", "HaiDang_mecha3", "KhacTu_mecha3",
]


@dataclass
class MemberRecord:
    """Represents an engineer / production team member."""

    account_id: str  # Unique account identifier (e.g. Son_mecha1)
    full_name: str  # Full name or display name
    department: str  # Department (e.g. Cơ 1, Cơ 2, Cơ 3)
    default_sub_unit: str = ""  # Default Sub-unit / Assy (e.g. DRUM, LSU, FUSER)
    machine_names: str = ""  # Dòng máy / Loại máy phụ trách (cách nhau dấu phẩy, e.g. 'Virgo, Iris 2024')
    is_active: bool = True  # Whether currently active in project
    notes: str = ""  # Optional notes
    id: int | None = None
    created_at: str | None = None
    updated_at: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class MemberDatabaseManager:
    """Manages member data operations with automatic LAN UNC sync and local caching."""

    def __init__(
        self,
        base_dir: Path | str | None = None,
        remote_db_path: Path | str | None = None,
        excel_path: Path | str | None = None,
    ) -> None:
        self.base_dir = Path(base_dir) if base_dir else get_app_base_dir()
        self._remote_check_cache: bool | None = None
        self._remote_check_time: float = 0.0

        cfg_path = self.base_dir / "config" / "settings.json"
        cfg_shared = ""
        cfg_excel = ""
        if safe_exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    cfg_shared = data.get("paths", {}).get("shared_db_path", "")
                    cfg_excel = data.get("paths", {}).get("member_excel_path", "")
            except Exception:
                pass

        is_test_dir = any(t in str(self.base_dir).lower() for t in ("pytest", "temp", "tmp"))
        if remote_db_path:
            self.remote_db_path = Path(remote_db_path)
        elif cfg_shared:
            self.remote_db_path = Path(cfg_shared)
        else:
            local_master = self.base_dir / "ssbom_master.db"
            if is_test_dir or safe_exists(local_master):
                self.remote_db_path = local_master
            else:
                self.remote_db_path = Path(DEFAULT_SHARED_DB_PATH)

        if excel_path:
            self.excel_path = Path(excel_path)
        elif cfg_excel:
            self.excel_path = Path(cfg_excel)
        else:
            local_xlsx = self.base_dir / "Danhsachthanhvien.xlsx"
            local_xlsm = self.base_dir / "Danhsachthanhvien.xlsm"
            if is_test_dir:
                self.excel_path = local_xlsx if safe_exists(local_xlsx) else (local_xlsm if safe_exists(local_xlsm) else local_xlsx)
            elif safe_exists(local_xlsx):
                self.excel_path = local_xlsx
            elif safe_exists(local_xlsm):
                self.excel_path = local_xlsm
            else:
                self.excel_path = Path(DEFAULT_MEMBER_EXCEL_PATH)

        self.is_test_dir = is_test_dir
        self._last_cache_sync_time: float = 0.0

        self.local_cache_dir = self.base_dir / "data"
        self.local_cache_dir.mkdir(parents=True, exist_ok=True)
        self.local_cache_path = self.local_cache_dir / "ssbom_master_cache.db"

        # Initialize schema and seed default data
        self.initialize_database()

    def is_remote_available(self) -> bool:
        """Check if remote UNC share or directory is accessible with caching."""
        now = time.time()
        if self._remote_check_cache is not None and (now - self._remote_check_time < 15.0):
            return self._remote_check_cache

        is_avail = False
        try:
            parent_dir = self.remote_db_path.parent
            if is_unc_path_accessible(parent_dir, timeout=0.5):
                is_avail = safe_exists(parent_dir)
        except Exception:
            is_avail = False

        self._remote_check_cache = is_avail
        self._remote_check_time = now
        return is_avail

    def _ensure_cache_fresh(self, force: bool = False) -> None:
        """Ensure local cache is up-to-date with remote DB using throttled mtime check."""
        now = time.time()
        throttle = 0.0 if self.is_test_dir else 5.0
        if not force and (now - self._last_cache_sync_time < throttle):
            return
        self._last_cache_sync_time = now

        if not self.is_remote_available():
            return

        remote_path = self.remote_db_path
        if remote_path == self.local_cache_path or not safe_exists(remote_path):
            return

        try:
            if not safe_exists(self.local_cache_path):
                self._safe_backup_db(remote_path, self.local_cache_path)
                return

            r_mtime = remote_path.stat().st_mtime
            l_mtime = self.local_cache_path.stat().st_mtime
            if r_mtime > l_mtime:
                logger.debug("Remote DB is newer (%.1f > %.1f), syncing cache...", r_mtime, l_mtime)
                self._safe_backup_db(remote_path, self.local_cache_path)
        except Exception as exc:
            logger.debug("Error checking cache freshness: %s", exc)

    def get_active_db_path(self) -> tuple[Path, bool]:
        """Determine whether to use remote network database or local cache.

        Returns:
            (active_path, is_remote_connected)
        """
        if self.is_remote_available():
            return self.remote_db_path, True
        return self.local_cache_path, False

    def get_read_db_path(self, force_refresh: bool = False) -> Path:
        """Return local cache path for local-first zero-latency read access."""
        self._ensure_cache_fresh(force=force_refresh)
        if safe_exists(self.local_cache_path):
            return self.local_cache_path
        active_path, _ = self.get_active_db_path()
        return active_path

    def sync_cache_from_remote(self, force: bool = False) -> bool:
        """Explicitly refresh local cache from remote DB if remote is available."""
        if not self.is_remote_available() or not safe_exists(self.remote_db_path):
            return False
        try:
            self._safe_backup_db(self.remote_db_path, self.local_cache_path)
            self._last_cache_sync_time = time.time()
            return True
        except Exception as exc:
            logger.error("Explicit cache sync failed: %s", exc)
            return False

    def get_member_excel_path(self) -> Path:
        """Resolve the active member Excel file path with fallback to candidate paths."""
        if safe_exists(self.excel_path):
            return self.excel_path
        # Check alternative extension (.xlsx vs .xlsm)
        alt_ext = ".xlsm" if self.excel_path.suffix.lower() == ".xlsx" else ".xlsx"
        alt_path = self.excel_path.with_suffix(alt_ext)
        if safe_exists(alt_path):
            return alt_path

        candidates = [
            self.base_dir / "Danhsachthanhvien.xlsx",
            self.base_dir / "Danhsachthanhvien.xlsm",
            self.base_dir / "data" / "Danhsachthanhvien.xlsx",
            self.base_dir / "data" / "Danhsachthanhvien.xlsm",
            self.base_dir / "config" / "Danhsachthanhvien.xlsx",
            self.base_dir / "config" / "Danhsachthanhvien.xlsm",
            Path("Danhsachthanhvien.xlsx"),
            Path("Danhsachthanhvien.xlsm"),
            Path("data/Danhsachthanhvien.xlsx"),
            Path("data/Danhsachthanhvien.xlsm"),
        ]
        for c in candidates:
            if safe_exists(c):
                return c
        return self.excel_path

    @staticmethod
    def _safe_backup_db(src_path: Path, dst_path: Path) -> None:
        """Safely copy live SQLite database using sqlite3 online backup API."""
        if not safe_exists(src_path):
            return
        src_conn = None
        dst_conn = None
        try:
            dst_path.parent.mkdir(parents=True, exist_ok=True)
            src_conn = sqlite3.connect(str(src_path), timeout=10.0)
            dst_conn = sqlite3.connect(str(dst_path), timeout=10.0)
            src_conn.backup(dst_conn)
        except Exception as exc:
            logger.debug("Safe SQLite backup skipped: %s", exc)
        finally:
            if src_conn is not None:
                try:
                    src_conn.close()
                except Exception:
                    pass
            if dst_conn is not None:
                try:
                    dst_conn.close()
                except Exception:
                    pass

    @contextmanager
    def _get_connection(self, db_path: Path):
        """Open an SQLite connection with 30s busy timeout and guaranteed clean closure."""
        db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(db_path), timeout=30.0)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA busy_timeout = 30000;")
            yield conn
        finally:
            try:
                conn.close()
            except Exception:
                pass

    @contextmanager
    def _write_transaction(self, db_path: Path):
        """Open an SQLite write transaction using BEGIN IMMEDIATE to avoid upgrade deadlocks."""
        db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(db_path), timeout=30.0, isolation_level=None)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA busy_timeout = 30000;")
            conn.execute("BEGIN IMMEDIATE;")
            yield conn
            conn.execute("COMMIT;")
        except Exception:
            try:
                conn.execute("ROLLBACK;")
            except Exception:
                pass
            raise
        finally:
            try:
                conn.close()
            except Exception:
                pass

    def _execute_dual_write(self, write_action, operation_name: str = "ghi dữ liệu") -> tuple[bool, Any, bool]:
        """Execute write_action(conn) on remote (with jitter retry) and immediately on local cache.

        write_action: callable(sqlite3.Connection) -> result (e.g. lastrowid, rowcount, or custom).
        Returns:
            (success: bool, result_or_error_msg: Any, used_fallback: bool)
        """
        remote_path, is_remote = self.get_active_db_path()
        used_fallback = False
        result = None

        # 1. Write to remote if available
        if is_remote and remote_path != self.local_cache_path:
            max_retries = 5
            remote_success = False
            for attempt in range(max_retries):
                try:
                    with self._write_transaction(remote_path) as conn:
                        result = write_action(conn)
                    remote_success = True
                    break
                except sqlite3.IntegrityError as int_err:
                    return False, str(int_err), False
                except LookupError as lk_err:
                    return False, str(lk_err), False
                except sqlite3.OperationalError as op_err:
                    err_str = str(op_err).lower()
                    if ("locked" in err_str or "busy" in err_str) and attempt < max_retries - 1:
                        delay = 0.04 * (1.5 ** attempt) + random.uniform(0.01, 0.06)
                        logger.warning(
                            "Remote DB locked during %s, retrying in %.3fs (attempt %d/%d)...",
                            operation_name,
                            delay,
                            attempt + 1,
                            max_retries,
                        )
                        time.sleep(delay)
                        continue
                    logger.warning("Remote DB write failed after retries: %s; falling back to local cache.", op_err)
                    used_fallback = True
                    break
                except Exception as ex:
                    logger.warning("Remote write error: %s; falling back to local cache.", ex)
                    used_fallback = True
                    break

        # 2. Dual-write directly to local cache
        try:
            with self._write_transaction(self.local_cache_path) as conn:
                local_res = write_action(conn)
                if result is None:
                    result = local_res
        except sqlite3.IntegrityError as int_err:
            return False, str(int_err), used_fallback
        except LookupError as lk_err:
            return False, str(lk_err), used_fallback
        except Exception as local_err:
            logger.error("Local cache dual-write error: %s", local_err)
            if used_fallback and result is None:
                return False, f"Lỗi ghi dữ liệu cục bộ: {local_err}", used_fallback

        self._last_cache_sync_time = time.time()
        return True, result, used_fallback

    def initialize_database(self) -> None:
        """Create tables if not existing, seed canonical engineers, and sync cache."""
        target_path, is_remote = self.get_active_db_path()

        try:
            with self._get_connection(target_path) as conn:
                is_network = str(target_path).startswith(r"\\") or str(target_path).startswith("//")
                try:
                    if is_network:
                        conn.execute("PRAGMA journal_mode = TRUNCATE;")
                    else:
                        conn.execute("PRAGMA journal_mode = WAL;")
                except sqlite3.OperationalError:
                    pass

                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='members';")
                table_exists = cursor.fetchone() is not None

                if not table_exists:
                    conn.execute(
                        """
                        CREATE TABLE IF NOT EXISTS members (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            account_id TEXT UNIQUE NOT NULL,
                            full_name TEXT NOT NULL,
                            department TEXT NOT NULL,
                            default_sub_unit TEXT DEFAULT '',
                            machine_names TEXT DEFAULT '',
                            is_active INTEGER DEFAULT 1,
                            notes TEXT DEFAULT '',
                            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                        );
                        """
                    )
                    conn.execute("CREATE INDEX IF NOT EXISTS idx_members_dept ON members(department);")
                    conn.execute("CREATE INDEX IF NOT EXISTS idx_members_active ON members(is_active);")

                    cursor.execute("SELECT COUNT(*) FROM members;")
                    count = cursor.fetchone()[0]

                    if count == 0:
                        logger.info("Seeding database '%s' with 38 initial engineers...", target_path)
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        seed_items: list[tuple[str, str, str, str, str, int, str, str, str]] = []
                        for name in SEED_ROSTER_MECHA_1:
                            seed_items.append((name, name, "Cơ 1", "", "(Tất cả)", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                        for name in SEED_ROSTER_MECHA_2:
                            seed_items.append((name, name, "Cơ 2", "", "(Tất cả)", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                        for name in SEED_ROSTER_MECHA_3:
                            seed_items.append((name, name, "Cơ 3", "", "(Tất cả)", 1, "Mặc định từ tenphong_pt", now_str, now_str))

                        conn.executemany(
                            """
                            INSERT OR IGNORE INTO members (
                                account_id, full_name, department, default_sub_unit, machine_names,
                                is_active, notes, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
                            """,
                            seed_items,
                        )
                        conn.commit()
                else:
                    # Auto-migration for existing databases
                    try:
                        cursor.execute("PRAGMA table_info(members);")
                        existing_cols = [r[1] for r in cursor.fetchall()]
                        if "machine_names" not in existing_cols:
                            conn.execute("ALTER TABLE members ADD COLUMN machine_names TEXT DEFAULT '';")
                            conn.commit()
                    except sqlite3.OperationalError:
                        pass

            # If remote was used and local cache is missing or older, update local cache
            if is_remote and safe_exists(target_path):
                should_sync = True
                if safe_exists(self.local_cache_path):
                    try:
                        should_sync = target_path.stat().st_mtime > self.local_cache_path.stat().st_mtime
                    except OSError:
                        should_sync = False
                if should_sync:
                    try:
                        self._safe_backup_db(target_path, self.local_cache_path)
                    except Exception as cache_err:
                        logger.debug("Failed to cache remote DB to local: %s", cache_err)

            # If only local was used but local didn't exist, it is now initialized
        except Exception as exc:
            logger.error("Error initializing member database: %s", exc)
            # Fallback to local cache initialization if remote failed during write
            if is_remote:
                logger.info("Retrying initialization on local cache...")
                try:
                    with self._get_connection(self.local_cache_path) as conn:
                        cursor = conn.cursor()
                        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='members';")
                        if not cursor.fetchone():
                            conn.execute(
                                """
                                CREATE TABLE IF NOT EXISTS members (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    account_id TEXT UNIQUE NOT NULL,
                                    full_name TEXT NOT NULL,
                                    department TEXT NOT NULL,
                                    default_sub_unit TEXT DEFAULT '',
                                    machine_names TEXT DEFAULT '',
                                    is_active INTEGER DEFAULT 1,
                                    notes TEXT DEFAULT '',
                                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                                );
                                """
                            )
                            cursor.execute("SELECT COUNT(*) FROM members;")
                            if cursor.fetchone()[0] == 0:
                                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                seed_items = []
                                for name in SEED_ROSTER_MECHA_1:
                                    seed_items.append((name, name, "Cơ 1", "", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                                for name in SEED_ROSTER_MECHA_2:
                                    seed_items.append((name, name, "Cơ 2", "", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                                for name in SEED_ROSTER_MECHA_3:
                                    seed_items.append((name, name, "Cơ 3", "", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                                conn.executemany(
                                    """
                                    INSERT OR IGNORE INTO members (
                                        account_id, full_name, department, default_sub_unit, machine_names,
                                        is_active, notes, created_at, updated_at
                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
                                    """,
                                    seed_items,
                                )
                                conn.commit()
                except Exception as local_err:
                    logger.error("Fatal: failed to initialize local cache DB: %s", local_err)

    def get_members(
        self,
        department: str | None = None,
        active_only: bool = False,
    ) -> list[MemberRecord]:
        """Query members filtered by department and active status (Local-First fast read)."""
        db_path = self.get_read_db_path()
        query = "SELECT * FROM members WHERE 1=1"
        params: list[Any] = []

        if department and department != "Tất cả":
            query += " AND department = ?"
            params.append(department)

        if active_only:
            query += " AND is_active = 1"

        query += " ORDER BY department ASC, account_id ASC;"

        results: list[MemberRecord] = []
        try:
            with self._get_connection(db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                for row in cursor.fetchall():
                    results.append(
                        MemberRecord(
                            id=row["id"],
                            account_id=row["account_id"],
                            full_name=row["full_name"],
                            department=row["department"],
                            default_sub_unit=row["default_sub_unit"] or "",
                            machine_names=row["machine_names"] if ("machine_names" in row.keys() and row["machine_names"]) else "",
                            is_active=bool(row["is_active"]),
                            notes=row["notes"] or "",
                            created_at=row["created_at"],
                            updated_at=row["updated_at"],
                        )
                    )
        except Exception as exc:
            logger.error("Error querying members from %s: %s", db_path, exc)
            # Fallback to local cache if error on remote
            if db_path != self.local_cache_path and self.local_cache_path.exists():
                try:
                    with self._get_connection(self.local_cache_path) as conn:
                        cursor = conn.cursor()
                        cursor.execute(query, params)
                        for row in cursor.fetchall():
                            results.append(
                                MemberRecord(
                                    id=row["id"],
                                    account_id=row["account_id"],
                                    full_name=row["full_name"],
                                    department=row["department"],
                                    default_sub_unit=row["default_sub_unit"] or "",
                                    machine_names=row["machine_names"] if ("machine_names" in row.keys() and row["machine_names"]) else "",
                                    is_active=bool(row["is_active"]),
                                    notes=row["notes"] or "",
                                    created_at=row["created_at"],
                                    updated_at=row["updated_at"],
                                )
                            )
                except Exception as local_err:
                    logger.error("Error reading fallback local cache: %s", local_err)

        return results

    def get_members_for_machine(self, machine_name: str) -> list[MemberRecord]:
        """Query all active members assigned to a specific machine/model.

        Matches comma-separated machine_names (case-insensitive, exact or substring).
        """
        clean_target = re.sub(r"\s+", "", machine_name.strip()).lower()
        all_active = self.get_members(active_only=True)
        if not clean_target:
            return all_active

        matched = []
        for m in all_active:
            raw_models = getattr(m, "machine_names", "") or ""
            m_models = [re.sub(r"\s+", "", mod.strip()).lower() for mod in raw_models.split(",") if mod.strip()]
            if any(clean_target == mod or clean_target in mod or mod in clean_target for mod in m_models):
                matched.append(m)

        return matched

    def add_member(self, member: MemberRecord, sync_to_excel: bool = False) -> tuple[bool, str]:
        """Add a new member to the database via fast dual-write micro-transaction."""
        if not member.account_id or not member.account_id.strip():
            return False, "Mã thành viên / Tài khoản không được để trống."

        if not member.department or not member.department.strip():
            return False, "Phòng ban không được để trống."

        full_name = member.full_name.strip() if member.full_name else member.account_id.strip()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        raw_mach = getattr(member, "machine_names", "") or ""
        clean_mach = ", ".join([" ".join(t.strip().split()) for t in raw_mach.split(",") if t.strip()])

        def _do_add(conn: sqlite3.Connection):
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO members (
                    account_id, full_name, department, default_sub_unit, machine_names,
                    is_active, notes, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
                """,
                (
                    member.account_id.strip(),
                    full_name,
                    member.department.strip(),
                    member.default_sub_unit.strip(),
                    clean_mach,
                    1 if member.is_active else 0,
                    member.notes.strip(),
                    now_str,
                    now_str,
                ),
            )
            return cursor.lastrowid

        ok, result, used_fallback = self._execute_dual_write(_do_add, operation_name=f"thêm '{member.account_id}'")
        if not ok:
            err_msg = str(result)
            if "unique" in err_msg.lower() or "trùng" in err_msg.lower() or "exists" in err_msg.lower():
                return False, f"Mã thành viên '{member.account_id}' đã tồn tại trong CSDL."
            return False, f"Lỗi lưu CSDL: {err_msg}"

        excel_note = ""
        if sync_to_excel:
            try:
                ok_ex, ex_msg = self.export_to_excel()
                if ok_ex:
                    excel_note = " (Đã đồng bộ sang file Excel)"
                else:
                    logger.warning("Could not sync added member to Excel: %s", ex_msg)
            except Exception as ex_err:
                logger.warning("Could not sync added member to Excel: %s", ex_err)

        fallback_note = " (Đã lưu vào bộ nhớ tạm cục bộ do CSDL mạng bận)" if used_fallback else ""
        return True, f"Đã thêm thành viên '{member.account_id}' thành công.{fallback_note}{excel_note}"

    def update_member(self, member: MemberRecord, sync_to_excel: bool = False) -> tuple[bool, str]:
        """Update existing member information via fast dual-write micro-transaction."""
        new_acc_id = member.account_id.strip() if member.account_id else ""
        if not new_acc_id:
            return False, "Mã thành viên không được để trống."

        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        raw_mach = getattr(member, "machine_names", "") or ""
        clean_mach = ", ".join([" ".join(t.strip().split()) for t in raw_mach.split(",") if t.strip()])
        full_name = member.full_name.strip() if member.full_name else new_acc_id

        def _do_update(conn: sqlite3.Connection):
            cursor = conn.cursor()
            if member.id is not None:
                cursor.execute(
                    "SELECT id FROM members WHERE account_id = ? AND id != ?;",
                    (new_acc_id, member.id),
                )
                if cursor.fetchone():
                    raise sqlite3.IntegrityError(f"Mã thành viên '{new_acc_id}' đã trùng với thành viên khác trong CSDL.")

                cursor.execute(
                    """
                    UPDATE members SET
                        account_id = ?,
                        full_name = ?,
                        department = ?,
                        default_sub_unit = ?,
                        machine_names = ?,
                        is_active = ?,
                        notes = ?,
                        updated_at = ?
                    WHERE id = ?;
                    """,
                    (
                        new_acc_id,
                        full_name,
                        member.department.strip(),
                        member.default_sub_unit.strip(),
                        clean_mach,
                        1 if member.is_active else 0,
                        member.notes.strip(),
                        now_str,
                        member.id,
                    ),
                )
                if cursor.rowcount == 0:
                    cursor.execute(
                        """
                        UPDATE members SET
                            full_name = ?,
                            department = ?,
                            default_sub_unit = ?,
                            machine_names = ?,
                            is_active = ?,
                            notes = ?,
                            updated_at = ?
                        WHERE account_id = ?;
                        """,
                        (
                            full_name,
                            member.department.strip(),
                            member.default_sub_unit.strip(),
                            clean_mach,
                            1 if member.is_active else 0,
                            member.notes.strip(),
                            now_str,
                            new_acc_id,
                        ),
                    )
            else:
                cursor.execute(
                    """
                    UPDATE members SET
                        full_name = ?,
                        department = ?,
                        default_sub_unit = ?,
                        machine_names = ?,
                        is_active = ?,
                        notes = ?,
                        updated_at = ?
                    WHERE account_id = ?;
                    """,
                    (
                        full_name,
                        member.department.strip(),
                        member.default_sub_unit.strip(),
                        clean_mach,
                        1 if member.is_active else 0,
                        member.notes.strip(),
                        now_str,
                        new_acc_id,
                    ),
                )

            if cursor.rowcount == 0:
                raise LookupError(f"Không tìm thấy thành viên '{new_acc_id}' để cập nhật.")
            return cursor.rowcount

        ok, result, used_fallback = self._execute_dual_write(_do_update, operation_name=f"cập nhật '{new_acc_id}'")
        if not ok:
            err_msg = str(result)
            if "unique" in err_msg.lower() or "trùng" in err_msg.lower():
                return False, f"Mã thành viên '{new_acc_id}' đã trùng với thành viên khác trong CSDL."
            if "không tìm thấy" in err_msg.lower():
                return False, err_msg
            return False, f"Lỗi cập nhật CSDL: {err_msg}"

        excel_note = ""
        if sync_to_excel:
            try:
                ok_ex, ex_msg = self.export_to_excel()
                if ok_ex:
                    excel_note = " (Đã đồng bộ 2 chiều sang file Excel)"
                else:
                    logger.warning("Could not sync updated member to Excel: %s", ex_msg)
            except Exception as ex_err:
                logger.warning("Could not sync updated member to Excel: %s", ex_err)

        fallback_note = " (Đã lưu vào bộ nhớ tạm cục bộ do CSDL mạng bận)" if used_fallback else ""
        return True, f"Đã cập nhật thông tin thành viên '{new_acc_id}' thành công.{fallback_note}{excel_note}"

    def delete_member(self, account_id: str, member_id: int | None = None, sync_to_excel: bool = False) -> tuple[bool, str]:
        """Delete member by member_id or account_id via fast dual-write micro-transaction."""
        acc_id = account_id.strip() if account_id else ""
        if not acc_id and member_id is None:
            return False, "Mã thành viên không hợp lệ."

        def _do_delete(conn: sqlite3.Connection):
            cursor = conn.cursor()
            if member_id is not None and acc_id:
                cursor.execute("DELETE FROM members WHERE id = ? OR account_id = ?;", (member_id, acc_id))
            elif member_id is not None:
                cursor.execute("DELETE FROM members WHERE id = ?;", (member_id,))
            else:
                cursor.execute("DELETE FROM members WHERE account_id = ?;", (acc_id,))

            if cursor.rowcount == 0:
                raise LookupError(f"Không tìm thấy thành viên '{acc_id}' để xóa.")
            return cursor.rowcount

        ok, result, used_fallback = self._execute_dual_write(_do_delete, operation_name=f"xóa '{acc_id}'")
        if not ok:
            err_msg = str(result)
            if "không tìm thấy" in err_msg.lower():
                return False, err_msg
            return False, f"Lỗi xóa dữ liệu: {err_msg}"

        if sync_to_excel:
            try:
                self.export_to_excel()
            except Exception:
                pass

        fallback_note = " (Đã xóa trên bộ nhớ tạm cục bộ do CSDL mạng bận)" if used_fallback else ""
        return True, f"Đã xóa thành viên '{acc_id or str(member_id)}' thành công.{fallback_note}"

    def import_from_excel(self, excel_path: Path | str | None = None, clean_legacy_seeds: bool = True) -> tuple[int, int, str]:
        """Import or update members from Danhsachthanhvien.xlsx / .xlsm sheet 'DS KDTVN'.

        Maps:
          Col B (2): Mã nhân viên
          Col C (3): Mã thành viên -> account_id
          Col D (4): Họ và tên -> full_name
          Col E (5): Phòng ban -> department
          Col F (6): Công đoạn mặc định -> default_sub_unit
          Col G (7): Dòng máy phụ trách -> machine_names

        Args:
            excel_path: Optional explicit file path.
            clean_legacy_seeds: If True, cleans up old placeholder seeds ('tenphong_pt')
                                that are not in the real imported Excel file.

        Returns:
            (added_count, updated_count, message)
        """
        target_path = Path(excel_path) if excel_path else self.get_member_excel_path()
        if not target_path.exists():
            return 0, 0, f"Không tìm thấy file Excel thành viên tại: {target_path}"

        try:
            wb = openpyxl.load_workbook(str(target_path), data_only=True)
        except Exception as err:
            logger.error("Failed to load Excel %s: %s", target_path, err)
            return 0, 0, f"Lỗi đọc file Excel {target_path.name}: {err}"

        sheet_name = "DS KDTVN" if "DS KDTVN" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        # 1. Identify header row and column indices
        col_acc = 3      # Default Col C: Mã thành viên
        col_name = 4     # Default Col D: Họ và tên
        col_dept = 5     # Default Col E: Phòng ban
        col_subunit = 6  # Default Col F: Công đoạn mặc định
        col_machine = 7  # Default Col G: Dòng máy phụ trách
        col_empid = 2    # Col B: Mã nhân viên

        header_row = 4
        for r in range(1, min(15, ws.max_row + 1)):
            for c in range(1, min(12, ws.max_column + 1)):
                val = str(ws.cell(row=r, column=c).value or "").strip().lower()
                if "mã thành viên" in val or "ma thanh vien" in val:
                    header_row = r
                    col_acc = c
                elif "họ và tên" in val or "ho va ten" in val or "氏名" in val:
                    col_name = c
                elif "phòng ban" in val or "phong ban" in val or "課" in val:
                    col_dept = c
                elif "công đoạn" in val or "cong doan" in val:
                    col_subunit = c
                elif "dòng máy" in val or "dong may" in val or "tên máy" in val:
                    col_machine = c
                elif "mã \nnhân viên" in val or "nhân viên" in val or "社員番号" in val:
                    col_empid = c

        added_count = 0
        updated_count = 0
        db_path, is_remote = self.get_active_db_path()

        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        imported_account_ids: set[str] = set()

        try:
            with self._write_transaction(db_path) as conn:
                cursor = conn.cursor()

                for r in range(header_row + 1, ws.max_row + 1):
                    raw_acc = ws.cell(row=r, column=col_acc).value
                    raw_name = ws.cell(row=r, column=col_name).value
                    raw_dept = ws.cell(row=r, column=col_dept).value
                    raw_subunit = ws.cell(row=r, column=col_subunit).value
                    raw_machine = ws.cell(row=r, column=col_machine).value
                    raw_empid = ws.cell(row=r, column=col_empid).value

                    # If row is completely empty, skip
                    if not any([raw_acc, raw_name, raw_dept, raw_subunit, raw_machine, raw_empid]):
                        continue

                    full_name = str(raw_name).strip() if raw_name is not None else ""
                    dept = str(raw_dept).strip() if raw_dept is not None else "Cơ 1"
                    subunit = str(raw_subunit).strip() if raw_subunit is not None else ""
                    machine = str(raw_machine).strip() if raw_machine is not None else ""
                    emp_id = str(raw_empid).strip() if raw_empid is not None else ""

                    account_id = str(raw_acc).strip() if raw_acc is not None else ""
                    if not account_id:
                        if full_name:
                            clean_fn = re.sub(r"\s+", "", full_name)
                            clean_dept = re.sub(r"[^A-Za-z0-9]", "", dept)
                            if "Điện" in dept or "Dien" in dept:
                                account_id = "HaiDang_dien"
                            else:
                                account_id = f"{clean_fn}_{clean_dept}" if clean_dept else clean_fn
                        elif emp_id:
                            account_id = f"NV_{emp_id}"
                        else:
                            continue

                    # Clean machine names preserving internal spaces
                    tokens = [" ".join(m.strip().split()) for m in machine.split(",") if m.strip()]
                    clean_machines = ", ".join(tokens)

                    imported_account_ids.add(account_id.lower())

                    # Check if exists
                    cursor.execute("SELECT id FROM members WHERE LOWER(account_id) = LOWER(?);", (account_id,))
                    row = cursor.fetchone()
                    if row:
                        cursor.execute(
                            """
                            UPDATE members SET
                                full_name = ?,
                                department = ?,
                                default_sub_unit = ?,
                                machine_names = ?,
                                notes = ?,
                                updated_at = ?
                            WHERE id = ?;
                            """,
                            (full_name or account_id, dept, subunit, clean_machines, f"Đồng bộ từ {target_path.name}", now_str, row[0]),
                        )
                        updated_count += 1
                    else:
                        cursor.execute(
                            """
                            INSERT INTO members (
                                account_id, full_name, department, default_sub_unit, machine_names,
                                is_active, notes, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, 1, ?, ?, ?);
                            """,
                            (account_id, full_name or account_id, dept, subunit, clean_machines, f"Đồng bộ từ {target_path.name}", now_str, now_str),
                        )
                        added_count += 1

                # Clean legacy placeholder seeds if requested
                if clean_legacy_seeds and imported_account_ids:
                    cursor.execute("SELECT id, account_id FROM members WHERE notes LIKE '%tenphong_pt%';")
                    legacy_rows = cursor.fetchall()
                    for lid, lacc in legacy_rows:
                        if lacc.lower() not in imported_account_ids:
                            cursor.execute("DELETE FROM members WHERE id = ?;", (lid,))

            wb.close()
            self._sync_active_with_alternate(db_path, is_remote)
            msg = f"Đã nạp thành công {added_count + updated_count} thành viên ({added_count} mới, {updated_count} cập nhật) từ file Excel {target_path.name}."
            logger.info(msg)
            return added_count + updated_count, updated_count, msg
        except Exception as exc:
            logger.error("Error during import_from_excel: %s", exc)
            return added_count + updated_count, updated_count, f"Lỗi nạp dữ liệu từ Excel: {exc}"

    def export_to_excel(self, excel_path: Path | str | None = None) -> tuple[bool, str]:
        """Export/write back current members from database to Danhsachthanhvien.xlsx / .xlsm.

        Preserves VBA macros (if .xlsm) and creates an automatic backup (.xlsx.bak / .xlsm.bak).
        Updates Cols C, D, E, F, G for matching account_ids and appends new members.
        """
        target_path = Path(excel_path) if excel_path else self.get_member_excel_path()
        if not target_path.exists():
            return False, f"Không tìm thấy file Excel để ghi tại: {target_path}"

        # 1. Create backup with proper suffix
        backup_path = target_path.with_suffix(f"{target_path.suffix}.bak")
        try:
            shutil.copy2(str(target_path), str(backup_path))
        except Exception as b_err:
            logger.warning("Could not create backup %s: %s", backup_path, b_err)

        # 2. Open workbook with VBA preservation if applicable
        is_vba = target_path.suffix.lower() in (".xlsm", ".xltm")
        try:
            wb = openpyxl.load_workbook(str(target_path), keep_vba=is_vba)
        except Exception as err:
            logger.error("Failed to load workbook %s for export: %s", target_path, err)
            return False, f"Không thể mở file Excel để ghi (file có thể đang bị khóa): {err}"

        sheet_name = "DS KDTVN" if "DS KDTVN" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        # 3. Identify header columns
        col_acc = 3      # Col C
        col_name = 4     # Col D
        col_dept = 5     # Col E
        col_subunit = 6  # Col F
        col_machine = 7  # Col G

        header_row = 4
        for r in range(1, min(15, ws.max_row + 1)):
            for c in range(1, min(12, ws.max_column + 1)):
                val = str(ws.cell(row=r, column=c).value or "").strip().lower()
                if "mã thành viên" in val or "ma thanh vien" in val:
                    header_row = r
                    col_acc = c
                elif "họ và tên" in val or "ho va ten" in val or "氏名" in val:
                    col_name = c
                elif "phòng ban" in val or "phong ban" in val or "課" in val:
                    col_dept = c
                elif "công đoạn" in val or "cong doan" in val:
                    col_subunit = c
                elif "dòng máy" in val or "dong may" in val or "tên máy" in val:
                    col_machine = c

        # 4. Map existing rows in Excel by account_id and full_name
        existing_by_acc: dict[str, int] = {}
        existing_by_name: dict[str, int] = {}
        last_non_empty_row = header_row

        for r in range(header_row + 1, ws.max_row + 1):
            acc_val = ws.cell(row=r, column=col_acc).value
            name_val = ws.cell(row=r, column=col_name).value
            if acc_val or name_val:
                last_non_empty_row = max(last_non_empty_row, r)
            if acc_val:
                existing_by_acc[str(acc_val).strip().lower()] = r
            if name_val:
                existing_by_name[str(name_val).strip().lower()] = r

        # 5. Fetch members from SQLite
        members = self.get_members(active_only=False)
        updated_rows = 0
        added_rows = 0

        next_append_row = last_non_empty_row + 1
        contiguous_last = header_row
        for r in range(header_row + 1, min(last_non_empty_row + 1, 100)):
            if ws.cell(row=r, column=col_acc).value or ws.cell(row=r, column=col_name).value:
                contiguous_last = r
        if contiguous_last > 0 and (last_non_empty_row - contiguous_last > 100):
            next_append_row = contiguous_last + 1

        for m in members:
            acc_key = m.account_id.strip().lower()
            name_key = m.full_name.strip().lower() if m.full_name else acc_key

            row_target: int | None = existing_by_acc.get(acc_key)
            if row_target is None and name_key in existing_by_name:
                row_target = existing_by_name[name_key]

            if row_target is not None:
                # Update existing row
                ws.cell(row=row_target, column=col_acc).value = m.account_id
                ws.cell(row=row_target, column=col_name).value = m.full_name
                ws.cell(row=row_target, column=col_dept).value = m.department
                if m.default_sub_unit:
                    ws.cell(row=row_target, column=col_subunit).value = m.default_sub_unit
                if m.machine_names:
                    ws.cell(row=row_target, column=col_machine).value = m.machine_names
                updated_rows += 1
            else:
                # Append new row
                r = next_append_row
                next_append_row += 1
                ws.cell(row=r, column=1).value = r - header_row
                ws.cell(row=r, column=col_acc).value = m.account_id
                ws.cell(row=r, column=col_name).value = m.full_name
                ws.cell(row=r, column=col_dept).value = m.department
                ws.cell(row=r, column=col_subunit).value = m.default_sub_unit or None
                ws.cell(row=r, column=col_machine).value = m.machine_names or None
                ws.cell(row=r, column=8).value = "Bộ phận kỹ thuật chế tạo"
                existing_by_acc[acc_key] = r
                added_rows += 1

        try:
            wb.save(str(target_path))
            wb.close()
            msg = f"Đã ghi 2 chiều thành công sang file Excel {target_path.name} ({updated_rows} cập nhật, {added_rows} thêm mới)."
            logger.info(msg)
            return True, msg
        except PermissionError:
            wb.close()
            err_msg = (
                f"Không thể ghi vào file '{target_path.name}'.\n"
                "File đang được mở bởi người khác hoặc đang ở chế độ chỉ đọc (Read-only).\n"
                "Vui lòng đóng file Excel trên mạng LAN và thử lại!"
            )
            logger.warning(err_msg)
            return False, err_msg
        except Exception as save_err:
            wb.close()
            err_msg = f"Lỗi lưu file Excel {target_path.name}: {save_err}"
            logger.error(err_msg)
            return False, err_msg

    def _sync_active_with_alternate(self, active_path: Path, is_remote: bool) -> None:
        """Mirror updated DB between remote and local cache."""
        try:
            if is_remote and safe_exists(active_path):
                self._safe_backup_db(active_path, self.local_cache_path)
            elif not is_remote and self.is_remote_available():
                self._safe_backup_db(self.local_cache_path, self.remote_db_path)
        except Exception as sync_err:
            logger.debug("Mirror sync skipped: %s", sync_err)

    def get_all_account_names(self, active_only: bool = True) -> list[str]:
        """Get flat list of account IDs (for QCompleter / dropdowns)."""
        members = self.get_members(active_only=active_only)
        if members:
            return [m.account_id for m in members]
        # Fallback to seed arrays if DB is completely empty
        return SEED_ROSTER_MECHA_1 + SEED_ROSTER_MECHA_2 + SEED_ROSTER_MECHA_3

    def get_departments(self) -> list[str]:
        """Get distinct departments list (Local-First fast read)."""
        db_path = self.get_read_db_path()
        try:
            with self._get_connection(db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT DISTINCT department FROM members ORDER BY department ASC;")
                rows = cursor.fetchall()
                if rows:
                    return [r[0] for r in rows if r[0]]
        except Exception:
            pass
        return ["Cơ 1", "Cơ 2", "Cơ 3"]

    def get_roster_by_department(self, active_only: bool = True) -> dict[str, list[str]]:
        """Get dictionary mapping department to list of engineer account IDs."""
        members = self.get_members(active_only=active_only)
        result: dict[str, list[str]] = {}
        for m in members:
            result.setdefault(m.department, []).append(m.account_id)
        if not result:
            result["Cơ 1"] = list(SEED_ROSTER_MECHA_1)
            result["Cơ 2"] = list(SEED_ROSTER_MECHA_2)
            result["Cơ 3"] = list(SEED_ROSTER_MECHA_3)
        return result


_member_db_instances: dict[str, MemberDatabaseManager] = {}


def get_member_database_manager(
    base_dir: Path | str | None = None,
    remote_db_path: Path | str | None = None,
    excel_path: Path | str | None = None,
) -> MemberDatabaseManager:
    """Return a cached or new MemberDatabaseManager instance."""
    key = f"{base_dir}_{remote_db_path}_{excel_path}"
    if key not in _member_db_instances:
        _member_db_instances[key] = MemberDatabaseManager(base_dir, remote_db_path, excel_path)
    return _member_db_instances[key]
