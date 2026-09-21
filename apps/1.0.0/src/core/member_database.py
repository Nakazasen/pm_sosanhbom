"""Centralized SQLite Member Management Database Module.

Provides shared database persistence and offline caching for Kyocera SSBOM:
- Remote primary DB (UNC): \\\\fstvn01\\Data\\00_KDTVN Common(KDTVN共通)\\⑤Production Engineering(製造技術)\\Hang muc can luu\\Vinh\\Pm_sosanhBOM\\ssbom_master.db
- Local offline cache fallback: <base_dir>/data/ssbom_master_cache.db
- Full CRUD operations for team personnel (engineers, departments, default sub-units).
- Automatic database migration and seeding with the 38 canonical engineers.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
import datetime
import json
import logging
import os
from pathlib import Path
import shutil
import sqlite3
from typing import Any

logger = logging.getLogger(__name__)

# Default UNC Network path requested by User
DEFAULT_SHARED_DB_PATH = (
    r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)"
    r"\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db"
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
    ) -> None:
        self.base_dir = Path(base_dir or r"D:\Sandbox\pm_sosanhbom")
        if remote_db_path:
            self.remote_db_path = Path(remote_db_path)
        else:
            local_master = self.base_dir / "ssbom_master.db"
            cfg_path = self.base_dir / "config" / "settings.json"
            cfg_shared = ""
            if cfg_path.exists():
                try:
                    with open(cfg_path, "r", encoding="utf-8") as fp:
                        data = json.load(fp)
                        cfg_shared = data.get("paths", {}).get("shared_db_path", "")
                except Exception:
                    pass
            if cfg_shared:
                self.remote_db_path = Path(cfg_shared)
            elif local_master.exists():
                self.remote_db_path = local_master
            else:
                self.remote_db_path = DEFAULT_SHARED_DB_PATH

        self.local_cache_dir = self.base_dir / "data"
        self.local_cache_dir.mkdir(parents=True, exist_ok=True)
        self.local_cache_path = self.local_cache_dir / "ssbom_master_cache.db"

        # Initialize schema and seed default data
        self.initialize_database()

    def is_remote_available(self) -> bool:
        """Check if remote UNC share or directory is accessible."""
        try:
            # Check parent directory or file access
            parent_dir = self.remote_db_path.parent
            if parent_dir.exists():
                return True
        except Exception:
            pass
        return False

    def get_active_db_path(self) -> tuple[Path, bool]:
        """Determine whether to use remote network database or local cache.
        
        Returns:
            (active_path, is_remote_connected)
        """
        if self.is_remote_available():
            return self.remote_db_path, True
        return self.local_cache_path, False

    def _get_connection(self, db_path: Path) -> sqlite3.Connection:
        """Open an SQLite connection with WAL mode and row factory."""
        db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(db_path), timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL;")
        return conn

    def initialize_database(self) -> None:
        """Create tables if not existing, seed canonical engineers, and sync cache."""
        target_path, is_remote = self.get_active_db_path()

        try:
            with self._get_connection(target_path) as conn:
                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS members (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        account_id TEXT UNIQUE NOT NULL,
                        full_name TEXT NOT NULL,
                        department TEXT NOT NULL,
                        default_sub_unit TEXT DEFAULT '',
                        is_active INTEGER DEFAULT 1,
                        notes TEXT DEFAULT '',
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                    );
                    """
                )
                conn.execute(
                    "CREATE INDEX IF NOT EXISTS idx_members_dept ON members(department);"
                )
                conn.execute(
                    "CREATE INDEX IF NOT EXISTS idx_members_active ON members(is_active);"
                )

                # Check if table is empty
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM members;")
                count = cursor.fetchone()[0]

                if count == 0:
                    logger.info("Seeding database '%s' with 38 initial engineers...", target_path)
                    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    seed_items: list[tuple[str, str, str, str, int, str, str, str]] = []
                    for name in SEED_ROSTER_MECHA_1:
                        seed_items.append((name, name, "Cơ 1", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                    for name in SEED_ROSTER_MECHA_2:
                        seed_items.append((name, name, "Cơ 2", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                    for name in SEED_ROSTER_MECHA_3:
                        seed_items.append((name, name, "Cơ 3", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))

                    conn.executemany(
                        """
                        INSERT OR IGNORE INTO members (
                            account_id, full_name, department, default_sub_unit,
                            is_active, notes, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
                        """,
                        seed_items,
                    )
                    conn.commit()

            # If remote was used, update local cache
            if is_remote and target_path.exists():
                try:
                    shutil.copy2(target_path, self.local_cache_path)
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
                        conn.execute(
                            """
                            CREATE TABLE IF NOT EXISTS members (
                                id INTEGER PRIMARY KEY AUTOINCREMENT,
                                account_id TEXT UNIQUE NOT NULL,
                                full_name TEXT NOT NULL,
                                department TEXT NOT NULL,
                                default_sub_unit TEXT DEFAULT '',
                                is_active INTEGER DEFAULT 1,
                                notes TEXT DEFAULT '',
                                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                            );
                            """
                        )
                        cursor = conn.cursor()
                        cursor.execute("SELECT COUNT(*) FROM members;")
                        if cursor.fetchone()[0] == 0:
                            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            seed_items = []
                            for name in SEED_ROSTER_MECHA_1:
                                seed_items.append((name, name, "Cơ 1", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                            for name in SEED_ROSTER_MECHA_2:
                                seed_items.append((name, name, "Cơ 2", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                            for name in SEED_ROSTER_MECHA_3:
                                seed_items.append((name, name, "Cơ 3", "", 1, "Mặc định từ tenphong_pt", now_str, now_str))
                            conn.executemany(
                                """
                                INSERT OR IGNORE INTO members (
                                    account_id, full_name, department, default_sub_unit,
                                    is_active, notes, created_at, updated_at
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
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
        """Query members filtered by department and active status."""
        db_path, _ = self.get_active_db_path()
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
                                    is_active=bool(row["is_active"]),
                                    notes=row["notes"] or "",
                                    created_at=row["created_at"],
                                    updated_at=row["updated_at"],
                                )
                            )
                except Exception as local_err:
                    logger.error("Error reading fallback local cache: %s", local_err)

        return results

    def add_member(self, member: MemberRecord) -> tuple[bool, str]:
        """Add a new member to the database and sync cache."""
        if not member.account_id or not member.account_id.strip():
            return False, "Mã thành viên / Tài khoản không được để trống."

        if not member.department or not member.department.strip():
            return False, "Phòng ban không được để trống."

        full_name = member.full_name.strip() if member.full_name else member.account_id.strip()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        db_path, is_remote = self.get_active_db_path()
        try:
            with self._get_connection(db_path) as conn:
                conn.execute(
                    """
                    INSERT INTO members (
                        account_id, full_name, department, default_sub_unit,
                        is_active, notes, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
                    """,
                    (
                        member.account_id.strip(),
                        full_name,
                        member.department.strip(),
                        member.default_sub_unit.strip(),
                        1 if member.is_active else 0,
                        member.notes.strip(),
                        now_str,
                        now_str,
                    ),
                )
                conn.commit()

            # Sync to local cache or remote
            self._sync_active_with_alternate(db_path, is_remote)
            return True, f"Đã thêm thành viên '{member.account_id}' thành công."
        except sqlite3.IntegrityError:
            return False, f"Mã thành viên '{member.account_id}' đã tồn tại trong CSDL."
        except Exception as exc:
            logger.error("Failed to add member: %s", exc)
            return False, f"Lỗi lưu CSDL: {exc}"

    def update_member(self, member: MemberRecord) -> tuple[bool, str]:
        """Update existing member information."""
        if not member.account_id or not member.account_id.strip():
            return False, "Mã thành viên không hợp lệ."

        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db_path, is_remote = self.get_active_db_path()

        try:
            with self._get_connection(db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    UPDATE members SET
                        full_name = ?,
                        department = ?,
                        default_sub_unit = ?,
                        is_active = ?,
                        notes = ?,
                        updated_at = ?
                    WHERE account_id = ?;
                    """,
                    (
                        member.full_name.strip(),
                        member.department.strip(),
                        member.default_sub_unit.strip(),
                        1 if member.is_active else 0,
                        member.notes.strip(),
                        now_str,
                        member.account_id.strip(),
                    ),
                )
                if cursor.rowcount == 0:
                    return False, f"Không tìm thấy thành viên '{member.account_id}' để cập nhật."
                conn.commit()

            self._sync_active_with_alternate(db_path, is_remote)
            return True, f"Đã cập nhật thông tin thành viên '{member.account_id}' thành công."
        except Exception as exc:
            logger.error("Failed to update member: %s", exc)
            return False, f"Lỗi cập nhật CSDL: {exc}"

    def delete_member(self, account_id: str) -> tuple[bool, str]:
        """Delete member by account_id."""
        if not account_id or not account_id.strip():
            return False, "Mã thành viên không hợp lệ."

        db_path, is_remote = self.get_active_db_path()
        try:
            with self._get_connection(db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM members WHERE account_id = ?;", (account_id.strip(),))
                if cursor.rowcount == 0:
                    return False, f"Không tìm thấy thành viên '{account_id}' để xóa."
                conn.commit()

            self._sync_active_with_alternate(db_path, is_remote)
            return True, f"Đã xóa thành viên '{account_id}' thành công."
        except Exception as exc:
            logger.error("Failed to delete member: %s", exc)
            return False, f"Lỗi xóa dữ liệu: {exc}"

    def _sync_active_with_alternate(self, active_path: Path, is_remote: bool) -> None:
        """Mirror updated DB between remote and local cache."""
        try:
            if is_remote and active_path.exists():
                shutil.copy2(active_path, self.local_cache_path)
            elif not is_remote and self.is_remote_available():
                shutil.copy2(self.local_cache_path, self.remote_db_path)
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
        """Get distinct departments list."""
        db_path, _ = self.get_active_db_path()
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
