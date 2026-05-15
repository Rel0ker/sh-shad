"""SQLite: справочники, дни изменений, строки изменений."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
import sys
from pathlib import Path
from typing import Any, Generator, Iterable


def get_db_path() -> Path:
    """Рядом с .exe при сборке PyInstaller; иначе — корень проекта."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "app.db"
    return Path(__file__).resolve().parent.parent / "app.db"


@contextmanager
def get_connection() -> Generator[sqlite3.Connection, None, None]:
    path = get_db_path()
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _resource_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "data"
    return Path(__file__).resolve().parent.parent / "data"


def init_db() -> None:
    data_dir = _resource_data_dir()
    seed_teachers = data_dir / "seed_teachers.json"
    seed_classes = data_dir / "seed_classes.json"

    with get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS teachers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fio TEXT NOT NULL UNIQUE,
                active INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS classes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                shift INTEGER NOT NULL CHECK (shift IN (1, 2))
            );
            
            CREATE TABLE IF NOT EXISTS subjects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                active INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS change_days (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day_id INTEGER NOT NULL REFERENCES change_days(id) ON DELETE CASCADE,
                class_id INTEGER REFERENCES classes(id) ON DELETE SET NULL,
                klass TEXT NOT NULL,
                lesson_no INTEGER NOT NULL,
                absent_fio TEXT NOT NULL DEFAULT '',
                replacement_fio TEXT NOT NULL DEFAULT '',
                subject TEXT NOT NULL DEFAULT '',
                room TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                sort_order INTEGER NOT NULL DEFAULT 0
            );

            CREATE INDEX IF NOT EXISTS idx_changes_day ON changes(day_id);
            """
        )

        # Seed empty lists from JSON if tables empty
        if seed_teachers.exists():
            teachers = json.loads(seed_teachers.read_text(encoding="utf-8"))
            if isinstance(teachers, list) and teachers:
                cur = conn.execute("SELECT COUNT(*) AS c FROM teachers")
                if cur.fetchone()["c"] == 0:
                    for fio in teachers:
                        if isinstance(fio, str) and fio.strip():
                            conn.execute(
                                "INSERT OR IGNORE INTO teachers (fio) VALUES (?)",
                                (fio.strip(),),
                            )

        if seed_classes.exists():
            classes = json.loads(seed_classes.read_text(encoding="utf-8"))
            if isinstance(classes, list) and classes:
                cur = conn.execute("SELECT COUNT(*) AS c FROM classes")
                if cur.fetchone()["c"] == 0:
                    for item in classes:
                        if isinstance(item, dict) and item.get("name"):
                            shift = int(item.get("shift", 1))
                            if shift not in (1, 2):
                                shift = 1
                            conn.execute(
                                "INSERT OR IGNORE INTO classes (name, shift) VALUES (?, ?)",
                                (str(item["name"]).strip(), shift),
                            )


def list_teachers(active_only: bool = True) -> list[dict[str, Any]]:
    q = "SELECT id, fio, active FROM teachers"
    if active_only:
        q += " WHERE active = 1"
    q += " ORDER BY fio COLLATE NOCASE"
    with get_connection() as conn:
        rows = conn.execute(q).fetchall()
    return [dict(r) for r in rows]


def list_classes() -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT id, name, shift FROM classes ORDER BY name COLLATE NOCASE"
        ).fetchall()
    return [dict(r) for r in rows]


def list_subjects(active_only: bool = True) -> list[dict[str, Any]]:
    q = "SELECT id, name, active FROM subjects"
    if active_only:
        q += " WHERE active = 1"
    q += " ORDER BY name COLLATE NOCASE"
    with get_connection() as conn:
        rows = conn.execute(q).fetchall()
    return [dict(r) for r in rows]


def add_teacher(fio: str) -> int:
    fio = fio.strip()
    if not fio:
        raise ValueError("Пустое ФИО")
    with get_connection() as conn:
        cur = conn.execute("INSERT OR IGNORE INTO teachers (fio) VALUES (?)", (fio,))
        if cur.rowcount == 0:
            raise ValueError("Учитель уже существует")
        return int(cur.lastrowid)


def add_class(name: str, shift: int) -> int:
    name = name.strip()
    if not name:
        raise ValueError("Пустое название класса")
    if shift not in (1, 2):
        raise ValueError("Смена должна быть 1 или 2")
    with get_connection() as conn:
        cur = conn.execute(
            "INSERT OR IGNORE INTO classes (name, shift) VALUES (?, ?)", (name, shift)
        )
        if cur.rowcount == 0:
            raise ValueError("Класс уже существует")
        return int(cur.lastrowid)


def add_subject(name: str) -> int:
    name = name.strip()
    if not name:
        raise ValueError("Пустой предмет")
    with get_connection() as conn:
        cur = conn.execute("INSERT OR IGNORE INTO subjects (name) VALUES (?)", (name,))
        if cur.rowcount == 0:
            raise ValueError("Предмет уже существует")
        return int(cur.lastrowid)


def delete_teacher(tid: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM teachers WHERE id = ?", (tid,))


def delete_class(cid: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM classes WHERE id = ?", (cid,))


def delete_subject(sid: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM subjects WHERE id = ?", (sid,))


def upsert_subjects(names: Iterable[str]) -> int:
    inserted = 0
    with get_connection() as conn:
        for name in names:
            n = (name or "").strip()
            if not n:
                continue
            cur = conn.execute(
                "INSERT OR IGNORE INTO subjects (name) VALUES (?)",
                (n,),
            )
            if cur.rowcount:
                inserted += 1
    return inserted


def upsert_teachers(fios: Iterable[str]) -> int:
    inserted = 0
    with get_connection() as conn:
        for fio in fios:
            f = (fio or "").strip()
            if not f:
                continue
            cur = conn.execute(
                "INSERT OR IGNORE INTO teachers (fio) VALUES (?)",
                (f,),
            )
            if cur.rowcount:
                inserted += 1
    return inserted


def upsert_classes_from_rows(rows: Iterable[dict[str, Any]]) -> int:
    """Новые классы из ввода — в справочник (смена 1 по умолчанию)."""
    inserted = 0
    with get_connection() as conn:
        for r in rows:
            name = (r.get("klass") or "").strip()
            if not name:
                continue
            exists = conn.execute(
                "SELECT id FROM classes WHERE name = ? COLLATE NOCASE",
                (name,),
            ).fetchone()
            if exists:
                continue
            shift = 1
            cid = r.get("class_id")
            if cid is not None:
                row = conn.execute(
                    "SELECT shift FROM classes WHERE id = ?", (int(cid),)
                ).fetchone()
                if row:
                    shift = int(row["shift"])
            cur = conn.execute(
                "INSERT OR IGNORE INTO classes (name, shift) VALUES (?, ?)",
                (name, shift),
            )
            if cur.rowcount:
                inserted += 1
    return inserted


def remember_from_change_rows(rows: Iterable[dict[str, Any]]) -> dict[str, int]:
    """Сохранить в справочники всё, что ввели в таблице (учителя, предметы, классы)."""
    teachers: set[str] = set()
    subjects: set[str] = set()
    row_list = list(rows)
    for r in row_list:
        for key in ("absent_fio", "replacement_fio"):
            f = (r.get(key) or "").strip()
            if f:
                teachers.add(f)
        subj = (r.get("subject") or "").strip()
        if subj:
            subjects.add(subj)
    return {
        "teachers": upsert_teachers(teachers),
        "subjects": upsert_subjects(subjects),
        "classes": upsert_classes_from_rows(row_list),
    }


def get_or_create_day(d: str) -> int:
    """d — YYYY-MM-DD. Возвращает id change_days."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT id FROM change_days WHERE date = ?", (d,)
        ).fetchone()
        if row:
            return int(row["id"])
        now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
        cur = conn.execute(
            "INSERT INTO change_days (date, created_at) VALUES (?, ?)",
            (d, now),
        )
        return int(cur.lastrowid)


def get_day_id(d: str) -> int | None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT id FROM change_days WHERE date = ?", (d,)
        ).fetchone()
    return int(row["id"]) if row else None


def replace_changes_for_day(day_id: int, rows: Iterable[dict[str, Any]]) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM changes WHERE day_id = ?", (day_id,))
        for i, r in enumerate(rows):
            conn.execute(
                """
                INSERT INTO changes (
                    day_id, class_id, klass, lesson_no,
                    absent_fio, replacement_fio, subject, room, note, sort_order
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    day_id,
                    r.get("class_id"),
                    r.get("klass") or "",
                    int(r["lesson_no"]),
                    r.get("absent_fio") or "",
                    r.get("replacement_fio") or "",
                    r.get("subject") or "",
                    r.get("room") or "",
                    r.get("note") or "",
                    i,
                ),
            )


def load_changes_for_day(d: str) -> list[dict[str, Any]]:
    day_id = get_day_id(d)
    if not day_id:
        return []
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, day_id, class_id, klass, lesson_no,
                   absent_fio, replacement_fio, subject, room, note, sort_order
            FROM changes WHERE day_id = ?
            ORDER BY sort_order, id
            """,
            (day_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def load_changes_between(date_from: str, date_to: str) -> list[dict[str, Any]]:
    """Все строки изменений за период [date_from, date_to] с полем day_date (YYYY-MM-DD)."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT c.id, c.day_id, c.class_id, c.klass, c.lesson_no,
                   c.absent_fio, c.replacement_fio, c.subject, c.room, c.note, c.sort_order,
                   cd.date AS day_date
            FROM changes c
            JOIN change_days cd ON c.day_id = cd.id
            WHERE cd.date >= ? AND cd.date <= ?
            ORDER BY cd.date, c.sort_order, c.id
            """,
            (date_from, date_to),
        ).fetchall()
    return [dict(r) for r in rows]


def class_shift_by_id(class_id: int | None) -> int | None:
    if not class_id:
        return None
    with get_connection() as conn:
        row = conn.execute(
            "SELECT shift FROM classes WHERE id = ?", (class_id,)
        ).fetchone()
    return int(row["shift"]) if row else None


def class_shift_by_name(name: str) -> int | None:
    name = (name or "").strip()
    if not name:
        return None
    with get_connection() as conn:
        row = conn.execute(
            "SELECT shift FROM classes WHERE name = ? COLLATE NOCASE",
            (name,),
        ).fetchone()
    return int(row["shift"]) if row else None


def today_iso() -> str:
    return date.today().isoformat()
