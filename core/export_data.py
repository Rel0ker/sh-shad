"""Подготовка строк для печати и экспорта из сырых записей changes."""

from __future__ import annotations

from typing import Any

from core import db
from core.class_levels import is_elementary_class, split_changes_by_level


def _shift_for_row(r: dict[str, Any]) -> int:
    cid = r.get("class_id")
    if cid:
        s = db.class_shift_by_id(int(cid))
        if s:
            return s
    name = (r.get("klass") or "").strip()
    if name:
        s = db.class_shift_by_name(name)
        if s:
            return s
    return 1


def lesson_range_for_shift(shift: int) -> tuple[int, int]:
    if shift == 2:
        return -1, 6
    return 0, 7


def validate_lesson(shift: int, lesson_no: int) -> bool:
    lo, hi = lesson_range_for_shift(shift)
    return lo <= lesson_no <= hi


def sort_key_teachers(r: dict[str, Any]) -> tuple[int, str]:
    return (int(r["lesson_no"]), (r.get("klass") or "").lower())


def build_teacher_rows(changes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for r in changes:
        rows.append(
            {
                "klass": r.get("klass") or "",
                "lesson_no": r.get("lesson_no"),
                "absent_fio": r.get("absent_fio") or "",
                "replacement_fio": r.get("replacement_fio") or "",
                "room": r.get("room") or "",
                "note": r.get("note") or "",
            }
        )
    rows.sort(key=sort_key_teachers)
    return rows


def build_student_rows(changes: list[dict[str, Any]]) -> tuple[list[dict], list[dict]]:
    """Две группы: смена 1 и смена 2. Только строки с предметом или кабинетом."""
    s1: list[dict[str, Any]] = []
    s2: list[dict[str, Any]] = []
    for r in changes:
        subj = (r.get("subject") or "").strip()
        room = (r.get("room") or "").strip()
        if not subj and not room:
            continue
        shift = _shift_for_row(r)
        row = {
            "klass": r.get("klass") or "",
            "lesson_no": r.get("lesson_no"),
            "subject": subj or "—",
            "room": room or "—",
            "note": r.get("note") or "",
        }
        if shift == 2:
            s2.append(row)
        else:
            s1.append(row)
    s1.sort(key=lambda x: (int(x["lesson_no"]), x["klass"].lower()))
    s2.sort(key=lambda x: (int(x["lesson_no"]), x["klass"].lower()))
    return s1, s2


def build_elementary_student_rows(changes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Начальная школа: одна таблица без деления на смены."""
    rows: list[dict[str, Any]] = []
    for r in changes:
        subj = (r.get("subject") or "").strip()
        room = (r.get("room") or "").strip()
        if not subj and not room:
            continue
        rows.append(
            {
                "klass": r.get("klass") or "",
                "lesson_no": r.get("lesson_no"),
                "subject": subj or "—",
                "room": room or "—",
                "note": r.get("note") or "",
            }
        )
    rows.sort(key=lambda x: (int(x["lesson_no"]), x["klass"].lower()))
    return rows


def changes_for_level(
    changes: list[dict[str, Any]], level: str
) -> list[dict[str, Any]]:
    elementary, main = split_changes_by_level(changes)
    if level == "elementary":
        return elementary
    return main


def xlsx_rows(changes: list[dict[str, Any]], date_iso: str) -> list[dict[str, Any]]:
    out = []
    for r in sorted(changes, key=sort_key_teachers):
        out.append(
            {
                "date": date_iso,
                "absent_fio": r.get("absent_fio") or "",
                "replacement_fio": r.get("replacement_fio") or "",
                "klass": r.get("klass") or "",
                "lesson_no": r.get("lesson_no"),
                "subject": r.get("subject") or "",
                "room": r.get("room") or "",
                "note": r.get("note") or "",
            }
        )
    return out
