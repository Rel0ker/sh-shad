"""Форматирование дат для заголовков."""

from __future__ import annotations

from datetime import datetime

_MONTHS = (
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
)


def format_date_ru(iso_date: str) -> str:
    d = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return f"{d.day} {_MONTHS[d.month - 1]} {d.year} г."
