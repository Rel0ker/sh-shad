"""Экспорт учётного XLSX."""

from __future__ import annotations

import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def build_changes_workbook(rows: list[dict[str, Any]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Замены"

    headers = [
        "Дата",
        "Отсутствующий",
        "Заменяющий",
        "Класс",
        "№ урока",
        "Предмет",
        "Кабинет",
        "Примечание",
    ]
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDDDDD")

    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, r in enumerate(rows, start=2):
        ws.cell(row=r_idx, column=1, value=r.get("date"))
        ws.cell(row=r_idx, column=2, value=r.get("absent_fio"))
        ws.cell(row=r_idx, column=3, value=r.get("replacement_fio"))
        ws.cell(row=r_idx, column=4, value=r.get("klass"))
        ws.cell(row=r_idx, column=5, value=r.get("lesson_no"))
        ws.cell(row=r_idx, column=6, value=r.get("subject"))
        ws.cell(row=r_idx, column=7, value=r.get("room"))
        ws.cell(row=r_idx, column=8, value=r.get("note"))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"

    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
