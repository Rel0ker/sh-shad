"""Начальная школа: 1–4 класс; основная: 5 и выше."""

from __future__ import annotations

import re
from typing import Any

_GRADE_RE = re.compile(r"^(\d{1,2})")


def parse_class_grade(klass: str) -> int | None:
    s = (klass or "").strip().replace(" ", "")
    if not s:
        return None
    m = _GRADE_RE.match(s)
    if not m:
        return None
    return int(m.group(1))


def is_elementary_class(klass: str) -> bool:
    g = parse_class_grade(klass)
    return g is not None and 1 <= g <= 4


def split_changes_by_level(
    changes: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    elementary: list[dict[str, Any]] = []
    main: list[dict[str, Any]] = []
    for r in changes:
        if is_elementary_class(r.get("klass") or ""):
            elementary.append(r)
        else:
            main.append(r)
    return elementary, main
