# Ранняя установка каталога браузеров (до импорта приложения)
import os
import sys
from pathlib import Path

_APP = "ScheduleChanges"


def _browsers_dir() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path.home() / ".local" / "share"
    return base / _APP / "ms-playwright"


try:
    p = _browsers_dir()
    p.mkdir(parents=True, exist_ok=True)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(p.resolve())
except OSError:
    pass
