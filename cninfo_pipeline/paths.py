from __future__ import annotations

import os
import sys
from pathlib import Path


APP_ID = "CNInfoReportCollector"


def resolve_app_data_dir() -> Path:
    if sys.platform == "win32":
        base_dir = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base_dir = Path.home() / "Library" / "Application Support"
    else:
        base_dir = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base_dir / APP_ID


def resolve_default_cache_dir() -> Path:
    return resolve_app_data_dir() / "cache"


def resolve_default_output_dir() -> Path:
    return resolve_app_data_dir() / "exports"


def ensure_writable_dir(path: str | Path) -> Path:
    resolved = Path(path).expanduser()
    resolved.mkdir(parents=True, exist_ok=True)

    probe_path = resolved / ".write_probe"
    probe_path.write_text("", encoding="utf-8")
    probe_path.unlink(missing_ok=True)
    return resolved
