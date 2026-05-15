from __future__ import annotations

import os
from pathlib import Path


APP_NAME = "pymacros"


def config_dir() -> Path:
    return _env_path("PYMACROS_CONFIG_HOME") or _appdata_dir("APPDATA", "Roaming")


def config_file() -> Path:
    return config_dir() / "config.json"


def data_dir() -> Path:
    return _env_path("PYMACROS_DATA_HOME") or _appdata_dir("LOCALAPPDATA", "Local")


def procedures_dir() -> Path:
    return config_dir() / "procedures"


def log_dir() -> Path:
    return _env_path("PYMACROS_LOG_HOME") or data_dir() / "logs"


def default_log_file() -> Path:
    return log_dir() / "pymacros.log"


def _env_path(name: str) -> Path | None:
    value = os.environ.get(name)
    return Path(value) if value else None


def _appdata_dir(env_name: str, fallback_kind: str) -> Path:
    value = os.environ.get(env_name)
    if value:
        return Path(value) / APP_NAME

    return Path.home() / "AppData" / fallback_kind / APP_NAME
