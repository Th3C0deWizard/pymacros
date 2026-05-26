from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from pymacros_cli import paths

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
IGNORED_WORKBOOK_DIRS = {
    ".git",
    ".mypy_cache",
    ".pytest_cache",
    ".ruff_cache",
    "__pycache__",
    "build",
    "dist",
    "venv",
    ".venv",
}


class ConfigError(Exception):
    pass


@dataclass(frozen=True)
class CliConfig:
    procedures_dir: Path = field(default_factory=paths.procedures_dir)
    visible: bool = True
    save: bool = True
    close_excel: bool = True
    read_only: bool = False
    update_links: int = 0
    log_file: Path = field(default_factory=paths.default_log_file)
    editor: str | None = None


def load_config(path: str | Path | None = None) -> CliConfig:
    config_path = Path(path) if path else paths.config_file()

    if not config_path.exists():
        return CliConfig()

    try:
        raw_config = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Invalid JSON config file {config_path}: {exc}") from exc

    if not isinstance(raw_config, dict):
        raise ConfigError(f"Config file must contain a JSON object: {config_path}")

    return CliConfig(
        procedures_dir=_path_value(raw_config, "procedures_dir", paths.procedures_dir()),
        visible=_bool_value(raw_config, "visible", True),
        save=_bool_value(raw_config, "save", True),
        close_excel=_bool_value(raw_config, "close_excel", True),
        read_only=_bool_value(raw_config, "read_only", False),
        update_links=_int_value(raw_config, "update_links", 0),
        log_file=_path_value(raw_config, "log_file", paths.default_log_file()),
        editor=_optional_string_value(raw_config, "editor"),
    )


def find_workbooks(root: str | Path = ".") -> list[Path]:
    root = Path(root)
    return sorted(
        path
        for path in root.rglob("*")
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS
        if not any(part in IGNORED_WORKBOOK_DIRS for part in path.relative_to(root).parts[:-1])
    )


def _path_value(config: dict, key: str, default: Path) -> Path:
    value = config.get(key)
    return Path(value) if value else default


def _bool_value(config: dict, key: str, default: bool) -> bool:
    value = config.get(key, default)

    if isinstance(value, bool):
        return value

    raise ConfigError(f"Config value {key!r} must be a boolean.")


def _int_value(config: dict, key: str, default: int) -> int:
    value = config.get(key, default)

    if isinstance(value, int) and not isinstance(value, bool):
        return value

    raise ConfigError(f"Config value {key!r} must be an integer.")


def _optional_string_value(config: dict, key: str) -> str | None:
    value = config.get(key)

    if value is None or value == "":
        return None

    if isinstance(value, str):
        return value

    raise ConfigError(f"Config value {key!r} must be a string.")
