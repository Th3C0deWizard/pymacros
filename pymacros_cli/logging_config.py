from __future__ import annotations

import logging
from pathlib import Path

from pymacros_cli import paths


LOG_FORMAT = "%(asctime)s %(levelname)s [%(name)s] %(message)s"


def configure_logging(*, debug: bool = False, log_file: str | Path | None = None) -> None:
    handlers: list[logging.Handler] = []
    log_path = Path(log_file) if log_file else paths.default_log_file()
    log_path.parent.mkdir(parents=True, exist_ok=True)
    handlers.append(logging.FileHandler(log_path, encoding="utf-8"))

    if debug:
        handlers.append(logging.StreamHandler())

    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format=LOG_FORMAT,
        handlers=handlers,
        force=True,
    )
