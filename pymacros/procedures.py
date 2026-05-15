from __future__ import annotations

import importlib.util
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Callable

from pymacros.context import ExcelContext
from pymacros.errors import InvalidProcedureError, ProcedureLoadError


Procedure = Callable[[ExcelContext], object]


@dataclass(frozen=True)
class ProcedureInfo:
    name: str
    description: str
    path: Path
    run: Procedure


def normalize_procedure_name(name: str) -> str:
    normalized = re.sub(r"[^0-9a-zA-Z_]+", "_", name.strip().lower())
    normalized = re.sub(r"_+", "_", normalized).strip("_")

    if not normalized:
        normalized = "procedure"

    if normalized[0].isdigit():
        normalized = f"procedure_{normalized}"

    return normalized


def procedure_filename(name: str) -> str:
    return f"{normalize_procedure_name(name)}.py"


def default_procedure_name(path: str | Path) -> str:
    return Path(path).stem.replace("_", " ").title()


def build_procedure_boilerplate(
    name: str,
    *,
    description: str = "Describe what this procedure does.",
) -> str:
    display_name = name.strip() or "New Procedure"

    return (
        "from pymacros import ExcelContext\n"
        "\n"
        "\n"
        f"NAME = {display_name!r}\n"
        f"DESCRIPTION = {description!r}\n"
        "\n"
        "\n"
        "def run(ctx: ExcelContext) -> None:\n"
        "    ws = ctx.active_sheet\n"
        "    # Write your Excel automation here.\n"
    )


def load_procedure(path: str | Path) -> ProcedureInfo:
    path = Path(path)

    if not path.exists():
        raise ProcedureLoadError(f"No existe el procedimiento: {path}")

    module_name = f"_pymacros_procedure_{path.stem}_{abs(hash(path.resolve()))}"
    spec = importlib.util.spec_from_file_location(module_name, path)

    if spec is None or spec.loader is None:
        raise ProcedureLoadError(f"No se pudo cargar el procedimiento: {path}")

    module = importlib.util.module_from_spec(spec)

    try:
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
    except Exception as exc:
        raise ProcedureLoadError(f"Error cargando el procedimiento {path}: {exc}") from exc
    finally:
        sys.modules.pop(module_name, None)

    return validate_procedure_module(module, path=path)


def validate_procedure_module(
    module: ModuleType,
    *,
    path: str | Path,
) -> ProcedureInfo:
    run = getattr(module, "run", None)

    if run is None:
        raise InvalidProcedureError("El procedimiento debe definir run(ctx).")

    if not callable(run):
        raise InvalidProcedureError("run debe ser callable.")

    path = Path(path)

    return ProcedureInfo(
        name=getattr(module, "NAME", None) or default_procedure_name(path),
        description=getattr(module, "DESCRIPTION", None) or "",
        path=path,
        run=run,
    )
