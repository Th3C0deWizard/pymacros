from __future__ import annotations

from pathlib import Path

from pymacros.errors import ProcedureExistsError, ProcedureNotFoundError
from pymacros.procedures import (
    ProcedureInfo,
    build_procedure_boilerplate,
    load_procedure,
    procedure_filename,
)


class ProcedureRegistry:
    def __init__(self, root: str | Path):
        self.root = Path(root)

    def path_for(self, name: str) -> Path:
        return self.root / procedure_filename(name)

    def resolve(self, name: str | Path) -> Path:
        path = Path(name)

        if path.suffix == ".py":
            if path.is_absolute() or path.parent != Path("."):
                return path

            return self.root / path

        return self.path_for(str(name))

    def get(self, name: str | Path) -> ProcedureInfo:
        path = self.resolve(name)

        if not path.exists():
            raise ProcedureNotFoundError(f"No existe el procedimiento: {path}")

        return load_procedure(path)

    def list(self) -> list[ProcedureInfo]:
        if not self.root.exists():
            return []

        return [load_procedure(path) for path in sorted(self.root.glob("*.py"))]

    def create(
        self,
        name: str,
        *,
        description: str = "Describe what this procedure does.",
        force: bool = False,
    ) -> Path:
        path = self.path_for(name)

        if path.exists() and not force:
            raise ProcedureExistsError(f"El procedimiento ya existe: {path}")

        self.root.mkdir(parents=True, exist_ok=True)
        path.write_text(
            build_procedure_boilerplate(name, description=description),
            encoding="utf-8",
        )

        return path

    def delete(self, name: str) -> Path:
        path = self.path_for(name)

        if not path.exists():
            raise ProcedureNotFoundError(f"No existe el procedimiento: {path}")

        path.unlink()
        return path

    def rename(self, old_name: str, new_name: str, *, force: bool = False) -> Path:
        old_path = self.path_for(old_name)
        new_path = self.path_for(new_name)

        if not old_path.exists():
            raise ProcedureNotFoundError(f"No existe el procedimiento: {old_path}")

        if new_path.exists() and not force:
            raise ProcedureExistsError(f"El procedimiento ya existe: {new_path}")

        old_path.rename(new_path)
        return new_path
