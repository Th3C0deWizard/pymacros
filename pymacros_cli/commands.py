from __future__ import annotations

import os
import logging
import shutil
import shlex
import subprocess
from argparse import Namespace
from pathlib import Path
from urllib.parse import urlparse

import pymacros
from pymacros_cli import paths
from pymacros_cli.config import find_workbooks
from pymacros_cli.prompts import CliError, choose_searchable, confirm


logger = logging.getLogger(__name__)


def run_command(args: Namespace) -> int:
    workbook = _resolve_workbook(args.workbook)
    registry = pymacros.ProcedureRegistry(args.procedures_dir)
    procedure = _resolve_procedure(registry, args.procedure)

    logger.info(
        "Running procedure %s on workbook %s save=%s read_only=%s update_links=%s",
        procedure.path,
        workbook,
        args.save,
        args.read_only,
        args.update_links,
    )
    config = pymacros.ExcelConfig(visible=args.visible)
    pymacros.run_workbook(
        workbook,
        procedure.run,
        config=config,
        save=args.save,
        close_workbook=args.close_excel,
        close_excel=args.close_excel,
        read_only=args.read_only,
        update_links=args.update_links,
    )

    save_status = "saved" if args.save else "not saved"
    lifecycle_status = "Excel closed" if args.close_excel else "Excel kept open"
    logger.info("Procedure completed successfully")
    print(f"OK: ran {procedure.name} on {workbook} ({save_status}, {lifecycle_status}).")
    return 0


def procedures_list_command(args: Namespace) -> int:
    registry = pymacros.ProcedureRegistry(args.procedures_dir)
    logger.info("Listing procedures in %s", args.procedures_dir)
    procedures = registry.list()

    if not procedures:
        print(f"No procedures found in {Path(args.procedures_dir)}.")
        return 0

    for procedure in procedures:
        description = f" - {procedure.description}" if procedure.description else ""
        print(f"{procedure.path.stem}: {procedure.name}{description}")

    return 0


def procedures_create_command(args: Namespace) -> int:
    registry = pymacros.ProcedureRegistry(args.procedures_dir)
    logger.info("Creating procedure %s in %s", args.name, args.procedures_dir)
    path = registry.create(
        args.name,
        description=args.description,
        force=args.force,
    )

    print(f"Created procedure: {path}")
    return 0


def procedures_show_command(args: Namespace) -> int:
    registry = pymacros.ProcedureRegistry(args.procedures_dir)
    procedure = registry.get(args.name)

    print(f"Name: {procedure.name}")
    print(f"Description: {procedure.description or '(none)'}")
    print(f"Path: {procedure.path}")
    return 0


def procedures_rename_command(args: Namespace) -> int:
    registry = pymacros.ProcedureRegistry(args.procedures_dir)
    logger.info("Renaming procedure %s to %s", args.old_name, args.new_name)
    path = registry.rename(args.old_name, args.new_name, force=args.force)

    print(f"Renamed procedure: {path}")
    return 0


def procedures_delete_command(args: Namespace) -> int:
    registry = pymacros.ProcedureRegistry(args.procedures_dir)

    if not args.yes and not confirm(f"Delete procedure {args.name}?"):
        print("Cancelled.")
        return 0

    logger.info("Deleting procedure %s", args.name)
    path = registry.delete(args.name)
    print(f"Deleted procedure: {path}")
    return 0


def procedures_edit_command(args: Namespace) -> int:
    registry = pymacros.ProcedureRegistry(args.procedures_dir)
    path = registry.resolve(args.name)

    if not path.exists():
        raise pymacros.ProcedureNotFoundError(f"No existe el procedimiento: {path}")

    return _open_with_editor(path, editor=args.editor, fallback="notepad.exe")


def procedures_edit_folder_command(args: Namespace) -> int:
    procedures_dir = Path(args.procedures_dir)
    procedures_dir.mkdir(parents=True, exist_ok=True)
    return _open_with_editor(
        procedures_dir,
        editor=args.editor,
        fallback=_default_folder_editor(),
    )


def _resolve_workbook(workbook: str | None) -> str | Path:
    if workbook:
        if _is_web_url(workbook):
            return workbook

        path = Path(workbook)
        if not path.exists():
            raise CliError(f"Workbook not found: {path}")
        return path

    workbooks = find_workbooks()
    if not workbooks:
        return _prompt_workbook_path()

    return choose_searchable(
        "Select workbook:",
        workbooks,
        format_option=lambda path: str(path),
    )


def _resolve_procedure(
    registry: pymacros.ProcedureRegistry,
    procedure_name: str | None,
) -> pymacros.ProcedureInfo:
    if procedure_name:
        return registry.get(procedure_name)

    procedures = registry.list()
    procedure = choose_searchable(
        "Select procedure:",
        procedures,
        format_option=_format_procedure_choice,
    )
    return procedure


def _prompt_workbook_path() -> str | Path:
    print("No Excel workbooks found in the current directory.")
    raw_path = input("Workbook path: ").strip().strip('"')

    if not raw_path:
        raise CliError("Workbook path is required.")

    if _is_web_url(raw_path):
        return raw_path

    path = Path(raw_path)
    if not path.exists():
        raise CliError(f"Workbook not found: {path}")

    return path


def default_procedures_dir() -> str:
    return str(paths.procedures_dir())


def _is_web_url(path: str) -> bool:
    parsed = urlparse(path)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _format_procedure_choice(procedure: pymacros.ProcedureInfo) -> str:
    description = f" - {procedure.description}" if procedure.description else ""
    return f"{procedure.name} ({procedure.path.stem}){description}"


def _open_with_editor(path: Path, *, editor: str | None, fallback: str) -> int:
    command = editor or os.environ.get("EDITOR") or fallback

    if not command:
        raise CliError("Set EDITOR or pass --editor to open this path.")

    editor_command = _editor_command(command, path)
    logger.info("Opening %s with editor command %s", path, editor_command)
    return _call_editor(editor_command)


def _default_folder_editor() -> str:
    if shutil.which("code"):
        return "code"

    return "explorer.exe"


def _editor_command(command: str, path: Path) -> list[str]:
    try:
        parts = shlex.split(command)
    except ValueError as exc:
        raise CliError(f"Invalid editor command {command!r}: {exc}") from exc

    if not parts:
        raise CliError("Editor command cannot be empty.")

    executable = shutil.which(parts[0]) or parts[0]
    return [executable, *parts[1:], str(path)]


def _call_editor(command: list[str]) -> int:
    shell = Path(command[0]).suffix.lower() in {".bat", ".cmd"}
    return subprocess.call(command, shell=shell)
