from __future__ import annotations

import argparse
import logging
from collections.abc import Sequence

import pymacros
from pymacros_cli import commands
from pymacros_cli.config import CliConfig, ConfigError, load_config
from pymacros_cli.logging_config import configure_logging
from pymacros_cli.prompts import CliError


logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pymacros",
        description="Run Python procedures against Excel workbooks.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="pymacros CLI 0.1.2",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging and tracebacks for handled errors.",
    )
    parser.add_argument(
        "--log-file",
        help="Write logs to a file.",
    )
    parser.add_argument(
        "--config-file",
        help="Read CLI defaults from this JSON config file.",
    )
    subparsers = parser.add_subparsers(dest="command")

    run_parser = subparsers.add_parser(
        "run",
        help="Run a Python procedure on an Excel workbook.",
    )
    run_parser.add_argument(
        "workbook",
        nargs="?",
        help="Workbook path. If omitted, choose from discovered Excel files.",
    )
    run_parser.add_argument(
        "procedure",
        nargs="?",
        help="Procedure name. If omitted, choose from available procedures.",
    )
    run_parser.add_argument(
        "--procedures-dir",
        default=None,
        help="Directory that contains procedure .py files.",
    )
    save_group = run_parser.add_mutually_exclusive_group()
    save_group.add_argument("--save", action="store_true", help="Save workbook changes.")
    save_group.add_argument(
        "--no-save",
        action="store_false",
        dest="save",
        help="Do not save workbook changes.",
    )
    run_parser.set_defaults(save=None)
    lifecycle_group = run_parser.add_mutually_exclusive_group()
    lifecycle_group.add_argument(
        "--close-excel",
        action="store_true",
        dest="close_excel",
        help="Close the workbook and Excel process after running.",
    )
    lifecycle_group.add_argument(
        "--keep-open",
        action="store_false",
        dest="close_excel",
        help="Leave the workbook and Excel process open after running.",
    )
    run_parser.set_defaults(close_excel=None)
    visibility_group = run_parser.add_mutually_exclusive_group()
    visibility_group.add_argument(
        "--visible",
        action="store_true",
        dest="visible",
        help="Show Excel while running.",
    )
    visibility_group.add_argument(
        "--hidden",
        action="store_false",
        dest="visible",
        help="Hide Excel while running.",
    )
    run_parser.set_defaults(visible=None)
    run_parser.add_argument(
        "--read-only",
        action="store_true",
        default=None,
        help="Open the workbook read-only.",
    )
    run_parser.add_argument(
        "--read-write",
        action="store_false",
        dest="read_only",
        help="Open the workbook read/write.",
    )
    run_parser.add_argument(
        "--update-links",
        type=int,
        default=None,
        help="Excel UpdateLinks value used when opening the workbook.",
    )
    run_parser.set_defaults(func=commands.run_command)

    procedures_parser = subparsers.add_parser(
        "procedures",
        help="Manage Python procedure files.",
    )
    procedure_subparsers = procedures_parser.add_subparsers(dest="procedure_command")

    list_parser = procedure_subparsers.add_parser("list", help="List procedures.")
    _add_procedures_dir_argument(list_parser)
    list_parser.set_defaults(func=commands.procedures_list_command)

    create_parser = procedure_subparsers.add_parser("create", help="Create a procedure.")
    create_parser.add_argument("name")
    _add_procedures_dir_argument(create_parser)
    create_parser.add_argument(
        "--description",
        default="Describe what this procedure does.",
    )
    create_parser.add_argument("--force", action="store_true")
    create_parser.set_defaults(func=commands.procedures_create_command)

    show_parser = procedure_subparsers.add_parser("show", help="Show procedure details.")
    show_parser.add_argument("name")
    _add_procedures_dir_argument(show_parser)
    show_parser.set_defaults(func=commands.procedures_show_command)

    edit_parser = procedure_subparsers.add_parser("edit", help="Open a procedure in an editor.")
    edit_parser.add_argument("name")
    _add_procedures_dir_argument(edit_parser)
    edit_parser.add_argument("--editor")
    edit_parser.set_defaults(func=commands.procedures_edit_command)

    edit_folder_parser = procedure_subparsers.add_parser(
        "edit-folder",
        help="Open the procedures folder in an editor.",
    )
    _add_procedures_dir_argument(edit_folder_parser)
    edit_folder_parser.add_argument("--editor")
    edit_folder_parser.set_defaults(func=commands.procedures_edit_folder_command)

    rename_parser = procedure_subparsers.add_parser("rename", help="Rename a procedure.")
    rename_parser.add_argument("old_name")
    rename_parser.add_argument("new_name")
    _add_procedures_dir_argument(rename_parser)
    rename_parser.add_argument("--force", action="store_true")
    rename_parser.set_defaults(func=commands.procedures_rename_command)

    delete_parser = procedure_subparsers.add_parser("delete", help="Delete a procedure.")
    delete_parser.add_argument("name")
    _add_procedures_dir_argument(delete_parser)
    delete_parser.add_argument("--yes", action="store_true", help="Skip confirmation.")
    delete_parser.set_defaults(func=commands.procedures_delete_command)

    return parser


def _add_procedures_dir_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--procedures-dir",
        default=None,
        help="Directory that contains procedure .py files.",
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config_file)
        _apply_config_defaults(args, config)
    except ConfigError as exc:
        configure_logging(debug=args.debug, log_file=args.log_file)
        logger.error("Could not load config: %s", exc, exc_info=args.debug)
        parser.exit(1, f"Error: {exc}\n")

    configure_logging(debug=args.debug, log_file=args.log_file)
    logger.debug("Parsed CLI args: %s", args)

    if not hasattr(args, "func"):
        parser.print_help()
        return 0

    try:
        return args.func(args)
    except (CliError, pymacros.PymacrosError) as exc:
        logger.error("Command failed: %s", exc, exc_info=args.debug)
        parser.exit(1, f"Error: {exc}\n")
    except Exception as exc:
        logger.exception("Unexpected command failure")

        if args.debug:
            parser.exit(1, f"Unexpected error: {exc}\n")

        parser.exit(1, "Unexpected error. Re-run with --debug for details.\n")


def _apply_config_defaults(args: argparse.Namespace, config: CliConfig) -> None:
    if getattr(args, "log_file", None) is None:
        args.log_file = str(config.log_file)

    if hasattr(args, "procedures_dir") and args.procedures_dir is None:
        args.procedures_dir = str(config.procedures_dir)

    if hasattr(args, "save") and args.save is None:
        args.save = config.save

    if hasattr(args, "close_excel") and args.close_excel is None:
        args.close_excel = config.close_excel

    if hasattr(args, "visible") and args.visible is None:
        args.visible = config.visible

    if hasattr(args, "read_only") and args.read_only is None:
        args.read_only = config.read_only

    if hasattr(args, "update_links") and args.update_links is None:
        args.update_links = config.update_links

    if hasattr(args, "editor") and args.editor is None:
        args.editor = config.editor
