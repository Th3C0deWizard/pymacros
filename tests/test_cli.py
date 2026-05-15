import json

import pytest

import pymacros_cli.commands as commands
from pymacros_cli import paths
from pymacros_cli.config import ConfigError, find_workbooks, load_config
from pymacros_cli.prompts import CliError, choose_one, choose_searchable, confirm
from pymacros_cli.app import build_parser, main


@pytest.fixture(autouse=True)
def isolate_app_paths(monkeypatch, tmp_path):
    monkeypatch.setenv("PYMACROS_CONFIG_HOME", str(tmp_path / "config"))
    monkeypatch.setenv("PYMACROS_DATA_HOME", str(tmp_path / "data"))
    monkeypatch.setenv("PYMACROS_LOG_HOME", str(tmp_path / "logs"))


def test_cli_parser_uses_pymacros_program_name():
    parser = build_parser()

    assert parser.prog == "pymacros"


def test_cli_main_prints_help_without_command(capsys):
    result = main([])
    captured = capsys.readouterr()

    assert result == 0
    assert "Run Python procedures against Excel workbooks." in captured.out


def test_cli_procedure_crud_commands(tmp_path, capsys):
    procedures_dir = tmp_path / "procedures"

    assert main([
        "procedures",
        "create",
        "My Report",
        "--procedures-dir",
        str(procedures_dir),
        "--description",
        "Formats report",
    ]) == 0
    assert (procedures_dir / "my_report.py").exists()

    assert main([
        "procedures",
        "list",
        "--procedures-dir",
        str(procedures_dir),
    ]) == 0
    captured = capsys.readouterr()
    assert "my_report: My Report - Formats report" in captured.out

    assert main([
        "procedures",
        "show",
        "My Report",
        "--procedures-dir",
        str(procedures_dir),
    ]) == 0
    captured = capsys.readouterr()
    assert "Name: My Report" in captured.out
    assert "Description: Formats report" in captured.out

    assert main([
        "procedures",
        "rename",
        "My Report",
        "Daily Report",
        "--procedures-dir",
        str(procedures_dir),
    ]) == 0
    assert not (procedures_dir / "my_report.py").exists()
    assert (procedures_dir / "daily_report.py").exists()

    assert main([
        "procedures",
        "delete",
        "Daily Report",
        "--procedures-dir",
        str(procedures_dir),
        "--yes",
    ]) == 0
    assert not (procedures_dir / "daily_report.py").exists()


def test_cli_run_delegates_to_pymacros(monkeypatch, tmp_path, capsys):
    calls = []
    workbook = tmp_path / "book.xlsx"
    procedures_dir = tmp_path / "procedures"
    workbook.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )

    def fake_run_workbook(
        path,
        procedure,
        *,
        config,
        save,
        close_workbook,
        close_excel,
        read_only,
        update_links,
    ):
        calls.append((path, procedure, config.visible, save, close_workbook, close_excel, read_only, update_links))

    monkeypatch.setattr(commands.pymacros, "run_workbook", fake_run_workbook)

    result = main([
        "run",
        str(workbook),
        "My Report",
        "--procedures-dir",
        str(procedures_dir),
        "--save",
        "--hidden",
        "--read-only",
        "--update-links",
        "3",
    ])
    captured = capsys.readouterr()

    assert result == 0
    assert len(calls) == 1
    assert calls[0][0] == workbook
    assert calls[0][2:] == (False, True, True, True, True, 3)
    assert "OK: ran My Report" in captured.out
    assert "Excel closed" in captured.out


def test_cli_run_keep_open_leaves_workbook_and_excel_open(monkeypatch, tmp_path, capsys):
    calls = []
    workbook = tmp_path / "book.xlsx"
    procedures_dir = tmp_path / "procedures"
    workbook.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )

    def fake_run_workbook(path, procedure, **kwargs):
        calls.append((path, kwargs))

    monkeypatch.setattr(commands.pymacros, "run_workbook", fake_run_workbook)

    result = main([
        "run",
        str(workbook),
        "My Report",
        "--procedures-dir",
        str(procedures_dir),
        "--keep-open",
    ])
    captured = capsys.readouterr()

    assert result == 0
    assert calls[0][1]["close_workbook"] is False
    assert calls[0][1]["close_excel"] is False
    assert "Excel kept open" in captured.out


def test_cli_run_prompts_for_workbook_and_procedure(monkeypatch, tmp_path, capsys):
    calls = []
    procedures_dir = tmp_path / "procedures"
    workbook_a = tmp_path / "a.xlsx"
    workbook_b = tmp_path / "nested" / "b.xlsm"
    workbook_a.write_text("fake", encoding="utf-8")
    workbook_b.parent.mkdir()
    workbook_b.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "first.py").write_text(
        "NAME = 'First'\n"
        "DESCRIPTION = 'First description'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    (procedures_dir / "second.py").write_text(
        "NAME = 'Second'\n"
        "DESCRIPTION = 'Second description'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    choices = iter(["", "2", "", "1"])
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("builtins.input", lambda prompt: next(choices))
    monkeypatch.setattr(
        commands.pymacros,
        "run_workbook",
        lambda path, procedure, **kwargs: calls.append((path, procedure, kwargs)),
    )

    result = main([
        "run",
        "--procedures-dir",
        str(procedures_dir),
    ])
    captured = capsys.readouterr()

    assert result == 0
    assert calls[0][0] == workbook_b.relative_to(tmp_path)
    assert "Select workbook: Type search text to filter results" in captured.out
    assert "1. a.xlsx" in captured.out
    assert "2. nested" in captured.out
    assert "Select procedure: Type search text to filter results" in captured.out
    assert "First (first) - First description" in captured.out
    assert "Second (second) - Second description" in captured.out


def test_cli_run_search_filters_workbook_and_procedure(monkeypatch, tmp_path, capsys):
    calls = []
    procedures_dir = tmp_path / "procedures"
    workbook_a = tmp_path / "sales.xlsx"
    workbook_b = tmp_path / "inventory.xlsx"
    workbook_a.write_text("fake", encoding="utf-8")
    workbook_b.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "format_sales.py").write_text(
        "NAME = 'Format Sales'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    (procedures_dir / "format_inventory.py").write_text(
        "NAME = 'Format Inventory'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    choices = iter(["inventory", "1", "sales", "1"])
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("builtins.input", lambda prompt: next(choices))
    monkeypatch.setattr(
        commands.pymacros,
        "run_workbook",
        lambda path, procedure, **kwargs: calls.append((path, procedure, kwargs)),
    )

    result = main([
        "run",
        "--procedures-dir",
        str(procedures_dir),
    ])
    captured = capsys.readouterr()

    assert result == 0
    assert calls[0][0] == workbook_b.relative_to(tmp_path)
    assert "Found 1 match(es):" in captured.out
    assert "inventory.xlsx" in captured.out
    assert "Format Sales (format_sales)" in captured.out


def test_cli_run_prompts_even_with_single_workbook_and_procedure(monkeypatch, tmp_path, capsys):
    calls = []
    workbook = tmp_path / "only.xlsx"
    procedures_dir = tmp_path / "procedures"
    workbook.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "only.py").write_text(
        "NAME = 'Only Procedure'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    choices = iter(["", "1", "", "1"])
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr("builtins.input", lambda prompt: next(choices))
    monkeypatch.setattr(
        commands.pymacros,
        "run_workbook",
        lambda path, procedure, **kwargs: calls.append((path, procedure, kwargs)),
    )

    result = main([
        "run",
        "--procedures-dir",
        str(procedures_dir),
    ])
    captured = capsys.readouterr()

    assert result == 0
    assert calls[0][0] == workbook.relative_to(tmp_path)
    assert "Select workbook: Type search text to filter results" in captured.out
    assert "Found 1 match(es):" in captured.out
    assert "1. only.xlsx" in captured.out
    assert "Select procedure: Type search text to filter results" in captured.out
    assert "1. Only Procedure (only)" in captured.out


def test_cli_run_prompts_for_manual_workbook_path_when_none_found(monkeypatch, tmp_path):
    calls = []
    workbook = tmp_path / "manual.xlsx"
    procedures_dir = tmp_path / "procedures"
    workbook.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    choices = iter([str(workbook), "", "1"])
    monkeypatch.chdir(tmp_path / "procedures")
    monkeypatch.setattr("builtins.input", lambda prompt: next(choices))
    monkeypatch.setattr(
        commands.pymacros,
        "run_workbook",
        lambda path, procedure, **kwargs: calls.append((path, procedure, kwargs)),
    )

    result = main([
        "run",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert calls[0][0] == workbook


def test_cli_run_reports_missing_workbook(capsys):
    try:
        main(["run", "missing.xlsx", "anything"])
    except SystemExit as exc:
        assert exc.code == 1

    captured = capsys.readouterr()
    assert "Workbook not found" in captured.err


def test_cli_debug_logs_traceback_for_handled_error(capsys):
    try:
        main(["--debug", "run", "missing.xlsx", "anything"])
    except SystemExit as exc:
        assert exc.code == 1

    captured = capsys.readouterr()
    assert "Traceback" in captured.err
    assert "Workbook not found" in captured.err


def test_cli_writes_logs_to_file(monkeypatch, tmp_path):
    workbook = tmp_path / "book.xlsx"
    procedures_dir = tmp_path / "procedures"
    log_file = tmp_path / "pymacros.log"
    workbook.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )

    monkeypatch.setattr(commands.pymacros, "run_workbook", lambda *args, **kwargs: None)

    result = main([
        "--log-file",
        str(log_file),
        "run",
        str(workbook),
        "My Report",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert "Running procedure" in log_file.read_text(encoding="utf-8")


def test_cli_writes_logs_to_windows_default_location(monkeypatch, tmp_path):
    workbook = tmp_path / "book.xlsx"
    procedures_dir = tmp_path / "procedures"
    workbook.write_text("fake", encoding="utf-8")
    procedures_dir.mkdir()
    (procedures_dir / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )

    monkeypatch.setattr(commands.pymacros, "run_workbook", lambda *args, **kwargs: None)

    result = main([
        "run",
        str(workbook),
        "My Report",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert "Running procedure" in paths.default_log_file().read_text(encoding="utf-8")


def test_windows_default_paths_use_appdata_env(monkeypatch, tmp_path):
    monkeypatch.delenv("PYMACROS_CONFIG_HOME", raising=False)
    monkeypatch.delenv("PYMACROS_DATA_HOME", raising=False)
    monkeypatch.delenv("PYMACROS_LOG_HOME", raising=False)
    monkeypatch.setenv("APPDATA", str(tmp_path / "Roaming"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "Local"))

    assert paths.config_file() == tmp_path / "Roaming" / "pymacros" / "config.json"
    assert paths.procedures_dir() == tmp_path / "Roaming" / "pymacros" / "procedures"
    assert paths.default_log_file() == tmp_path / "Local" / "pymacros" / "logs" / "pymacros.log"


def test_load_config_reads_json_values(tmp_path):
    config_file = tmp_path / "config.json"
    config_file.write_text(
        json.dumps({
            "procedures_dir": str(tmp_path / "procedures"),
            "visible": False,
            "save": True,
            "close_excel": False,
            "read_only": True,
            "update_links": 2,
            "log_file": str(tmp_path / "app.log"),
            "editor": "code",
        }),
        encoding="utf-8",
    )

    config = load_config(config_file)

    assert config.procedures_dir == tmp_path / "procedures"
    assert config.visible is False
    assert config.save is True
    assert config.close_excel is False
    assert config.read_only is True
    assert config.update_links == 2
    assert config.log_file == tmp_path / "app.log"
    assert config.editor == "code"


def test_load_config_rejects_invalid_json_value(tmp_path):
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps({"visible": "no"}), encoding="utf-8")

    with pytest.raises(ConfigError):
        load_config(config_file)


def test_load_config_rejects_invalid_editor_value(tmp_path):
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps({"editor": 123}), encoding="utf-8")

    with pytest.raises(ConfigError):
        load_config(config_file)


def test_cli_uses_config_defaults_and_allows_cli_override(monkeypatch, tmp_path):
    calls = []
    workbook = tmp_path / "book.xlsx"
    config_procedures = tmp_path / "config_procedures"
    cli_procedures = tmp_path / "cli_procedures"
    config_file = tmp_path / "config.json"
    workbook.write_text("fake", encoding="utf-8")
    config_procedures.mkdir()
    cli_procedures.mkdir()
    (config_procedures / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    (cli_procedures / "my_report.py").write_text(
        "NAME = 'My Report'\n"
        "\n"
        "def run(ctx):\n"
        "    return None\n",
        encoding="utf-8",
    )
    config_file.write_text(
        json.dumps({
            "procedures_dir": str(config_procedures),
            "visible": False,
            "save": True,
            "close_excel": False,
            "read_only": True,
            "update_links": 7,
        }),
        encoding="utf-8",
    )

    def fake_run_workbook(path, procedure, *, config, save, close_workbook, close_excel, read_only, update_links):
        calls.append((path, config.visible, save, close_workbook, close_excel, read_only, update_links))

    monkeypatch.setattr(commands.pymacros, "run_workbook", fake_run_workbook)

    assert main([
        "--config-file",
        str(config_file),
        "run",
        str(workbook),
        "My Report",
    ]) == 0
    assert calls[-1] == (workbook, False, True, False, False, True, 7)

    assert main([
        "--config-file",
        str(config_file),
        "run",
        str(workbook),
        "My Report",
        "--procedures-dir",
        str(cli_procedures),
        "--no-save",
        "--visible",
        "--read-write",
        "--close-excel",
        "--update-links",
        "3",
    ]) == 0
    assert calls[-1] == (workbook, True, False, True, True, False, 3)


def test_cli_delete_can_be_cancelled(monkeypatch, tmp_path, capsys):
    procedures_dir = tmp_path / "procedures"
    main([
        "procedures",
        "create",
        "Keep Me",
        "--procedures-dir",
        str(procedures_dir),
    ])
    monkeypatch.setattr(commands, "confirm", lambda message: False)

    result = main([
        "procedures",
        "delete",
        "Keep Me",
        "--procedures-dir",
        str(procedures_dir),
    ])
    captured = capsys.readouterr()

    assert result == 0
    assert "Cancelled." in captured.out
    assert (procedures_dir / "keep_me.py").exists()


def test_cli_edit_uses_requested_editor(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    main([
        "procedures",
        "create",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
    ])
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "procedures",
        "edit",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
        "--editor",
        "fake-editor",
    ])

    assert result == 0
    assert calls == [(["fake-editor", str(procedures_dir / "edit_me.py")], {"shell": False})]


def test_cli_edit_defaults_to_notepad_instead_of_file_association(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    main([
        "procedures",
        "create",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
    ])
    monkeypatch.delenv("EDITOR", raising=False)
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "procedures",
        "edit",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert len(calls) == 1
    assert calls[0][0][0].lower().endswith("notepad.exe")
    assert calls[0][0][1:] == [str(procedures_dir / "edit_me.py")]
    assert calls[0][1] == {"shell": False}


def test_cli_edit_folder_opens_procedures_dir_with_editor(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "procedures",
        "edit-folder",
        "--procedures-dir",
        str(procedures_dir),
        "--editor",
        "fake-editor",
    ])

    assert result == 0
    assert procedures_dir.exists()
    assert calls == [(["fake-editor", str(procedures_dir)], {"shell": False})]


def test_cli_edit_folder_defaults_to_code_when_available(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    monkeypatch.delenv("EDITOR", raising=False)
    monkeypatch.setattr(commands.shutil, "which", lambda command: "code.cmd" if command == "code" else None)
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "procedures",
        "edit-folder",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert calls == [(["code.cmd", str(procedures_dir)], {"shell": True})]


def test_cli_edit_uses_configured_editor(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps({"editor": "configured-editor"}), encoding="utf-8")
    main([
        "procedures",
        "create",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
    ])
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "--config-file",
        str(config_file),
        "procedures",
        "edit",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert calls == [(["configured-editor", str(procedures_dir / "edit_me.py")], {"shell": False})]


def test_cli_editor_argument_overrides_configured_editor(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps({"editor": "configured-editor"}), encoding="utf-8")
    main([
        "procedures",
        "create",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
    ])
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "--config-file",
        str(config_file),
        "procedures",
        "edit",
        "Edit Me",
        "--procedures-dir",
        str(procedures_dir),
        "--editor",
        "cli-editor",
    ])

    assert result == 0
    assert calls == [(["cli-editor", str(procedures_dir / "edit_me.py")], {"shell": False})]


def test_cli_edit_folder_uses_configured_editor(monkeypatch, tmp_path):
    calls = []
    procedures_dir = tmp_path / "procedures"
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps({"editor": "configured-editor"}), encoding="utf-8")
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = main([
        "--config-file",
        str(config_file),
        "procedures",
        "edit-folder",
        "--procedures-dir",
        str(procedures_dir),
    ])

    assert result == 0
    assert calls == [(["configured-editor", str(procedures_dir)], {"shell": False})]


def test_editor_command_resolves_code_cmd_and_uses_shell(monkeypatch, tmp_path):
    calls = []
    procedure_path = tmp_path / "procedure.py"
    procedure_path.write_text("def run(ctx):\n    pass\n", encoding="utf-8")
    monkeypatch.setattr(commands.shutil, "which", lambda command: "C:/bin/code.cmd" if command == "code" else None)
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = commands._open_with_editor(procedure_path, editor="code", fallback="notepad.exe")

    assert result == 0
    assert calls == [(["C:/bin/code.cmd", str(procedure_path)], {"shell": True})]


def test_editor_command_preserves_editor_arguments(monkeypatch, tmp_path):
    calls = []
    procedure_path = tmp_path / "procedure.py"
    procedure_path.write_text("def run(ctx):\n    pass\n", encoding="utf-8")
    monkeypatch.setattr(commands.shutil, "which", lambda command: "C:/bin/code.cmd" if command == "code" else None)
    monkeypatch.setattr(commands.subprocess, "call", lambda command, **kwargs: calls.append((command, kwargs)) or 0)

    result = commands._open_with_editor(procedure_path, editor="code --reuse-window", fallback="notepad.exe")

    assert result == 0
    assert calls == [(["C:/bin/code.cmd", "--reuse-window", str(procedure_path)], {"shell": True})]


def test_find_workbooks_returns_excel_files_only(tmp_path):
    xlsx = tmp_path / "a.xlsx"
    xlsm = tmp_path / "nested" / "b.xlsm"
    ignored = tmp_path / "venv" / "ignored.xlsx"
    txt = tmp_path / "notes.txt"
    xlsx.write_text("", encoding="utf-8")
    xlsm.parent.mkdir()
    xlsm.write_text("", encoding="utf-8")
    ignored.parent.mkdir()
    ignored.write_text("", encoding="utf-8")
    txt.write_text("", encoding="utf-8")

    assert find_workbooks(tmp_path) == [xlsx, xlsm]


def test_choose_one_accepts_numbered_input(monkeypatch, capsys):
    monkeypatch.setattr("builtins.input", lambda prompt: "2")

    selected = choose_one("Select item:", ["one", "two"])
    captured = capsys.readouterr()

    assert selected == "two"
    assert "Select item:" in captured.out


def test_choose_one_formats_single_option(capsys):
    selected = choose_one(
        "Select item:",
        ["one"],
        format_option=lambda value: f"formatted {value}",
    )
    captured = capsys.readouterr()

    assert selected == "one"
    assert "Select item: formatted one" in captured.out


def test_choose_searchable_filters_options(monkeypatch, capsys):
    choices = iter(["tw", "1"])
    monkeypatch.setattr("builtins.input", lambda prompt: next(choices))

    selected = choose_searchable("Select item:", ["one", "two"])
    captured = capsys.readouterr()

    assert selected == "two"
    assert "Found 1 match(es):" in captured.out
    assert "1. two" in captured.out


def test_choose_searchable_prompts_for_single_option(monkeypatch, capsys):
    choices = iter(["", "1"])
    monkeypatch.setattr("builtins.input", lambda prompt: next(choices))

    selected = choose_searchable("Select item:", ["one"])
    captured = capsys.readouterr()

    assert selected == "one"
    assert "Select item: Type search text to filter results" in captured.out
    assert "1. one" in captured.out


def test_choose_one_rejects_bad_input(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda prompt: "bad")

    try:
        choose_one("Select item:", ["one", "two"])
    except CliError as exc:
        assert "numero" in str(exc)
    else:
        raise AssertionError("Expected CliError")


def test_confirm_accepts_yes(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda prompt: "yes")

    assert confirm("Continue?") is True
