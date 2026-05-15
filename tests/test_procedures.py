from types import ModuleType

import pytest

from pymacros import (
    InvalidProcedureError,
    ProcedureLoadError,
    build_procedure_boilerplate,
    load_procedure,
    normalize_procedure_name,
    procedure_filename,
    validate_procedure_module,
)


def test_normalize_procedure_name_makes_safe_python_stem():
    assert normalize_procedure_name(" 123 Report! Final ") == "procedure_123_report_final"
    assert normalize_procedure_name("") == "procedure"
    assert procedure_filename("My Report") == "my_report.py"


def test_boilerplate_is_valid_python():
    source = build_procedure_boilerplate("My Procedure")

    compile(source, "procedure.py", "exec")
    assert "def run(ctx: ExcelContext) -> None:" in source


def test_validate_procedure_module_accepts_callable_run(tmp_path):
    module = ModuleType("valid_procedure")

    def run(ctx):
        return ctx

    module.run = run
    module.NAME = "Custom Name"
    module.DESCRIPTION = "Custom description"

    info = validate_procedure_module(module, path=tmp_path / "valid_procedure.py")

    assert info.name == "Custom Name"
    assert info.description == "Custom description"
    assert info.run is run


def test_validate_procedure_module_uses_metadata_defaults(tmp_path):
    module = ModuleType("format_report")


    def run(ctx):
        return ctx


    module.run = run

    info = validate_procedure_module(module, path=tmp_path / "format_report.py")

    assert info.name == "Format Report"
    assert info.description == ""


def test_validate_procedure_module_rejects_missing_run(tmp_path):
    module = ModuleType("missing_run")


    with pytest.raises(InvalidProcedureError):
        validate_procedure_module(module, path=tmp_path / "missing_run.py")


def test_validate_procedure_module_rejects_non_callable_run(tmp_path):
    module = ModuleType("bad_run")
    module.run = "not callable"

    with pytest.raises(InvalidProcedureError):
        validate_procedure_module(module, path=tmp_path / "bad_run.py")


def test_load_procedure_imports_valid_python_file(tmp_path):
    path = tmp_path / "hello.py"
    path.write_text(
        "NAME = 'Hello'\n"
        "DESCRIPTION = 'Says hello'\n"
        "\n"
        "def run(ctx):\n"
        "    return 'ok'\n",
        encoding="utf-8",
    )

    info = load_procedure(path)

    assert info.name == "Hello"
    assert info.description == "Says hello"
    assert info.run(None) == "ok"


def test_load_procedure_rejects_missing_file(tmp_path):
    with pytest.raises(ProcedureLoadError):
        load_procedure(tmp_path / "missing.py")
