# Project Plan

## Purpose

`pymacros` is intended to be a lightweight wrapper around `pywin32` for Excel interop through Excel COM objects. It should make it easy to create and manage Excel processes, open or create workbooks, and execute custom Python procedures against a workbook.

The project should evolve into two independent packages:

- `pymacros`: the core library for Excel COM interop and procedure execution.
- `pymacros_cli`: the CLI app that uses `pymacros` as its backend.

The dependency direction must be one-way: `pymacros_cli` can import `pymacros`, but `pymacros` must never import or depend on `pymacros_cli`.

## Package Boundaries

### Core Library: `pymacros`

The library owns all Excel and procedure execution logic:

- Excel COM lifecycle management.
- Excel application configuration.
- Workbook opening, creation, saving, and closing.
- `ExcelContext`, the object passed to custom procedures.
- Procedure contracts and validation.
- Loading procedure modules from Python files.
- Running procedures on workbooks.
- Core errors that callers can catch and render however they want.
- Low-level procedure registry utilities that are useful outside the CLI.

The library must remain usable directly from Python scripts without installing or invoking the CLI.

### CLI App: `pymacros_cli`

The CLI owns user interaction and application configuration:

- Command parsing.
- Interactive workbook selection.
- Interactive procedure selection.
- Procedure CRUD command UX.
- Pretty terminal output, tables, confirmations, and progress messages.
- CLI config files and default user paths.
- Mapping library errors to friendly user-facing messages.

The CLI should not contain Excel COM lifecycle logic. It should delegate execution to `pymacros`.

## Suggested Structure

```text
pymacros/
  __init__.py
  context.py
  excel.py
  procedures.py
  registry.py
  errors.py

pymacros_cli/
  __init__.py
  __main__.py
  app.py
  commands.py
  prompts.py
  config.py
  formatting.py

procedures/
  example_procedure.py

tests/
  test_context.py
  test_procedures.py
  test_registry.py
```

## Core Library API

The public API should focus on making custom procedures easy to write and execute.

Example direct library usage:

```python
from pymacros import run_workbook


def run(ctx):
    ctx.write("A1", "Hello from pymacros")


run_workbook("report.xlsx", run, save=True)
```

Procedure files should use one consistent callable name:

```python
from pymacros import ExcelContext


NAME = "Format Report"
DESCRIPTION = "Formats the active report worksheet."


def run(ctx: ExcelContext) -> None:
    ws = ctx.active_sheet
    ws.Range("A1").Font.Bold = True
```

Recommended public concepts:

- `ExcelConfig`: Excel application flags such as visibility, alerts, screen updating, and events.
- `ExcelContext`: procedure-facing workbook context.
- `ExcelSession`: owns Excel COM lifecycle.
- `run_workbook(...)`: convenience helper for opening a workbook, running a procedure, optionally saving, and closing.
- `load_procedure(...)`: load and validate a procedure file.
- `ProcedureInfo`: metadata used by both library consumers and the CLI.
- `ProcedureRegistry`: list, create, rename, and delete procedure files in a directory.

`ExcelContext` should keep direct access to raw COM objects while also providing convenience helpers:

- `ctx.app`
- `ctx.workbook`
- `ctx.active_sheet`
- `ctx.sheet("Hoja1")`
- `ctx.range("A1", sheet="Hoja1")`
- `ctx.read("A1", sheet="Hoja1")`
- `ctx.write("A1", value, sheet="Hoja1")`
- `ctx.save()`
- `ctx.close(save_changes=False)`

## Procedure Contract

A custom procedure is a Python module that exposes a callable `run(ctx)` function.

Optional metadata:

```python
NAME = "Human Friendly Name"
DESCRIPTION = "Short description for lists and help output."
```

The library should validate procedure modules and raise clear library exceptions when:

- The file does not exist.
- The file cannot be imported.
- `run` is missing.
- `run` is not callable.

Procedure exceptions should bubble up so both scripts and the CLI can decide how to handle them.

## Procedure Management

The core library should expose low-level, reusable operations:

- Normalize procedure names into safe Python filenames.
- Generate boilerplate procedure content.
- Validate procedure modules.
- Read procedure metadata.
- List procedure files from a directory.
- Create, rename, and delete procedure files.

The CLI should expose those operations as user-facing CRUD commands:

```bash
pymacros procedures list
pymacros procedures create my_procedure
pymacros procedures show my_procedure
pymacros procedures edit my_procedure
pymacros procedures rename old_name new_name
pymacros procedures delete my_procedure
```

Initial default procedure location can be project-local:

```text
./procedures/
```

A user-level procedure directory can be added later through CLI config if needed.

## CLI UX

The CLI should be easy to use in both direct and interactive modes.

Direct execution:

```bash
pymacros run workbook.xlsx procedure_name --save
pymacros run workbook.xlsx procedure_name --no-save
```

Interactive execution:

```bash
pymacros run
```

If the workbook path is omitted, the CLI should scan the current directory for Excel workbooks and prompt the user when there are multiple options.

Supported workbook extensions should include:

- `.xlsx`
- `.xlsm`
- `.xls`

If the procedure name is omitted, the CLI should list available procedures and prompt the user to choose one.

The default save behavior should be safe. Prefer `--no-save` by default and require `--save` to persist workbook changes.

The CLI should clearly report:

- Workbook path.
- Procedure name.
- Whether changes were saved.
- Success or failure.

Recommended CLI dependencies:

- `typer` for commands and help output.
- `rich` for tables, status messages, and readable errors.
- Optional `InquirerPy` for interactive selection prompts.

Until packaging is added, run the CLI as:

```bash
./venv/Scripts/python.exe -m pymacros_cli
```

Later, add a `pyproject.toml` console script so users can run:

```bash
pymacros
```

## Boilerplate Generation

`pymacros procedures create name` should generate a Python file with the initial structure for a custom procedure.

Example generated file:

```python
from pymacros import ExcelContext


NAME = "My Procedure"
DESCRIPTION = "Describe what this procedure does."


def run(ctx: ExcelContext) -> None:
    ws = ctx.active_sheet
    # Write your Excel automation here.
```

The CLI should prevent accidental overwrite unless the user passes `--force` or confirms the operation interactively.

## Error Handling

The library should define specific exceptions for core failures, for example:

- `PymacrosError`
- `WorkbookNotFoundError`
- `ProcedureLoadError`
- `InvalidProcedureError`
- `ProcedureExistsError`
- `ProcedureNotFoundError`

The CLI should catch those exceptions and render concise, friendly messages. Unexpected exceptions can show richer debug output when a `--debug` option is passed.

## Testing Phase

A dedicated testing phase is required for the core library before relying on the CLI.

Unit tests should focus on `pymacros` and avoid launching Excel by default. Use fakes or mocks for COM objects where needed.

Core unit test coverage should include:

- `ExcelContext.sheet`, `range`, `read`, `write`, `save`, and `close`.
- Procedure validation succeeds when `run(ctx)` exists.
- Procedure validation fails when `run` is missing or not callable.
- Procedure metadata defaults work when `NAME` or `DESCRIPTION` are missing.
- Procedure filename normalization is safe and predictable.
- Boilerplate generation creates valid Python.
- Registry list, create, rename, and delete behavior works in a temporary directory.
- `run_workbook(...)` calls the expected session methods using mocks or fakes.

Excel automation tests that launch real Excel should be optional integration tests, not part of the default unit test suite.

Recommended test dependency:

```text
pytest
```

Expected verification commands after the refactor:

```bash
./venv/Scripts/python.exe -m py_compile pymacros/*.py pymacros_cli/*.py
./venv/Scripts/python.exe -m pytest tests
```

Optional Excel integration tests can be separated and explicitly marked:

```bash
./venv/Scripts/python.exe -m pytest tests/integration -m excel
```

## Implementation Order

1. Create the `pymacros` core package for Excel COM lifecycle logic.
2. Add the refined `ExcelContext`, session, and `run_workbook(...)` public API.
3. Add procedure loading, metadata, validation, boilerplate generation, and registry utilities.
4. Add core unit tests for context helpers, procedure loading, and registry behavior.
5. Refactor `main.py` to use the new public library API as a sample.
6. Create the independent `pymacros_cli` package.
7. Add CLI commands for workbook execution and procedure CRUD.
8. Add CLI-specific tests after the core library tests are passing.
9. Add packaging metadata and a console script if desired.
10. Update `AGENTS.md` with the new package boundaries, commands, and test workflow.

## Key Design Rule

If a feature is about Excel sessions, workbooks, procedures, or execution, it belongs in `pymacros`.

If a feature is about prompts, command names, terminal display, config files, or user interaction, it belongs in `pymacros_cli`.
