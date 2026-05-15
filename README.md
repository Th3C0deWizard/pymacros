# pymacros

`pymacros` is a small Windows-only Python library and CLI for running custom Python procedures against Excel workbooks through Excel COM objects.

## Requirements

- Windows
- Microsoft Excel installed
- Python 3.11+
- Project virtualenv at `./venv/Scripts/python.exe`

## Install

```bash
make install
```

This installs the package in editable mode with development dependencies.

## CLI

Show help:

```bash
make cli ARGS="--help"
```

Run a procedure on a workbook:

```bash
make cli ARGS="run path/to/workbook.xlsx procedure_name --save"
```

Run interactively with searchable workbook and procedure selection:

```bash
make cli ARGS="run"
```

Keep Excel open after running:

```bash
make cli ARGS="run path/to/workbook.xlsx procedure_name --keep-open"
```

Procedure management:

```bash
make cli ARGS="procedures list"
make cli ARGS="procedures create my_procedure"
make cli ARGS="procedures show my_procedure"
make cli ARGS="procedures edit my_procedure"
make cli ARGS="procedures edit-folder"
make cli ARGS="procedures rename old_name new_name"
make cli ARGS="procedures delete my_procedure"
```

## Procedure Files

Procedures are Python files with a `run(ctx)` function:

```python
from pymacros import ExcelContext

NAME = "Format Report"
DESCRIPTION = "Formats the active report worksheet."


def run(ctx: ExcelContext) -> None:
    ws = ctx.active_sheet
    ws.Range("A1").Value = "Hello from pymacros"
    ctx.save()
```

By default, CLI-created procedures live in:

```text
%APPDATA%\pymacros\procedures
```

## Python API

```python
from pymacros import ExcelContext, run_workbook


def run(ctx: ExcelContext) -> None:
    ctx.write("A1", "Hello")


run_workbook("report.xlsx", run, save=True)
```

For lower-level control, use `ExcelSession`.

## Configuration

The CLI reads JSON config from:

```text
%APPDATA%\pymacros\config.json
```

Example:

```json
{
  "procedures_dir": "C:/Users/you/AppData/Roaming/pymacros/procedures",
  "editor": "code",
  "visible": true,
  "save": false,
  "close_excel": true,
  "read_only": false,
  "update_links": 0
}
```

Use a different config file with:

```bash
make cli ARGS="--config-file path/to/config.json run"
```

## Logs

Logs default to:

```text
%LOCALAPPDATA%\pymacros\logs\pymacros.log
```

Useful flags:

```bash
make cli ARGS="--debug run"
make cli ARGS="--log-file path/to/pymacros.log run"
```

## Development

Syntax check without starting Excel:

```bash
make check
```

Run unit tests:

```bash
make test
```

Remove generated artifacts:

```bash
make clean
```

Unit tests use fakes/mocks and do not launch Excel.
