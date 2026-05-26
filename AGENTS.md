# AGENTS.md

## Runtime
- This repo automates desktop Excel through COM (`pywin32`, `pythoncom`, `win32com.client`); use Windows Python with Microsoft Excel installed, not WSL/Linux Python.
- Use the project venv interpreter: `./venv/Scripts/python.exe`.
- Install the editable package and dev dependencies with `make install`.

## Entry Points
- `pymacros/excel.py` owns the COM lifecycle. `ExcelSession` calls `pythoncom.CoInitialize()`, creates a fresh Excel instance with `DispatchEx("Excel.Application")`, restores app flags, quits Excel, and calls `CoUninitialize()` on exit.
- `pymacros/procedures.py` and `pymacros/registry.py` are core library modules for loading, validating, generating, and managing Python procedure files; keep CLI concerns out of them.
- `pymacros_cli/` is the independent CLI package; it may import `pymacros`, but `pymacros` must not import CLI modules.
- Run the CLI with `make cli ARGS="..."`, `./venv/Scripts/python.exe -m pymacros_cli`, or the installed `./venv/Scripts/pymacros.exe`; key commands are `run` and `procedures list/create/show/edit/edit-folder/rename/delete`.
- CLI config is Windows-only JSON at `%APPDATA%\pymacros\config.json` by default; pass global `--config-file <path>` to override. It can set `procedures_dir`, `visible`, `save`, `close_excel`, `read_only`, `update_links`, `log_file`, and `editor`.
- CLI logs default to `%LOCALAPPDATA%\pymacros\logs\pymacros.log`; use global `--debug` for stderr tracebacks and `--log-file <path>` to override the log file.
- `procedures/` contains user-created procedure files. Procedure files should define `run(ctx: ExcelContext)` and optional `NAME` / `DESCRIPTION` metadata.

## Gotchas
- `ExcelSession.open_book(...)` accepts local workbook paths and HTTP/HTTPS workbook URLs such as SharePoint/OneDrive links; only local paths are checked with `Path.exists()`. SharePoint `?web=1` browser params are stripped before calling Excel.
- `ExcelSession.run_on_book(..., save=True)` and `run_workbook(..., save=True)` save by default before closing; pass `save=False` when changes must be discarded.

## Verification
- Prefer `make check` for a syntax-only check that does not start Excel.
- Prefer `make test` for unit tests; these tests use fakes/mocks and must not launch Excel.
