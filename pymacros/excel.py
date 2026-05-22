from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, TypeVar
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

import pythoncom
import win32com.client as win32

from pymacros.context import ExcelContext
from pymacros.errors import (
    ExcelSessionError,
    ProcedureExecutionError,
    WorkbookCloseError,
    WorkbookNotFoundError,
    WorkbookOpenError,
    WorkbookSaveError,
)


T = TypeVar("T")
Procedure = Callable[[ExcelContext], T]


@dataclass
class ExcelConfig:
    visible: bool = True
    display_alerts: bool = False
    screen_updating: bool = False
    enable_events: bool = False


class ExcelSession:
    def __init__(
        self,
        config: ExcelConfig | None = None,
        *,
        close_on_exit: bool = True,
        manage_com: bool = True,
    ):
        self.config = config or ExcelConfig()
        self.close_on_exit = close_on_exit
        self.manage_com = manage_com
        self.app = None
        self._previous_state: dict[str, Any] = {}

    def __enter__(self):
        if self.manage_com:
            pythoncom.CoInitialize()

        # DispatchEx crea una nueva instancia de Excel,
        # en vez de conectarse a una ya abierta.
        self.app = win32.DispatchEx("Excel.Application")

        self._capture_previous_state()
        self._apply_config()

        return self

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            if self.close_on_exit:
                self.quit()
            else:
                self.detach()
        finally:
            if self.manage_com:
                pythoncom.CoUninitialize()

    def _capture_previous_state(self):
        self._previous_state = {
            "Visible": self.app.Visible,
            "DisplayAlerts": self.app.DisplayAlerts,
            "ScreenUpdating": self.app.ScreenUpdating,
            "EnableEvents": self.app.EnableEvents,
        }

    def _apply_config(self):
        self.app.Visible = self.config.visible
        self.app.DisplayAlerts = self.config.display_alerts
        self.app.ScreenUpdating = self.config.screen_updating
        self.app.EnableEvents = self.config.enable_events

    def restore_state(self, *, include_visible: bool = True):
        for key, value in self._previous_state.items():
            if key == "Visible" and not include_visible:
                continue

            try:
                setattr(self.app, key, value)
            except Exception:
                pass

    def open_book(
        self,
        path: str | Path,
        *,
        read_only: bool = False,
        update_links: int = 0,
    ):
        if self.app is None:
            raise ExcelSessionError("Excel session is not started.")

        target = _workbook_target(path)

        try:
            return self.app.Workbooks.Open(
                str(target),
                UpdateLinks=update_links,
                ReadOnly=read_only,
            )
        except Exception as exc:
            raise WorkbookOpenError(f"Could not open workbook {target}: {exc}") from exc

    def new_book(self):
        if self.app is None:
            raise ExcelSessionError("Excel session is not started.")

        return self.app.Workbooks.Add()

    def run(
        self,
        workbook: Any,
        procedure: Procedure[T],
    ) -> T:
        ctx = ExcelContext(
            app=self.app,
            workbook=workbook,
        )

        workbook.Activate()

        return procedure(ctx)

    def run_on_book(
        self,
        path: str | Path,
        procedure: Procedure[T],
        *,
        save: bool = False,
        close: bool = True,
        read_only: bool = False,
        update_links: int = 0,
    ) -> T:
        target = _workbook_target(path)
        workbook = self.open_book(
            path,
            read_only=read_only,
            update_links=update_links,
        )
        primary_error: BaseException | None = None

        try:
            try:
                result = self.run(workbook, procedure)
            except Exception as exc:
                primary_error = ProcedureExecutionError(
                    f"Procedure failed while running on workbook {target}: {exc}"
                )
                raise primary_error from exc

            if save:
                try:
                    workbook.Save()
                except Exception as exc:
                    primary_error = WorkbookSaveError(
                        f"Could not save workbook {target}: {exc}"
                    )
                    raise primary_error from exc

            return result

        finally:
            if close:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception as exc:
                    close_error = WorkbookCloseError(
                        f"Could not close workbook {target}: {exc}"
                    )

                    if primary_error is None:
                        raise close_error from exc

                    primary_error.add_note(f"Additionally failed to close workbook: {exc}")

    def quit(self):
        if self.app is None:
            return

        try:
            self.restore_state()
        except Exception:
            pass

        try:
            self.app.Quit()
        finally:
            self.app = None

    def detach(self):
        if self.app is None:
            return

        try:
            self.restore_state(include_visible=False)
        except Exception:
            pass

        self.app = None

def run_workbook(
    path: str | Path,
    procedure: Procedure[T],
    *,
    config: ExcelConfig | None = None,
    save: bool = False,
    close_workbook: bool = True,
    close_excel: bool = True,
    manage_com: bool = True,
    read_only: bool = False,
    update_links: int = 0,
) -> T:
    with ExcelSession(config, close_on_exit=close_excel, manage_com=manage_com) as excel:
        return excel.run_on_book(
            path,
            procedure,
            save=save,
            close=close_workbook,
            read_only=read_only,
            update_links=update_links,
        )


def _workbook_target(path: str | Path) -> str | Path:
    if _is_web_url(path):
        return _excel_url_target(str(path))

    resolved_path = Path(path).resolve()

    if not resolved_path.exists():
        raise WorkbookNotFoundError(f"Workbook not found: {resolved_path}")

    return resolved_path


def _is_web_url(path: str | Path) -> bool:
    if not isinstance(path, str):
        return False

    parsed = urlparse(path)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _excel_url_target(url: str) -> str:
    parsed = urlparse(url)
    query_params = [
        (key, value)
        for key, value in parse_qsl(parsed.query, keep_blank_values=True)
        if key.lower() != "web"
    ]
    query = urlencode(query_params, doseq=True)
    return urlunparse(parsed._replace(query=query, fragment=""))
