import pytest

from pymacros import (
    ProcedureExecutionError,
    WorkbookCloseError,
    WorkbookNotFoundError,
    WorkbookSaveError,
)
import pymacros.excel as excel_module


def test_run_workbook_uses_session_to_run_and_close(monkeypatch, tmp_path):
    calls = []
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")

    def procedure(ctx):
        return "result"

    class FakeSession:
        def __init__(self, config, *, close_on_exit=True):
            calls.append(("init", config, close_on_exit))

        def __enter__(self):
            calls.append(("enter",))
            return self

        def __exit__(self, exc_type, exc_value, traceback):
            calls.append(("exit", exc_type))

        def run_on_book(self, path, proc, *, save, close, read_only, update_links):
            calls.append(("run_on_book", path, proc, save, close, read_only, update_links))
            return proc(None)

    monkeypatch.setattr(excel_module, "ExcelSession", FakeSession)

    result = excel_module.run_workbook(
        workbook_path,
        procedure,
        config="config",
        save=True,
        read_only=True,
        update_links=3,
    )

    assert result == "result"
    assert calls == [
        ("init", "config", True),
        ("enter",),
        ("run_on_book", workbook_path, procedure, True, True, True, 3),
        ("exit", None),
    ]


def test_run_workbook_can_keep_workbook_and_excel_open(monkeypatch, tmp_path):
    calls = []
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")

    def procedure(ctx):
        return "result"

    class FakeSession:
        def __init__(self, config, *, close_on_exit=True):
            calls.append(("init", config, close_on_exit))

        def __enter__(self):
            calls.append(("enter",))
            return self

        def __exit__(self, exc_type, exc_value, traceback):
            calls.append(("exit", exc_type))

        def run_on_book(self, path, proc, *, save, close, read_only, update_links):
            calls.append(("run_on_book", path, proc, save, close, read_only, update_links))
            return proc(None)

    monkeypatch.setattr(excel_module, "ExcelSession", FakeSession)

    result = excel_module.run_workbook(
        workbook_path,
        procedure,
        config="config",
        save=True,
        close_workbook=False,
        close_excel=False,
        read_only=True,
        update_links=3,
    )

    assert result == "result"
    assert calls == [
        ("init", "config", False),
        ("enter",),
        ("run_on_book", workbook_path, procedure, True, False, True, 3),
        ("exit", None),
    ]


class FakeWorkbooks:
    def __init__(self, workbook):
        self.workbook = workbook
        self.open_calls = []

    def Open(self, path, *, UpdateLinks, ReadOnly):
        self.open_calls.append((path, UpdateLinks, ReadOnly))
        return self.workbook


class FakeApp:
    def __init__(self, workbook):
        self.Workbooks = FakeWorkbooks(workbook)


class FakeWorkbook:
    def __init__(self, *, save_error=None, close_error=None):
        self.save_error = save_error
        self.close_error = close_error
        self.activated = False
        self.saved = False
        self.closed = False

    def Activate(self):
        self.activated = True

    def Save(self):
        if self.save_error:
            raise self.save_error

        self.saved = True

    def Close(self, *, SaveChanges):
        if self.close_error:
            raise self.close_error

        self.closed = True


class FakeQuitApp(FakeApp):
    def __init__(self):
        super().__init__(FakeWorkbook())
        self.Visible = True
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.EnableEvents = True
        self.quit_called = False

    def Quit(self):
        self.quit_called = True


def test_excel_session_detaches_instead_of_quitting_when_close_on_exit_false(monkeypatch):
    monkeypatch.setattr(excel_module.pythoncom, "CoUninitialize", lambda: None)
    app = FakeQuitApp()
    session = excel_module.ExcelSession(close_on_exit=False)
    session.app = app

    session.__exit__(None, None, None)

    assert app.quit_called is False
    assert session.app is None


def test_excel_session_keep_open_restores_interactive_flags_without_hiding(monkeypatch):
    monkeypatch.setattr(excel_module.pythoncom, "CoUninitialize", lambda: None)
    app = FakeQuitApp()
    app.Visible = True
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    app.EnableEvents = False
    session = excel_module.ExcelSession(close_on_exit=False)
    session.app = app
    session._previous_state = {
        "Visible": False,
        "DisplayAlerts": True,
        "ScreenUpdating": True,
        "EnableEvents": True,
    }

    session.__exit__(None, None, None)

    assert app.Visible is True
    assert app.DisplayAlerts is True
    assert app.ScreenUpdating is True
    assert app.EnableEvents is True
    assert app.quit_called is False
    assert session.app is None


def test_run_on_book_raises_workbook_not_found(tmp_path):
    session = excel_module.ExcelSession()
    session.app = FakeApp(FakeWorkbook())

    with pytest.raises(WorkbookNotFoundError):
        session.run_on_book(tmp_path / "missing.xlsx", lambda ctx: None)


def test_run_on_book_wraps_procedure_error_and_closes_workbook(tmp_path):
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    workbook = FakeWorkbook()
    session = excel_module.ExcelSession()
    session.app = FakeApp(workbook)

    def procedure(ctx):
        raise ValueError("boom")

    with pytest.raises(ProcedureExecutionError) as exc_info:
        session.run_on_book(workbook_path, procedure)

    assert "boom" in str(exc_info.value)
    assert workbook.closed is True


def test_run_on_book_wraps_save_error_and_closes_workbook(tmp_path):
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    workbook = FakeWorkbook(save_error=RuntimeError("save failed"))
    session = excel_module.ExcelSession()
    session.app = FakeApp(workbook)

    with pytest.raises(WorkbookSaveError) as exc_info:
        session.run_on_book(workbook_path, lambda ctx: None, save=True)

    assert "save failed" in str(exc_info.value)
    assert workbook.closed is True


def test_run_on_book_raises_close_error_after_success(tmp_path):
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    workbook = FakeWorkbook(close_error=RuntimeError("close failed"))
    session = excel_module.ExcelSession()
    session.app = FakeApp(workbook)

    with pytest.raises(WorkbookCloseError) as exc_info:
        session.run_on_book(workbook_path, lambda ctx: "ok")

    assert "close failed" in str(exc_info.value)


def test_run_on_book_does_not_mask_procedure_error_with_close_error(tmp_path):
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    workbook = FakeWorkbook(close_error=RuntimeError("close failed"))
    session = excel_module.ExcelSession()
    session.app = FakeApp(workbook)

    def procedure(ctx):
        raise RuntimeError("procedure failed")

    with pytest.raises(ProcedureExecutionError) as exc_info:
        session.run_on_book(workbook_path, procedure)

    assert "procedure failed" in str(exc_info.value)
    assert "Additionally failed to close workbook" in exc_info.value.__notes__[0]
