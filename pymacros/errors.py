class PymacrosError(Exception):
    """Base exception for pymacros errors."""


class ExcelSessionError(PymacrosError):
    """Raised when an Excel session cannot complete an operation."""


class WorkbookNotFoundError(ExcelSessionError):
    """Raised when the requested workbook path does not exist."""


class WorkbookOpenError(ExcelSessionError):
    """Raised when Excel cannot open a workbook."""


class ProcedureExecutionError(ExcelSessionError):
    """Raised when a procedure fails while running on a workbook."""


class WorkbookSaveError(ExcelSessionError):
    """Raised when Excel cannot save a workbook."""


class WorkbookCloseError(ExcelSessionError):
    """Raised when Excel cannot close a workbook."""


class ProcedureLoadError(PymacrosError):
    """Raised when a procedure file cannot be loaded."""


class InvalidProcedureError(PymacrosError):
    """Raised when a procedure module does not expose a valid run(ctx)."""


class ProcedureExistsError(PymacrosError):
    """Raised when creating or renaming would overwrite a procedure."""


class ProcedureNotFoundError(PymacrosError):
    """Raised when a procedure file does not exist."""
