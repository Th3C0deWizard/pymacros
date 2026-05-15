from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass
class ExcelContext:
    app: Any
    workbook: Any

    @property
    def active_sheet(self):
        return self.app.ActiveSheet

    def sheet(self, name: str):
        return self.workbook.Worksheets(name)

    def range(self, address: str, sheet: str | None = None):
        ws = self.sheet(sheet) if sheet else self.active_sheet
        return ws.Range(address)

    def read(self, address: str, sheet: str | None = None):
        return self.range(address, sheet).Value

    def write(self, address: str, value: Any, sheet: str | None = None):
        self.range(address, sheet).Value = value

    def save(self):
        self.workbook.Save()

    def save_as(self, path: str | Path):
        self.workbook.SaveAs(str(Path(path).resolve()))

    def close(self, save_changes: bool = False):
        self.workbook.Close(SaveChanges=save_changes)
