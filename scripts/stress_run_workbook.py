from __future__ import annotations

import time
import traceback
from datetime import datetime

from pymacros import ExcelContext, run_workbook


# Edit these values before running the script on Windows.
WORKBOOK_PATH = r"C:\Users\tec.planeacion\Dev\pymacros\data\stress_test.xlsx"
ITERATIONS = 50
DELAY_SECONDS = 1


def smoke_test(ctx: ExcelContext) -> dict[str, object]:
    """Exercise several Excel COM operations and return only plain Python values."""
    app = ctx.app
    workbook = ctx.workbook
    sheet = ctx.active_sheet
    sheet_count = workbook.Worksheets.Count
    sheet_names = [workbook.Worksheets(index).Name for index in range(1, sheet_count + 1)]
    original_sheet_name = sheet.Name
    original_a1 = sheet.Range("A1").Value
    used_range_address = sheet.UsedRange.Address
    temp_sheet = None

    app.StatusBar = f"pymacros stress test {datetime.now():%H:%M:%S}"

    try:
        temp_sheet = workbook.Worksheets.Add(After=workbook.Worksheets(sheet_count))
        temp_sheet.Name = f"_pymacros_stress_{datetime.now():%H%M%S}"

        temp_sheet.Range("A1").Value = "pymacros stress test"
        temp_sheet.Range("A2:C4").Value = (
            (1, 2, 3),
            (4, 5, 6),
            (7, 8, 9),
        )
        temp_sheet.Range("A6").Formula = "=SUM(A2:C4)"
        temp_sheet.Range("A1:C1").Font.Bold = True
        temp_sheet.Range("A1:C6").Columns.AutoFit()
        temp_sheet.Calculate()

        formula_result = temp_sheet.Range("A6").Value
        written_value = temp_sheet.Range("B3").Value
        sheet.Activate()

    finally:
        if temp_sheet is not None:
            temp_sheet.Delete()
        app.StatusBar = False

    return {
        "workbook": workbook.Name,
        "active_sheet": original_sheet_name,
        "first_sheet": sheet_names[0] if sheet_names else None,
        "sheets": sheet_count,
        "a1": original_a1,
        "used_range": used_range_address,
        "written_value": written_value,
        "formula_result": formula_result,
    }


def main() -> int:
    if WORKBOOK_PATH == r"C:\path\to\workbook.xlsx":
        print("Edit WORKBOOK_PATH in this script before running it.")
        return 2

    print("pymacros run_workbook stress test")
    print(f"Workbook: {WORKBOOK_PATH}")
    print(f"Iterations: {ITERATIONS}")
    print(f"Delay: {DELAY_SECONDS} seconds")
    print()

    completed = 0
    started_at = time.perf_counter()

    try:
        for iteration in range(1, ITERATIONS + 1):
            print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] iteration {iteration}/{ITERATIONS}: start")
            iteration_started_at = time.perf_counter()

            try:
                result = run_workbook(WORKBOOK_PATH, smoke_test)
            except Exception:
                print(f"iteration {iteration}: FAILED")
                traceback.print_exc()
                return 1

            completed += 1
            elapsed = time.perf_counter() - iteration_started_at
            print(f"iteration {iteration}: ok in {elapsed:.2f}s -> {result}")

            if iteration < ITERATIONS:
                print(f"waiting {DELAY_SECONDS} seconds...")
                time.sleep(DELAY_SECONDS)

    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        return 130

    total_elapsed = time.perf_counter() - started_at
    print()
    print(f"Completed {completed}/{ITERATIONS} iterations in {total_elapsed:.2f}s.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
