from pymacros import ExcelContext


class FakeRange:
    def __init__(self):
        self.Value = None


class FakeWorksheet:
    def __init__(self):
        self.ranges = {}

    def Range(self, address):
        return self.ranges.setdefault(address, FakeRange())


class FakeWorkbook:
    def __init__(self):
        self.sheets = {"Hoja1": FakeWorksheet()}
        self.saved = False
        self.saved_as = None
        self.closed_with = None

    def Worksheets(self, name):
        return self.sheets[name]

    def Save(self):
        self.saved = True

    def SaveAs(self, path):
        self.saved_as = path

    def Close(self, *, SaveChanges):
        self.closed_with = SaveChanges


class FakeApp:
    def __init__(self, active_sheet):
        self.ActiveSheet = active_sheet


def test_context_reads_and_writes_active_sheet():
    workbook = FakeWorkbook()
    ctx = ExcelContext(app=FakeApp(workbook.sheets["Hoja1"]), workbook=workbook)

    ctx.write("A1", "value")

    assert ctx.read("A1") == "value"


def test_context_reads_and_writes_named_sheet():
    workbook = FakeWorkbook()
    ctx = ExcelContext(app=FakeApp(FakeWorksheet()), workbook=workbook)

    ctx.write("B2", 42, sheet="Hoja1")


    assert ctx.read("B2", sheet="Hoja1") == 42


def test_context_saves_and_closes_workbook(tmp_path):
    workbook = FakeWorkbook()
    ctx = ExcelContext(app=FakeApp(workbook.sheets["Hoja1"]), workbook=workbook)
    save_as_path = tmp_path / "out.xlsx"

    ctx.save()
    ctx.save_as(save_as_path)
    ctx.close(save_changes=True)

    assert workbook.saved is True
    assert workbook.saved_as == str(save_as_path.resolve())
    assert workbook.closed_with is True
