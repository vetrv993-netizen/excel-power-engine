from excel_power_engine import open_excel as mod


class FakeWindow:
    def __init__(self):
        self.activated = False
        self.WindowState = None
    def Activate(self):
        self.activated = True


class FakeRange:
    Row = 266
    Column = 22
    def __init__(self):
        self.activated = False
        self.selected = False
    def Activate(self):
        self.activated = True
    def Select(self):
        self.selected = True


class FakeSheet:
    def __init__(self, rng):
        self.rng = rng
        self.activated = False
    def Activate(self):
        self.activated = True
    def Range(self, cell):
        assert cell == "V266"
        return self.rng


class FakeSheets:
    def __init__(self, sheet):
        self.sheet = sheet
    def __call__(self, name):
        assert name == "الشهر الثاني"
        return self.sheet


class FakeBook:
    def __init__(self, sheet, window):
        self.Worksheets = FakeSheets(sheet)
        self._window = window
        self.activated = False
    def Windows(self, i):
        return self._window
    def Activate(self):
        self.activated = True


class FakeActiveWindow:
    def __init__(self):
        self.ScrollRow = None
        self.ScrollColumn = None


class FakeExcel:
    Hwnd = 123
    def __init__(self):
        self.ActiveWindow = FakeActiveWindow()
        self.goto_called = False
    def Goto(self, rng, scroll):
        self.goto_called = (rng, scroll)


def test_navigation_selects_and_scrolls_target(monkeypatch):
    rng = FakeRange()
    ws = FakeSheet(rng)
    window = FakeWindow()
    wb = FakeBook(ws, window)
    xl = FakeExcel()

    monkeypatch.setattr(mod, "_bring_window_to_front", lambda hwnd: None)

    mod._activate_target_in_excel(xl, wb, "الشهر الثاني", "V266")

    assert wb.activated
    assert ws.activated
    assert rng.activated
    assert rng.selected
    assert xl.goto_called == (rng, True)
    assert xl.ActiveWindow.ScrollRow == 261
    assert xl.ActiveWindow.ScrollColumn == 18
    assert window.activated
    assert window.WindowState == -4137
