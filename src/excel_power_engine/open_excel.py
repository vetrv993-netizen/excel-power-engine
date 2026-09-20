from __future__ import annotations

import os
import subprocess
from pathlib import Path


def _vbs_string(value: str) -> str:
    """Quote a Python string as a VBScript string literal."""
    return '"' + value.replace('"', '""') + '"'


def _bring_window_to_front(hwnd: int) -> None:
    """Best-effort foreground activation for the Excel top-level window."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def _activate_target_in_excel(xl, wb, sheet: str, cell: str) -> None:
    """Force Excel to display the requested worksheet/cell visibly."""
    ws = wb.Worksheets(sheet)

    try:
        wb.Activate()
    except Exception:
        pass

    try:
        win = wb.Windows(1)
        win.Activate()
        # xlMaximized = -4137
        win.WindowState = -4137
    except Exception:
        win = None

    ws.Activate()
    rng = ws.Range(cell)

    try:
        rng.Activate()
    except Exception:
        pass
    try:
        rng.Select()
    except Exception:
        pass
    try:
        xl.Goto(rng, True)
    except Exception:
        pass

    try:
        wb.Activate()
    except Exception:
        pass
    if win is not None:
        try:
            win.Activate()
        except Exception:
            pass

    # Scroll the worksheet so the target cell is inside the visible grid.
    try:
        aw = xl.ActiveWindow
        aw.ScrollRow = max(1, int(rng.Row) - 5)
        aw.ScrollColumn = max(1, int(rng.Column) - 4)
    except Exception:
        pass

    # Keep the exact target as the active selection after scrolling.
    try:
        ws.Activate()
        rng.Activate()
        rng.Select()
    except Exception:
        pass

    try:
        _bring_window_to_front(int(xl.Hwnd))
    except Exception:
        pass

    try:
        import win32gui
        win32gui.ShowWindow(int(xl.Hwnd), 9)
        win32gui.SetForegroundWindow(int(xl.Hwnd))
    except Exception:
        pass


def _open_with_pywin32(path: Path, sheet: str, cell: str) -> dict[str, object]:
    import win32com.client as win32

    try:
        xl = win32.GetActiveObject("Excel.Application")
    except Exception:
        xl = win32.DispatchEx("Excel.Application")

    xl.Visible = True
    try:
        xl.UserControl = True
    except Exception:
        pass

    target = str(path).casefold()
    wb = None
    for candidate in xl.Workbooks:
        try:
            if str(candidate.FullName).casefold() == target:
                wb = candidate
                break
        except Exception:
            continue

    if wb is None:
        wb = xl.Workbooks.Open(
            str(path),
            ReadOnly=False,
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
        )

    _activate_target_in_excel(xl, wb, str(sheet), str(cell))

    return {
        "opened": True,
        "path": str(path),
        "sheet": sheet,
        "cell": cell,
        "visual_only": True,
        "navigation": "pywin32-activate-goto-scroll",
    }


def _open_with_vbs(path: Path, sheet: str, cell: str) -> dict[str, object]:
    """Fallback for systems where pywin32 Excel automation is unavailable."""
    import tempfile

    script = (
        "On Error Resume Next\n"
        "Set xl = GetObject(, \"Excel.Application\")\n"
        "If Err.Number <> 0 Then\n"
        "  Err.Clear\n"
        "  Set xl = CreateObject(\"Excel.Application\")\n"
        "End If\n"
        "On Error GoTo 0\n"
        "xl.Visible = True\n"
        f"Set wb = xl.Workbooks.Open({_vbs_string(str(path))})\n"
        f"Set ws = wb.Worksheets({_vbs_string(str(sheet))})\n"
        "wb.Activate\n"
        "ws.Activate\n"
        f"Set rng = ws.Range({_vbs_string(str(cell))})\n"
        "rng.Select\n"
        "xl.Goto rng, True\n"
        "On Error Resume Next\n"
        "wb.Windows(1).Activate\n"
        "xl.Visible = True\n"
        "xl.UserControl = True\n"
    )

    with tempfile.NamedTemporaryFile("w", suffix=".vbs", delete=False, encoding="utf-8") as f:
        f.write(script)
        vbs = f.name

    try:
        completed = subprocess.run(
            ["cscript.exe", "//nologo", vbs],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout or "VBScript failed").strip()
            raise RuntimeError(f"Excel visual navigation failed: {detail}")
    finally:
        try:
            os.remove(vbs)
        except OSError:
            pass

    return {
        "opened": True,
        "path": str(path),
        "sheet": sheet,
        "cell": cell,
        "visual_only": True,
        "navigation": "vbscript-activate-goto",
    }


def open_in_excel(path: str | Path, sheet: str | None = None, cell: str | None = None) -> dict[str, object]:
    """Open a workbook in desktop Excel and optionally navigate to a cell.

    Visual-only: this helper does not edit workbook contents. Navigation uses
    pywin32 first so the workbook, worksheet, cell, and Excel window are all
    explicitly activated and brought to the foreground; VBScript is a fallback.
    """
    path = Path(path).resolve()
    if os.name != "nt":
        raise RuntimeError("open-excel is supported on Windows only.")
    if not path.exists():
        raise FileNotFoundError(path)

    if not sheet or not cell:
        try:
            import win32com.client as win32
            try:
                xl = win32.GetActiveObject("Excel.Application")
            except Exception:
                xl = win32.DispatchEx("Excel.Application")
            xl.Visible = True
            wb = xl.Workbooks.Open(str(path), ReadOnly=False, UpdateLinks=0, IgnoreReadOnlyRecommended=True)
            wb.Activate()
            try:
                _bring_window_to_front(int(xl.Hwnd))
            except Exception:
                pass
            return {"opened": True, "path": str(path), "sheet": sheet, "cell": cell, "visual_only": True, "navigation": "pywin32-open"}
        except Exception:
            subprocess.Popen(["cmd", "/c", "start", "", str(path)], close_fds=True)
            return {"opened": True, "path": str(path), "sheet": sheet, "cell": cell, "visual_only": True, "navigation": "shell-open"}

    try:
        return _open_with_pywin32(path, str(sheet), str(cell))
    except Exception:
        return _open_with_vbs(path, str(sheet), str(cell))
