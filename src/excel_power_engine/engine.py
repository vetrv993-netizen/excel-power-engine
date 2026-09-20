from __future__ import annotations
from pathlib import Path
from .inspect import inspect_workbook
from .models import WorkbookProfile
from .smart_engine import SmartEditEngine, build_plan
from .formula_intelligence import cell_info, trace_formula, formula_issues

class ExcelEngine:
    def __init__(self):
        self.smart = SmartEditEngine()

    def profile(self, path: str|Path, deep: bool=False) -> WorkbookProfile:
        return inspect_workbook(path, deep=deep)

    def choose(self, path: str|Path, operation: str, *, recalculate: bool=False, prefer_native: bool=False) -> str:
        return build_plan(path, operation, recalculate=recalculate, prefer_native=prefer_native).engine

    def plan(self, path: str|Path, operation: str='cell-edit', *, recalculate: bool=False, prefer_native: bool=False):
        return self.smart.plan(path, operation, recalculate=recalculate, prefer_native=prefer_native)

    def smart_edit(self, source, sheet, edits, output=None, *, recalculate=False, prefer_native=False, create_backup=True):
        return self.smart.edit_cells(source, sheet, edits, output, recalculate=recalculate, prefer_native=prefer_native, create_backup=create_backup)

    def cell_info(self, path, sheet, cell):
        return cell_info(path, sheet, cell)

    def formula_trace(self, path, sheet, cell, *, max_depth=6):
        return trace_formula(path, sheet, cell, max_depth=max_depth)

    def formula_issues(self, path):
        return formula_issues(path)

    def search(self, path, term, **kwargs):
        from .search_engine import search_workbook
        return search_workbook(path, term, **kwargs)

    def resolve_targets(self, path, sheet, text, **kwargs):
        from .target_engine import resolve_target_specs
        return resolve_target_specs(path, sheet, text, **kwargs)
