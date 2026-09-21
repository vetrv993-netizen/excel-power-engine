from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

try:
    _version_file = Path(__file__).resolve().parents[2] / "VERSION.txt"
    __version__ = _version_file.read_text(encoding="utf-8").strip()
except (OSError, UnicodeError):
    try:
        __version__ = version("excel-power-engine")
    except PackageNotFoundError:
        __version__ = "unknown"

from .engine import ExcelEngine
from .inspect import inspect_workbook
from .safe_edit import EditOperation, SafeEditor
from .smart_engine import SmartEditEngine, SmartPlan, build_plan
from .formula_intelligence import CellSnapshot, cell_info, trace_formula, formula_issues, extract_references
from .bulk import BulkChange, parse_matrix, preview_file as bulk_preview_file, execute_file as bulk_execute_file

__all__=[
    'ExcelEngine','inspect_workbook','EditOperation','SafeEditor',
    'SmartEditEngine','SmartPlan','build_plan',
    'CellSnapshot','cell_info','trace_formula','formula_issues','extract_references',
    'BulkChange','parse_matrix','bulk_preview_file','bulk_execute_file',
    'SearchHit','search_workbook','WorkbookContext','SheetContext','analyze_workbook','WorkbookSession','TargetSet','SearchResultSet','VerificationResult','verify_workbook_change','TargetSpec','parse_target_specs','resolve_target_specs','ExcelOperations','run_workflow','validate_workflow','CommandIntent','CommandPlan','parse_command','build_command_plan','SemanticWorkbook','SemanticCandidate','SemanticResolution','normalize_text','inspect_cell','smart_suggestions','classify_operation','SAFE','WARNING','HIGH_RISK','OperationHistory',
]

from .search_engine import SearchHit, search_workbook
from .target_engine import TargetSpec, parse_target_specs, resolve_target_specs
from .format_engine import ExcelOperations
from .workflow import run_workflow, validate_workflow

from .workbook_context import WorkbookContext, SheetContext, analyze_workbook
from .session import WorkbookSession, TargetSet, SearchResultSet
from .verification import VerificationResult, verify_workbook_change

from .command_intelligence import CommandIntent, CommandPlan, parse_command, build_command_plan
from .semantic_engine import SemanticWorkbook, SemanticCandidate, SemanticResolution, normalize_text
from .inspector import inspect_cell
from .suggestions import smart_suggestions
from .risk import classify_operation, SAFE, WARNING, HIGH_RISK
from .history import OperationHistory
