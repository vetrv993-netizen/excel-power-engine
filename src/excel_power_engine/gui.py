from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path
from typing import Any

from .audit import log_event, read_recent
from .engine import ExcelEngine
from .open_excel import open_in_excel
from .safe_edit import EditOperation
from .sheet_viewer import read_region, workbook_sheets
from .formula_intelligence import cell_info, trace_formula
from .bulk import parse_matrix, read_data_file, preview_matrix, execute_matrix
from .search_engine import search_workbook
from .target_engine import parse_target_specs
from .format_engine import ExcelOperations
from .operations import preview_operations
from .workbook_context import analyze_workbook, WorkbookContext
from .orchestrator import preview_workspace, run_workspace, rollback_workspace
from .theme import ThemeManager
from .session import WorkbookSession
from . import __version__


def _qapp():
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("Excel Power Engine")
    app.setOrganizationName("Excel Power Engine")
    try:
        from PySide6.QtCore import Qt
        app.setLayoutDirection(Qt.RightToLeft)
    except Exception:
        pass
    return app


try:
    from PySide6.QtWidgets import QMainWindow as _QMainWindow
except ImportError:
    class _QMainWindow:
        pass


class MainWindow(_QMainWindow):
    def __init__(self):
        from PySide6.QtWidgets import (
            QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
            QFileDialog, QTabWidget, QComboBox, QSpinBox, QGroupBox, QFormLayout,
            QTableWidget, QTableWidgetItem, QAbstractItemView, QTextEdit, QMessageBox,
            QCheckBox
        )
        super().__init__()
        self.QWidget = QWidget
        self.QVBoxLayout = QVBoxLayout
        self.QHBoxLayout = QHBoxLayout
        self.QLabel = QLabel
        self.QLineEdit = QLineEdit
        self.QPushButton = QPushButton
        self.QFileDialog = QFileDialog
        self.QTabWidget = QTabWidget
        self.QComboBox = QComboBox
        self.QSpinBox = QSpinBox
        self.QGroupBox = QGroupBox
        self.QFormLayout = QFormLayout
        self.QTableWidget = QTableWidget
        self.QTableWidgetItem = QTableWidgetItem
        self.QAbstractItemView = QAbstractItemView
        self.QTextEdit = QTextEdit
        self.QMessageBox = QMessageBox
        self.QCheckBox = QCheckBox

        from PySide6.QtWidgets import QApplication
        self.theme_manager = ThemeManager(QApplication.instance())
        self.engine = ExcelEngine()
        self.current_path: Path | None = None
        self.context: WorkbookContext | None = None
        self.session: WorkbookSession | None = None
        self.last_edit_sheet: str | None = None
        self.last_edit_cell: str | None = None
        display_version = ".".join(__version__.split(".")[:2])
        self.setWindowTitle(f"Excel Power Engine v{display_version} — مساحة العمل الذكية")
        self.resize(1280, 820)
        self.tabs = self.QTabWidget()
        self.setCentralWidget(self.tabs)
        self._build_file_tab()
        self._build_context_tab()
        self._build_viewer_tab()
        self._build_edit_tab()
        self._build_intelligence_tab()
        self._build_operations_tab()
        self._build_workspace_tab()
        self._build_settings_tab()
        self._build_audit_tab()
        self._refresh_audit()

    def _build_settings_tab(self):
        root = self.QWidget()
        layout = self.QVBoxLayout(root)
        box = self.QGroupBox("الإعدادات")
        form = self.QFormLayout(box)
        self.theme_combo = self.QComboBox()
        self.theme_combo.addItem("نهاري", "light")
        self.theme_combo.addItem("ليلي", "dark")
        self.theme_combo.addItem("تلقائي", "auto")
        self.theme_combo.setCurrentIndex(
            max(0, self.theme_combo.findData(self.theme_manager.preference()))
        )
        self.theme_combo.currentIndexChanged.connect(self._theme_changed)
        form.addRow("مظهر الواجهة:", self.theme_combo)
        self.theme_status = self.QLabel()
        form.addRow("الوضع الحالي:", self.theme_status)
        layout.addWidget(box)
        layout.addStretch()
        self.tabs.addTab(root, "الإعدادات")
        self._apply_theme(self.theme_manager.preference())

    def _theme_changed(self):
        self._apply_theme(self.theme_combo.currentData())

    def _apply_theme(self, preference):
        active = self.theme_manager.apply(preference)
        if hasattr(self, "theme_status"):
            labels = {"light": "نهاري", "dark": "ليلي"}
            self.theme_status.setText(labels[active])

    def _build_file_tab(self):
        root = self.QWidget()
        layout = self.QVBoxLayout(root)
        box = self.QGroupBox("الملف الحالي")
        form = self.QFormLayout(box)
        row = self.QHBoxLayout()
        self.path_edit = self.QLineEdit()
        self.browse_btn = self.QPushButton("اختيار ملف Excel")
        self.browse_btn.clicked.connect(self._choose_file)
        row.addWidget(self.path_edit, 1); row.addWidget(self.browse_btn)
        form.addRow("المسار:", row)
        self.sheet_combo = self.QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._sheet_changed)
        form.addRow("الورقة:", self.sheet_combo)
        layout.addWidget(box)
        buttons = self.QHBoxLayout()
        for text, slot in [
            ("فحص سريع", self._inspect),
            ("فحص عميق", self._inspect_deep),
            ("عرض Excel", self._open_excel),
        ]:
            b=self.QPushButton(text); b.clicked.connect(slot); buttons.addWidget(b)
        layout.addLayout(buttons)
        self.profile_text = self.QTextEdit(); self.profile_text.setReadOnly(True)
        layout.addWidget(self.profile_text, 1)
        self.tabs.addTab(root, "الملف والفحص")

    def _build_context_tab(self):
        root = self.QWidget(); layout = self.QVBoxLayout(root)
        head = self.QHBoxLayout()
        self.context_file_label = self.QLabel("لا يوجد ملف محدد")
        self.context_status = self.QLabel("لم يتم تحليل المصنف بعد")
        refresh = self.QPushButton("تحليل ذكي للمصنف")
        refresh.clicked.connect(self._refresh_context)
        apply_btn = self.QPushButton("تطبيق المقترحات على الأدوات")
        apply_btn.clicked.connect(self._apply_context_suggestions)
        head.addWidget(self.context_file_label, 2); head.addWidget(self.context_status, 2); head.addWidget(refresh); head.addWidget(apply_btn)
        layout.addLayout(head)

        summary = self.QGroupBox("ملخص المصنف")
        sf = self.QFormLayout(summary)
        self.context_summary = self.QTextEdit(); self.context_summary.setReadOnly(True); self.context_summary.setMaximumHeight(170)
        sf.addRow(self.context_summary)
        layout.addWidget(summary)

        cols_box = self.QGroupBox("الحقول والأعمدة المكتشفة — اقتراحات جاهزة")
        cf = self.QFormLayout(cols_box)
        self.context_columns = self.QTableWidget(); self.context_columns.setColumnCount(4)
        self.context_columns.setHorizontalHeaderLabels(["العمود", "اسم الحقل", "النوع", "عينات"])
        self.context_columns.setEditTriggers(self.QAbstractItemView.NoEditTriggers)
        self.context_columns.setSelectionMode(self.QAbstractItemView.SingleSelection)
        cf.addRow(self.context_columns)
        layout.addWidget(cols_box, 2)

        sug_box = self.QGroupBox("المقترحات المستخدمة تلقائيًا")
        qf = self.QFormLayout(sug_box)
        self.context_suggestions = self.QTextEdit(); self.context_suggestions.setReadOnly(True)
        qf.addRow(self.context_suggestions)
        layout.addWidget(sug_box, 1)
        self.tabs.addTab(root, "سياق المصنف الذكي")

    def _refresh_context(self):
        if not self.current_path:
            return self._warn("اختر ملفًا أولًا")
        try:
            selected = self.sheet_combo.currentText() or None
            cell = self.viewer_cell.text().strip().upper() or None
            self.session = WorkbookSession.open(
                self.current_path,
                selected_sheet=selected,
                selected_cell=cell,
            )
            self.context = self.session.workbook_context
            self._render_context()
            self._apply_context_suggestions()
            log_event("workbook-context", path=str(self.current_path), details={"sheet_count": self.context.sheet_count, "selected_sheet": self.context.selected_sheet})
        except Exception as exc:
            self._error(exc)

    def _render_context(self):
        if not self.context:
            return
        selected = self.context.sheet(self.context.selected_sheet)
        self.context_file_label.setText(self.current_path.name if self.current_path else "")
        if not selected:
            self.context_status.setText("لا توجد ورقة مكتشفة")
            return
        self.context_status.setText(f"تحليل مكتمل ✓  |  {self.context.sheet_count} ورقة  |  {selected.name}")
        summary = {
            "الملف": str(self.context.path),
            "XLSM": self.context.xlsm,
            "VBA": self.context.has_vba,
            "الورقة الحالية": selected.name,
            "النطاق المستخدم": selected.used_range,
            "صف العناوين المكتشف": selected.header_row,
            "بداية البيانات": selected.data_start_row,
            "نهاية البيانات": selected.data_end_row,
            "عدد حقول الصيغ": selected.formula_cells,
            "عدد الخلايا ذات القيم": selected.value_cells,
            "الخلايا المدمجة": selected.merged_count,
            "محمية": selected.protected,
            "الألوان المقترحة": selected.dominant_fills,
            "ألوان الخط المقترحة": selected.dominant_font_colors,
        }
        self.context_summary.setPlainText(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
        self.context_columns.setRowCount(len(selected.columns))
        for r, col in enumerate(selected.columns):
            vals=[col.get("column",""), col.get("title",""), col.get("kind",""), " | ".join(col.get("samples",[])[:3])]
            for c,v in enumerate(vals):
                self.context_columns.setItem(r,c,self.QTableWidgetItem(str(v)))
        self.context_columns.resizeColumnsToContents()
        self.context_suggestions.setPlainText(json.dumps(self.context.suggestions, ensure_ascii=False, indent=2, default=str))

    def _apply_context_suggestions(self):
        if not self.context:
            return
        s = self.context.suggestions
        sheet = s.get("sheet") or self.sheet_combo.currentText()
        cell = s.get("cell") or "A1"
        target = s.get("target") or "A1"
        ranges = s.get("ranges") or target
        self.viewer_cell.setText(str(cell))
        self.info_cell.setText(str(cell))
        self.edit_cell.setText(str(cell))
        self.bulk_start.setText(str(cell))
        self.bulk_ranges.setText(str(ranges))
        self.op_ranges.setText(str(ranges))
        if s.get("fill"):
            self.op_fill.setText(str(s["fill"]))
        if s.get("font_color"):
            self.op_font.setText(str(s["font_color"]))
        for combo in (self.sheet_combo, self.info_sheet, self.bulk_sheet, self.edit_sheet, self.op_sheet, self.op_search_sheet):
            try:
                combo.setCurrentText(sheet)
            except Exception:
                pass
        selected = self.context.sheet(sheet)
        if selected and selected.columns:
            mapping = ", ".join(f"{c['title']}→{c['column']}" for c in selected.columns if c['title'] != c['column'] and c['title'])
            if hasattr(self, "bulk_data") and mapping:
                self.bulk_data.setPlaceholderText("مقترحات الحقول: " + mapping + "\nألصق البيانات أو استوردها هنا.")

    def _build_viewer_tab(self):
        root = self.QWidget(); layout = self.QVBoxLayout(root)
        top = self.QHBoxLayout()
        self.viewer_cell = self.QLineEdit("A1"); self.viewer_cell.setPlaceholderText("الخلية المركزية مثل V266")
        self.row_radius = self.QSpinBox(); self.row_radius.setRange(2, 50); self.row_radius.setValue(10)
        self.col_radius = self.QSpinBox(); self.col_radius.setRange(2, 30); self.col_radius.setValue(8)
        load = self.QPushButton("تحميل المعاينة"); load.clicked.connect(self._load_view)
        top.addWidget(self.QLabel("الخلية:")); top.addWidget(self.viewer_cell)
        top.addWidget(self.QLabel("صفوف:")); top.addWidget(self.row_radius)
        top.addWidget(self.QLabel("أعمدة:")); top.addWidget(self.col_radius)
        top.addWidget(load)
        layout.addLayout(top)
        self.viewer = self.QTableWidget(); self.viewer.setEditTriggers(self.QAbstractItemView.NoEditTriggers)
        self.viewer.setSelectionMode(self.QAbstractItemView.SingleSelection)
        layout.addWidget(self.viewer, 1)
        self.viewer_formula = self.QLineEdit(); self.viewer_formula.setReadOnly(True)
        layout.addWidget(self.QLabel("الصيغة/القيمة للخلية المحددة:")); layout.addWidget(self.viewer_formula)
        self.tabs.addTab(root, "معاينة الورقة")

    def _build_edit_tab(self):
        root = self.QWidget(); layout = self.QVBoxLayout(root)
        box = self.QGroupBox("Safe Edit / Smart Edit")
        form = self.QFormLayout(box)
        self.edit_sheet = self.QComboBox()
        self.edit_cell = self.QLineEdit("A1")
        self.edit_kind = self.QComboBox(); self.edit_kind.addItems(["قيمة", "صيغة"])
        self.edit_content = self.QLineEdit()
        self.output_edit = self.QLineEdit()
        self.backup_check = self.QCheckBox("إنشاء نسخة احتياطية"); self.backup_check.setChecked(True)
        form.addRow("الورقة:", self.edit_sheet)
        form.addRow("الخلية:", self.edit_cell)
        form.addRow("نوع التعديل:", self.edit_kind)
        form.addRow("المحتوى:", self.edit_content)
        form.addRow("ملف الناتج:", self.output_edit)
        form.addRow("الحماية:", self.backup_check)
        layout.addWidget(box)
        row = self.QHBoxLayout()
        plan=self.QPushButton("تحديد المحرك"); plan.clicked.connect(self._plan_edit)
        execute=self.QPushButton("تنفيذ Smart Edit"); execute.clicked.connect(self._execute_edit)
        visual=self.QPushButton("مشاهدة التعديل في Excel"); visual.clicked.connect(self._open_edited)
        row.addWidget(plan); row.addWidget(execute); row.addWidget(visual)
        layout.addLayout(row)
        self.edit_log = self.QTextEdit(); self.edit_log.setReadOnly(True); layout.addWidget(self.edit_log,1)
        self.tabs.addTab(root, "Safe Edit")

    def _build_intelligence_tab(self):
        root = self.QWidget(); layout = self.QVBoxLayout(root)

        info_box = self.QGroupBox("ذكاء الخلية والصيغة")
        info_form = self.QFormLayout(info_box)
        info_row = self.QHBoxLayout()
        self.info_sheet = self.QComboBox()
        self.info_cell = self.QLineEdit("A1")
        inspect_btn = self.QPushButton("تحليل الخلية")
        inspect_btn.clicked.connect(self._inspect_cell_intelligence)
        info_row.addWidget(self.info_sheet, 2); info_row.addWidget(self.info_cell, 1); info_row.addWidget(inspect_btn)
        info_form.addRow("الورقة / الخلية:", info_row)
        self.info_output = self.QTextEdit(); self.info_output.setReadOnly(True)
        info_form.addRow(self.info_output)
        layout.addWidget(info_box, 1)

        bulk_box = self.QGroupBox("الإدخال الشامل متعدد الخلايا / النطاق")
        bulk_form = self.QFormLayout(bulk_box)
        bulk_top = self.QHBoxLayout()
        self.bulk_sheet = self.QComboBox()
        self.bulk_start = self.QLineEdit("F11")
        self.bulk_ranges = self.QLineEdit()
        self.bulk_ranges.setPlaceholderText("مثال: F11:I70;F76:I135;F141:I200;F206:I265;F271:I330")
        self.bulk_file = self.QLineEdit()
        browse_bulk = self.QPushButton("استيراد CSV / Excel")
        browse_bulk.clicked.connect(self._choose_bulk_file)
        bulk_top.addWidget(self.bulk_sheet, 2)
        bulk_top.addWidget(self.QLabel("بداية النطاق:")); bulk_top.addWidget(self.bulk_start, 1)
        bulk_top.addWidget(self.QLabel("أو نطاقات متعددة:")); bulk_top.addWidget(self.bulk_ranges, 3)
        bulk_top.addWidget(browse_bulk)
        bulk_form.addRow("الوجهة:", bulk_top)
        self.bulk_data = self.QTextEdit()
        self.bulk_data.setPlaceholderText("ألصق البيانات هنا، مفصولة بعلامات TAB أو CSV. كل سطر = صف وكل عمود = خلية.")
        bulk_form.addRow("البيانات:", self.bulk_data)
        bulk_opts = self.QHBoxLayout()
        self.bulk_formulas = self.QCheckBox("اعتبار النص الذي يبدأ بـ = صيغة")
        self.bulk_formulas.setChecked(True)
        self.bulk_clear = self.QCheckBox("المربعات الفارغة تمسح البيانات الحالية")
        self.bulk_clear.setChecked(False)
        self.bulk_preview_btn = self.QPushButton("معاينة التغييرات")
        self.bulk_preview_btn.clicked.connect(self._bulk_preview)
        self.bulk_execute_btn = self.QPushButton("تنفيذ الإدخال الشامل")
        self.bulk_execute_btn.clicked.connect(self._bulk_execute)
        self.bulk_visual_btn = self.QPushButton("مشاهدة أول تعديل في Excel")
        self.bulk_visual_btn.clicked.connect(self._open_bulk_visual)
        for w in (self.bulk_formulas, self.bulk_clear, self.bulk_preview_btn, self.bulk_execute_btn, self.bulk_visual_btn): bulk_opts.addWidget(w)
        bulk_opts.addStretch()
        bulk_form.addRow("التنفيذ:", bulk_opts)
        self.bulk_output = self.QLineEdit()
        bulk_form.addRow("ملف الناتج:", self.bulk_output)
        self.bulk_preview_text = self.QTextEdit(); self.bulk_preview_text.setReadOnly(True)
        bulk_form.addRow("المعاينة:", self.bulk_preview_text)
        layout.addWidget(bulk_box, 2)
        self.tabs.addTab(root, "الذكاء والإدخال الشامل")

    def _sync_v06_sheets(self):
        if not self.current_path:
            return
        sheets = list(workbook_sheets(self.current_path).keys())
        for combo in (self.info_sheet, self.bulk_sheet):
            combo.clear(); combo.addItems(sheets)
        if sheets:
            self.info_sheet.setCurrentText(self.sheet_combo.currentText() or sheets[0])
            self.bulk_sheet.setCurrentText(self.sheet_combo.currentText() or sheets[0])

    def _inspect_cell_intelligence(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        sheet = self.info_sheet.currentText() or self.sheet_combo.currentText()
        cell = self.info_cell.text().strip().upper()
        if not sheet or not cell: return self._warn("حدد الورقة والخلية")
        try:
            result = {"cell": cell_info(self.current_path, sheet, cell).as_dict()}
            result["trace"] = trace_formula(self.current_path, sheet, cell, max_depth=5)
            self.info_output.setPlainText(json.dumps(result, ensure_ascii=False, indent=2, default=str))
            log_event("cell-intelligence", path=str(self.current_path), details={"sheet": sheet, "cell": cell})
        except Exception as exc: self._error(exc)

    def _choose_bulk_file(self):
        path, _ = self.QFileDialog.getOpenFileName(self, "اختر ملف بيانات", str(Path.home()), "Data (*.csv *.tsv *.txt *.xlsx *.xlsm)")
        if not path: return
        try:
            self.bulk_file.setText(path)
            data = read_data_file(path)
            lines = ["\t".join("" if v is None else str(v) for v in row) for row in data]
            self.bulk_data.setPlainText("\n".join(lines))
        except Exception as exc: self._error(exc)

    def _bulk_matrix(self):
        if self.bulk_data.toPlainText().strip():
            return parse_matrix(self.bulk_data.toPlainText())
        path = self.bulk_file.text().strip()
        if path:
            return read_data_file(path)
        return []

    def _bulk_preview(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        sheet = self.bulk_sheet.currentText() or self.sheet_combo.currentText()
        start = self.bulk_start.text().strip().upper()
        try:
            matrix = self._bulk_matrix()
            if not matrix: return self._warn("أدخل البيانات أو اختر ملف بيانات")
            if self.bulk_ranges.text().strip():
                from .bulk import preview_multi_ranges
                changes = preview_multi_ranges(self.current_path, sheet, self.bulk_ranges.text().strip(), matrix, interpret_formulas=self.bulk_formulas.isChecked(), clear_empty=self.bulk_clear.isChecked())
            else:
                changes = preview_matrix(self.current_path, sheet, start, matrix, interpret_formulas=self.bulk_formulas.isChecked(), clear_empty=self.bulk_clear.isChecked())
            changed = [c for c in changes if c.before != c.after]
            preview = {
                "الورقة": sheet,
                "بداية النطاق": start,
                "عدد الصفوف": len(matrix),
                "عدد الأعمدة": max((len(r) for r in matrix), default=0),
                "عدد الخلايا المتغيرة": len(changed),
                "التغييرات": [c.as_dict() for c in changed[:500]],
            }
            self.bulk_preview_text.setPlainText(json.dumps(preview, ensure_ascii=False, indent=2, default=str))
            log_event("bulk-preview", path=str(self.current_path), details={"sheet": sheet, "start_cell": start, "changed_count": len(changed)})
        except Exception as exc: self._error(exc)

    def _bulk_execute(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        sheet = self.bulk_sheet.currentText() or self.sheet_combo.currentText()
        start = self.bulk_start.text().strip().upper()
        try:
            matrix = self._bulk_matrix()
            if not matrix: return self._warn("أدخل البيانات أو اختر ملف بيانات")
            if not self.bulk_output.text().strip():
                self.bulk_output.setText(str(self.current_path.with_name(self.current_path.stem + "_BULK_EDIT" + self.current_path.suffix)))
            output = self.bulk_output.text().strip()
            # Require an explicit preview before commit when no preview text exists.
            if not self.bulk_preview_text.toPlainText().strip():
                self._bulk_preview()
                if not self.bulk_preview_text.toPlainText().strip(): return
            if self.bulk_ranges.text().strip():
                from .bulk import execute_multi_ranges
                result = execute_multi_ranges(self.current_path, sheet, self.bulk_ranges.text().strip(), matrix, output, interpret_formulas=self.bulk_formulas.isChecked(), clear_empty=self.bulk_clear.isChecked(), create_backup=True)
                first_cell = self.bulk_ranges.text().split(";", 1)[0].split(",", 1)[0].split(":", 1)[0].strip().upper()
            else:
                result = execute_matrix(self.current_path, sheet, start, matrix, output, interpret_formulas=self.bulk_formulas.isChecked(), clear_empty=self.bulk_clear.isChecked(), create_backup=True)
                first_cell = start
            self.bulk_preview_text.setPlainText(json.dumps(result, ensure_ascii=False, indent=2, default=str))
            self.last_edit_sheet = sheet
            self.last_edit_cell = first_cell
            self.last_output = Path(output)
            log_event("bulk-edit", path=str(self.current_path), details=result)
            self._open_bulk_visual()
        except Exception as exc: self._error(exc)

    def _open_bulk_visual(self):
        if not getattr(self, "last_output", None): return self._warn("نفّذ الإدخال الشامل أولًا")
        try:
            result = open_in_excel(self.last_output, self.last_edit_sheet, self.last_edit_cell)
            log_event("open-excel", path=str(self.last_output), details=result)
        except Exception as exc: self._error(exc)

    def _build_operations_tab(self):
        root=self.QWidget(); layout=self.QVBoxLayout(root)
        search_box=self.QGroupBox("البحث الذكي")
        sf=self.QFormLayout(search_box)
        row=self.QHBoxLayout()
        self.op_search_term=self.QLineEdit(); self.op_search_mode=self.QComboBox(); self.op_search_mode.addItems(["contains","exact","starts","ends","regex"])
        self.op_search_sheet=self.QComboBox(); self.op_search_btn=self.QPushButton("بحث")
        self.op_search_btn.clicked.connect(self._op_search)
        row.addWidget(self.QLabel("الكلمة:")); row.addWidget(self.op_search_term,2); row.addWidget(self.QLabel("النمط:")); row.addWidget(self.op_search_mode); row.addWidget(self.QLabel("الورقة:")); row.addWidget(self.op_search_sheet,2); row.addWidget(self.op_search_btn)
        sf.addRow(row)
        self.op_search_output=self.QTextEdit(); self.op_search_output.setReadOnly(True); sf.addRow(self.op_search_output)
        layout.addWidget(search_box,2)

        style_box=self.QGroupBox("تنسيق / حدود / دمج / طباعة / PDF")
        of=self.QFormLayout(style_box)
        row1=self.QHBoxLayout(); self.op_sheet=self.QComboBox(); self.op_ranges=self.QLineEdit(); self.op_ranges.setPlaceholderText("A1:C5;F11:I70")
        row1.addWidget(self.QLabel("الورقة:")); row1.addWidget(self.op_sheet,2); row1.addWidget(self.QLabel("النطاقات:")); row1.addWidget(self.op_ranges,5)
        of.addRow(row1)
        row2=self.QHBoxLayout(); self.op_fill=self.QLineEdit(); self.op_fill.setPlaceholderText("#FFF2CC"); self.op_font=self.QLineEdit(); self.op_font.setPlaceholderText("#000000"); self.op_bold=self.QCheckBox("غامق"); self.op_wrap=self.QCheckBox("التفاف"); self.op_merge=self.QCheckBox("دمج"); self.op_unmerge=self.QCheckBox("إلغاء الدمج")
        row2.addWidget(self.QLabel("تعبئة:")); row2.addWidget(self.op_fill); row2.addWidget(self.QLabel("لون الخط:")); row2.addWidget(self.op_font); row2.addWidget(self.op_bold); row2.addWidget(self.op_wrap); row2.addWidget(self.op_merge); row2.addWidget(self.op_unmerge)
        of.addRow(row2)
        row3=self.QHBoxLayout(); self.op_format_btn=self.QPushButton("تطبيق التنسيق"); self.op_borders_btn=self.QPushButton("تطبيق الحدود"); self.op_merge_btn=self.QPushButton("تنفيذ الدمج"); self.op_pdf_btn=self.QPushButton("تصدير PDF")
        self.op_format_btn.clicked.connect(self._op_format); self.op_borders_btn.clicked.connect(self._op_borders); self.op_merge_btn.clicked.connect(self._op_merge); self.op_pdf_btn.clicked.connect(self._op_pdf)
        for b in (self.op_format_btn,self.op_borders_btn,self.op_merge_btn,self.op_pdf_btn): row3.addWidget(b)
        of.addRow(row3)
        self.op_output=self.QLineEdit(); of.addRow("ملف الناتج:", self.op_output)
        self.op_output_log=self.QTextEdit(); self.op_output_log.setReadOnly(True); of.addRow("النتيجة:", self.op_output_log)
        layout.addWidget(style_box,2)
        self.tabs.addTab(root,"العمليات الشاملة")

    def _op_sheet_defaults(self):
        if not self.current_path: return
        sheets=list(workbook_sheets(self.current_path).keys())
        self.op_search_sheet.clear(); self.op_search_sheet.addItems(sheets)
        self.op_sheet.clear(); self.op_sheet.addItems(sheets)
        current=self.sheet_combo.currentText()
        if current:
            self.op_search_sheet.setCurrentText(current); self.op_sheet.setCurrentText(current)

    def _op_search(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        term=self.op_search_term.text().strip()
        if not term: return self._warn("اكتب كلمة أو جزءًا من كلمة")
        sheet=self.op_search_sheet.currentText() or None
        try:
            hits=search_workbook(self.current_path,term,sheets=[sheet] if sheet else None,mode=self.op_search_mode.currentText(),max_results=2000)
            payload={"term":term,"count":len(hits),"hits":[h.as_dict() for h in hits]}
            self.op_search_output.setPlainText(json.dumps(payload,ensure_ascii=False,indent=2,default=str))
            log_event("search",path=str(self.current_path),details={"term":term,"count":len(hits)})
        except Exception as exc: self._error(exc)

    def _op_selected(self):
        sheet=self.op_sheet.currentText() or self.sheet_combo.currentText()
        ranges=[x.strip() for x in self.op_ranges.text().split(";") if x.strip()]
        if not sheet: raise ValueError("اختر ورقة")
        if not ranges: raise ValueError("أدخل نطاقًا واحدًا أو عدة نطاقات")
        for r in ranges: parse_target_specs(r)
        return sheet,ranges

    def _op_output_path(self, suffix="OPERATIONS"):
        if not self.current_path: return None
        p=self.op_output.text().strip()
        return p or str(self.current_path.with_name(self.current_path.stem+"_"+suffix+self.current_path.suffix))

    def _op_format(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            sheet,ranges=self._op_selected()
            out=self._op_output_path("FORMAT")
            result=ExcelOperations().format(self.current_path,sheet,ranges,output=out,fill=self.op_fill.text().strip() or None,font_color=self.op_font.text().strip() or None,bold=self.op_bold.isChecked() if self.op_bold.isChecked() else None,wrap=self.op_wrap.isChecked() if self.op_wrap.isChecked() else None,backup=True)
            self.op_output_log.setPlainText(json.dumps(result.as_dict(),ensure_ascii=False,indent=2,default=str)); self.last_output=Path(out); self.last_edit_sheet=sheet; self.last_edit_cell=ranges[0].split(":",1)[0]
            log_event("format",path=str(self.current_path),details=result.as_dict()); self._open_bulk_visual()
        except Exception as exc: self._error(exc)

    def _op_borders(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            sheet,ranges=self._op_selected(); out=self._op_output_path("BORDERS")
            result=ExcelOperations().borders(self.current_path,sheet,ranges,output=out,backup=True)
            self.op_output_log.setPlainText(json.dumps(result.as_dict(),ensure_ascii=False,indent=2,default=str)); self.last_output=Path(out); self.last_edit_sheet=sheet; self.last_edit_cell=ranges[0].split(":",1)[0]; log_event("borders",path=str(self.current_path),details=result.as_dict()); self._open_bulk_visual()
        except Exception as exc: self._error(exc)

    def _op_merge(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            sheet,ranges=self._op_selected(); out=self._op_output_path("MERGE")
            unmerge=self.op_unmerge.isChecked() and not self.op_merge.isChecked()
            result=ExcelOperations().merge(self.current_path,sheet,ranges,output=out,unmerge=unmerge,backup=True)
            self.op_output_log.setPlainText(json.dumps(result.as_dict(),ensure_ascii=False,indent=2,default=str)); self.last_output=Path(out); self.last_edit_sheet=sheet; self.last_edit_cell=ranges[0].split(":",1)[0]; log_event("merge",path=str(self.current_path),details=result.as_dict()); self._open_bulk_visual()
        except Exception as exc: self._error(exc)

    def _op_pdf(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            sheet,ranges=self._op_selected()
            pdf=str(Path(self._op_output_path("PDF")).with_suffix(".pdf"))
            result=ExcelOperations().export_pdf(self.current_path,output_pdf=pdf,sheet=sheet,ranges=ranges)
            self.op_output_log.setPlainText(json.dumps(result,ensure_ascii=False,indent=2,default=str)); log_event("pdf",path=str(self.current_path),details=result)
        except Exception as exc: self._error(exc)


    def _build_workspace_tab(self):
        root=self.QWidget(); layout=self.QVBoxLayout(root)
        head=self.QHBoxLayout()
        self.ws_source_label=self.QLabel("لا يوجد ملف محدد")
        self.ws_output=self.QLineEdit()
        self.ws_output.setPlaceholderText("ملف الناتج؛ اتركه فارغًا لاستخدام اسم تلقائي")
        head.addWidget(self.ws_source_label,2); head.addWidget(self.ws_output,3)
        layout.addLayout(head)
        hint=self.QLabel("مساحة عمل موحدة: بحث + إدخال شامل + تنسيق + حدود + دمج + طباعة/PDF + تراجع")
        layout.addWidget(hint)
        self.ws_json=self.QTextEdit()
        self.ws_json.setPlainText(json.dumps({"operations":[{"kind":"search","sheet":"","payload":{"term":"التهاب","mode":"contains"}},{"kind":"format","sheet":"الشهر الخامس","ranges":["H11:I70"],"payload":{"fill":"#FFF2CC","bold":True}}]}, ensure_ascii=False, indent=2))
        layout.addWidget(self.ws_json,2)
        row=self.QHBoxLayout()
        self.ws_preview_btn=self.QPushButton("معاينة العملية"); self.ws_preview_btn.clicked.connect(self._workspace_preview)
        self.ws_run_btn=self.QPushButton("تنفيذ العملية الموحدة"); self.ws_run_btn.clicked.connect(self._workspace_run)
        self.ws_rollback_btn=self.QPushButton("تراجع عن آخر عملية"); self.ws_rollback_btn.clicked.connect(self._workspace_rollback)
        self.ws_apply_context_btn=self.QPushButton("استخدام سياق المصنف"); self.ws_apply_context_btn.clicked.connect(self._workspace_from_context)
        for b in (self.ws_preview_btn,self.ws_run_btn,self.ws_rollback_btn,self.ws_apply_context_btn): row.addWidget(b)
        layout.addLayout(row)
        self.ws_result=self.QTextEdit(); self.ws_result.setReadOnly(True); layout.addWidget(self.ws_result,2)
        self.tabs.addTab(root,"مساحة العمل الذكية")

    def _workspace_from_context(self):
        if not self.context:
            return self._warn("حلّل المصنف أولًا")
        s=self.context.suggestions
        sheet=s.get("sheet") or self.sheet_combo.currentText()
        target=s.get("ranges") or s.get("target") or "A1"
        payload={"fill":s.get("fill"),"font_color":s.get("font_color")}
        payload={k:v for k,v in payload.items() if v}
        op={"kind":"format","sheet":sheet,"ranges":[str(target)],"payload":payload}
        self.ws_json.setPlainText(json.dumps({"operations":[op]},ensure_ascii=False,indent=2))
        self._workspace_preview()

    def _workspace_payload(self):
        if not self.current_path:
            raise ValueError("اختر ملفًا أولًا")
        try:
            return json.loads(self.ws_json.toPlainText())
        except json.JSONDecodeError as exc:
            raise ValueError(f"JSON غير صحيح: {exc}") from exc

    def _workspace_preview(self):
        try:
            result=preview_workspace(self.session or self.current_path,self._workspace_payload())
            self.ws_source_label.setText(self.current_path.name)
            self.ws_result.setPlainText(json.dumps(result,ensure_ascii=False,indent=2,default=str))
            log_event("workspace-preview",path=str(self.current_path),details={"operation_count":result.get("operation_count")})
        except Exception as exc: self._error(exc)

    def _workspace_run(self):
        try:
            out=self.ws_output.text().strip()
            if not out:
                out=str(self.current_path.with_name(self.current_path.stem+"_WORKSPACE"+self.current_path.suffix))
            result=run_workspace(self.session or self.current_path,self._workspace_payload(),out,backup=True,visible=False,open_after=True)
            self.ws_result.setPlainText(json.dumps(result,ensure_ascii=False,indent=2,default=str))
            self.last_output=Path(out)
            self.last_edit_sheet=(result.get("applied_operations") or [{}])[-1].get("result",{}).get("sheet") if result.get("applied_operations") else self.sheet_combo.currentText()
            if result.get("visual_verification",{}).get("opened"):
                log_event("workspace-visual-check",path=str(out),details=result.get("visual_verification",{}))
        except Exception as exc: self._error(exc)

    def _workspace_rollback(self):
        try:
            out=self.ws_output.text().strip()
            if not out:
                out=str(self.current_path.with_name(self.current_path.stem+"_WORKSPACE"+self.current_path.suffix))
            result=rollback_workspace(out)
            self.ws_result.setPlainText(json.dumps(result,ensure_ascii=False,indent=2,default=str))
        except Exception as exc: self._error(exc)

    def _build_audit_tab(self):
        root=self.QWidget(); layout=self.QVBoxLayout(root)
        row=self.QHBoxLayout(); refresh=self.QPushButton("تحديث السجل"); refresh.clicked.connect(self._refresh_audit); row.addWidget(refresh); row.addStretch()
        layout.addLayout(row)
        self.audit_text=self.QTextEdit(); self.audit_text.setReadOnly(True); layout.addWidget(self.audit_text,1)
        self.tabs.addTab(root, "سجل العمليات")

    def _choose_file(self):
        path, _ = self.QFileDialog.getOpenFileName(self, "اختر ملف Excel", str(Path.home()), "Excel (*.xlsx *.xlsm)")
        if path:
            self._set_path(Path(path))

    def _set_path(self, path: Path):
        self.current_path=path.resolve()
        self.session=WorkbookSession.open(self.current_path)
        self.context=self.session.workbook_context
        self.path_edit.setText(str(self.current_path))
        if hasattr(self, "ws_source_label"): self.ws_source_label.setText(self.current_path.name)
        sheets=list(workbook_sheets(self.current_path).keys())
        self.sheet_combo.clear(); self.sheet_combo.addItems(sheets)
        self.edit_sheet.clear(); self.edit_sheet.addItems(sheets)
        if sheets:
            self._sync_v06_sheets()
            self._op_sheet_defaults()
            self._load_view()
            self._refresh_context()
            self._inspect()

    def _sheet_changed(self, sheet):
        if sheet:
            self._sync_v06_sheets()
            self._op_sheet_defaults()
            self._load_view()
            if self.context:
                self._refresh_context()

    def _inspect(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            p=self.engine.profile(self.current_path, deep=False)
            self.profile_text.setPlainText(json.dumps({
                "الملف": str(p.path), "النوع": p.suffix, "VBA": p.has_vba,
                "OOXML": p.is_ooxml, "الأوراق": list(p.sheet_parts), "الميزات": sorted(p.features)
            },ensure_ascii=False,indent=2))
            log_event("inspect", path=str(self.current_path), details={"deep":False})
        except Exception as exc: self._error(exc)

    def _inspect_deep(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            p=self.engine.profile(self.current_path, deep=True)
            self.profile_text.setPlainText(json.dumps({
                "الملف": str(p.path), "النوع": p.suffix, "VBA": p.has_vba,
                "OOXML": p.is_ooxml, "الأوراق": p.sheet_parts,
                "features": sorted(p.features), "integrity": p.integrity,
                "sheet_profiles": p.sheet_profiles,
            },ensure_ascii=False,indent=2))
            log_event("inspect-deep", path=str(self.current_path), details={"deep":True})
        except Exception as exc: self._error(exc)

    def _load_view(self):
        if not self.current_path or not self.sheet_combo.currentText(): return
        try:
            cols, rows, target = read_region(self.current_path, self.sheet_combo.currentText(), center_cell=self.viewer_cell.text().strip() or None,
                                             row_radius=self.row_radius.value(), col_radius=self.col_radius.value())
            self.viewer.clear(); self.viewer.setColumnCount(len(cols)+1); self.viewer.setRowCount(len(rows));
            self.viewer.setHorizontalHeaderLabels(["صف"]+cols)
            target_pos=None
            for ri,(row_num,row_data) in enumerate(rows):
                self.viewer.setItem(ri,0,self.QTableWidgetItem(str(row_num)))
                for ci,item in enumerate(row_data,1):
                    text="" if item is None else (f"={item.formula}" if item.formula else item.value)
                    self.viewer.setItem(ri,ci,self.QTableWidgetItem(text))
                    if item and item.ref.upper()==(target or "").upper(): target_pos=(ri,ci)
            self.viewer.resizeColumnsToContents()
            if target_pos: self.viewer.setCurrentCell(*target_pos)
            self.viewer_formula.setText("" if not target_pos else self.viewer.item(*target_pos).text())
        except Exception as exc: self._error(exc)

    def _plan_edit(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        try:
            plan=self.engine.plan(self.current_path,"cell-edit")
            self.edit_log.setPlainText(json.dumps(plan.as_dict(),ensure_ascii=False,indent=2))
            log_event("plan-edit", path=str(self.current_path), details=plan.as_dict())
        except Exception as exc: self._error(exc)

    def _execute_edit(self):
        if not self.current_path: return self._warn("اختر ملفًا أولًا")
        sheet=self.edit_sheet.currentText(); cell=self.edit_cell.text().strip().upper(); content=self.edit_content.text()
        if not sheet or not cell: return self._warn("اختر الورقة والخلية")
        output=self.output_edit.text().strip() or str(self.current_path.with_name(self.current_path.stem+"_SMART_EDIT"+self.current_path.suffix))
        try:
            op=EditOperation(cell, formula=content if self.edit_kind.currentText()=="صيغة" else None,
                             value=None if self.edit_kind.currentText()=="صيغة" else content)
            result=self.engine.smart_edit(self.current_path,sheet,[op],output,recalculate=False,create_backup=self.backup_check.isChecked())
            self.edit_log.setPlainText(json.dumps(result,ensure_ascii=False,indent=2))
            self.last_edit_sheet=sheet; self.last_edit_cell=cell; self.last_output=Path(output)
            log_event("smart-edit", path=str(self.current_path), details=result)
            self._open_edited()
            self.tabs.setCurrentIndex(1)
            self.current_path=self.last_output; self.path_edit.setText(str(self.current_path))
            sheets=list(workbook_sheets(self.current_path).keys()); self.sheet_combo.clear(); self.sheet_combo.addItems(sheets); self.edit_sheet.clear(); self.edit_sheet.addItems(sheets)
            self.sheet_combo.setCurrentText(sheet); self.edit_sheet.setCurrentText(sheet); self.viewer_cell.setText(cell); self._load_view()
        except Exception as exc: self._error(exc)

    def _open_excel(self):
        """Open the current workbook and navigate to the selected cell visually."""
        if not self.current_path:
            return self._warn("اختر ملفًا أولًا")
        sheet = self.sheet_combo.currentText() or None
        cell = (self.viewer_cell.text().strip().upper() if self.viewer_cell.text().strip() else "A1")
        if not sheet:
            return self._warn("اختر ورقة أولًا")
        try:
            result = open_in_excel(self.current_path, sheet, cell)
            log_event("open-excel", path=str(self.current_path), details=result)
        except Exception as exc:
            self._error(exc)

    def _open_edited(self):
        if not getattr(self,"last_output",None):
            if self.current_path and self.edit_cell.text().strip():
                path=self.current_path; sheet=self.edit_sheet.currentText(); cell=self.edit_cell.text().strip().upper()
            else: return self._warn("نفّذ تعديلًا أولًا")
        else:
            path=self.last_output; sheet=self.last_edit_sheet; cell=self.last_edit_cell
        try:
            result=open_in_excel(path,sheet,cell)
            log_event("open-excel", path=str(path), details=result)
        except Exception as exc: self._error(exc)

    def _refresh_audit(self):
        self.audit_text.setPlainText("\n\n".join(json.dumps(x,ensure_ascii=False,indent=2) for x in read_recent(200)))

    def _warn(self,msg): self.QMessageBox.warning(self,"Excel Power Engine",msg)
    def _error(self,exc):
        log_event("error",status="error",details={"message":str(exc),"traceback":traceback.format_exc()})
        self.QMessageBox.critical(self,"خطأ",str(exc))


def main() -> int:
    app=_qapp(); w=MainWindow(); w.show(); return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
