# Excel Power Engine v0.10

محرك Excel/XLSM متعدد المسارات مع ذكاء الخلايا، الإدخال الشامل، البحث، التنسيق، الحدود، الدمج، الطباعة، PDF وواجهة عربية.

## أبرز v0.10
- WorkbookSession مركزي يثبت بصمة الملف ويشارك السياق بين Workspace وGUI.
- Search Result Sets تتحول إلى Target Sets قابلة لإعادة الاستخدام مع provenance وتجميع حسب الورقة.
- Verification Gate حقيقي قبل اعتماد كل خطوة تعديل: وجود الناتج، القيم المستهدفة، أخطاء الصيغ، VBA/x14، والأجزاء المتغيرة.
- `keyed-bulk` متاح داخل Workspace مع سياسة رفض المفاتيح المفقودة أو المكررة افتراضيًا.
- Search: exact / contains / starts / ends / regex.
- Target Engine: cell / range / multi-range / row / column.
- Native formatting and borders on existing XLSX/XLSM.
- Merge / unmerge.
- Print setup and Excel-native PDF export.
- Multi-operation workflow with preview/validation.
- Backup + VBA/x14 preservation path for Native Excel operations.
- Arabic GUI tab: العمليات الشاملة.

## تشغيل Windows

```powershell
.\scripts\bootstrap_windows.ps1 -Full
.\scripts\run_gui.ps1
```

## CLI أمثلة

```powershell
# بحث
excel-power search file.xlsm "التهاب" --mode contains

# استهداف عدة نطاقات
excel-power resolve-targets file.xlsm --sheet "الشهر الخامس" --targets "F11:I70;F76:I135"

# تنسيق
excel-power format file.xlsm --sheet "الشهر الخامس" --ranges "H11:H70;H76:H135" --fill "#FFF2CC" --bold

# دمج
excel-power merge file.xlsx --sheet Sheet1 --ranges "A1:C1"

# PDF
excel-power export-pdf file.xlsm --sheet "الشهر الخامس" --ranges "A1:AI332" --output-pdf report.pdf

# Workflow
excel-power workflow-plan file.xlsm --workflow examples/workflow_demo.json --output result.xlsm
excel-power workflow-run file.xlsm --workflow examples/workflow_demo.json --output result.xlsm
```

## v0.8 — Workbook Intelligence Context

عند إدراج ملف Excel، يحلّل المحرك المصنف مرة واحدة ويبني سياقًا مركزيًا مشتركًا بين المعاينة، ذكاء الخلية، الإدخال الشامل، والعمليات الشاملة. تشمل المقترحات الورقة، الخلية، النطاقات، الحقول، أنواع البيانات، الألوان، ونطاقات العمل.

التحليل قراءة فقط ولا يحفظ المصنف. راجع `docs/V08_WORKBOOK_CONTEXT_AR.md`.

## v0.10 — Verified Smart Workspace

- Preview → Backup → Execute → Verify → Visual Check هو مسار Workspace.
- `search` يسجل Result Set؛ وتستطيع العملية التالية استخدام `target_set` مع `selected_hits`.
- كل خطوة تعديل تكتب في staging، وتُرفض قبل النشر إن فشل Verification Gate.
- ينشئ Workspace manifest بجانب الناتج (`.transaction.json`) ويستخدمه rollback لتحديد backup.
- التحقق البصري له سياسة `optional` أو `required`؛ تتطلب السياسة الثانية Excel/Windows وهدفًا بصريًا قابلًا للفتح.
- Arabic GUI tab: `مساحة العمل الذكية`.
