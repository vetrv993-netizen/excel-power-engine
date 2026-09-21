# Excel Power Engine v0.11

منصة Excel/XLSM ذكية متعددة المسارات مع فهم أوامر طبيعية، فهم دلالي للمصنف، تحرير يدوي، بحث، تنسيق، حدود، دمج، طباعة، PDF وواجهة عربية حديثة.

## أبرز v0.11
- Command Box عربي يحوّل الأوامر الطبيعية إلى Intent وخطة تنفيذ قابلة للمعاينة.
- Semantic Workbook يربط أسماء الأعمدة الحقيقية بالنطاقات المكتشفة، مع Confidence/Ambiguity.
- قوائم منسدلة ديناميكية للأوراق والأعمدة والنطاقات والخلايا المكتشفة من الملف المرفوع.
- Smart Inspector + Smart Suggestions + Operation History.
- عمليات جديدة: عرض العمود، ارتفاع الصف، إخفاء/إظهار، محاذاة والتفاف النص.
- نفس Workspace/Verification/Backup/Rollback الخاص بـ v0.10 لكل تنفيذ ذكي.

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

# فهم أمر طبيعي دون تنفيذ
excel-power command-plan file.xlsx "لوّن عمود الطلاب بالأحمر" --sheet "الطلاب"

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

## v0.11 — Smart Excel Platform

تضيف v0.11 طبقة فهم أوامر طبيعية فوق محرك v0.10، مع فهم دلالي لأسماء الأعمدة، قوائم منسدلة مشتقة من الملف، Preview، تصنيف المخاطر، Smart Inspector، الاقتراحات الذكية، وسجل العمليات، مع الحفاظ على Verification Gate وRollback.

راجع `docs/V11_SMART_PLATFORM_AR.md`.
