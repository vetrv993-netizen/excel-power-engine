# خارطة الطريق

## v0.2 — مكتملة
Deep Inspector + SafeEdit + Backup + Integrity checks.

## v0.3 — مكتملة
Smart Edit Engine لاختيار OpenPyXL أو OOXML أو Excel Native حسب طبيعة العملية.

## v0.4 — مكتملة
Smart Recalculate + Excel Native Bridge + فحوص VBA/x14 وأخطاء الصيغ.

## v0.5 — مكتملة
واجهة PySide6 عربية + Sheet viewer + Safe Edit UI + Audit log + فتح Excel بصريًا وتحديد الخلية.

## v0.6 — مكتملة
### 1) Cell Intelligence
- فهم نوع الخلية: قيمة / صيغة / خطأ / فارغة.
- قراءة الصيغة والقيمة المخزنة والتنسيق والحماية.
- استخراج المراجع المباشرة من الصيغة.
- تتبع سلسلة الاعتماد والخلايا المتأثرة.
- كشف #REF! داخل نص الصيغة والقيم الخطأ المخزنة.

### 2) Bulk Data Engine
- لصق بيانات متعددة الصفوف والأعمدة في نطاق واحد.
- معاينة كل خلية قبل التنفيذ: القديم مقابل الجديد.
- استيراد CSV/TSV وXLSX/XLSM للقراءة ثم الإدخال الشامل.
- تفسير الصيغ التي تبدأ بـ = عند الحاجة.
- تنفيذ جماعي عبر SafeEdit مع Backup وحماية VBA وx14.
- فتح Excel بعد التنفيذ وإظهار أول خلية معدلة بصريًا.

## v0.7 — مكتملة
البحث، محرك الأهداف، التنسيق والحدود والدمج والطباعة وPDF وWorkflow.

## v0.8 — مكتملة
Workbook Context للقراءة والتحليل والاقتراحات.

## v0.9 — مكتملة
Workspace Orchestrator مع backup وrollback.

## v0.10 — مكتملة
WorkbookSession، Search Result/Target Sets، keyed-bulk، Verification Gate، وtransaction manifest.

## حالة الاختبارات
تغطي suite الحالية 41 اختبارًا بعد إضافة regression لـWorkspace Search وSearch→Target→Bulk وVerification/Keyed Mapping.
