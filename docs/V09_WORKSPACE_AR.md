# v0.9 — مساحة العمل الذكية والمنسق الموحد

v0.9 يضيف طبقة Workspace Orchestrator توحد سياق الملف وتنفيذ العمليات في معاملة واحدة.

## العمليات
- search: بحث واستخدام النتائج كهدف للعملية التالية.
- bulk: إدخال جماعي إلى خلية بداية أو عدة نطاقات.
- format / borders / merge / unmerge: عمليات تنسيق وهيكل.
- print-setup / pdf: إعداد الطباعة والتصدير.
- open-excel: تحقق بصري في Excel.

## المسار
`Preview → Backup → Execute → Verify → Visual Check`

كل عملية لها نتيجة، ويمكن استعادة نسخة العملية من `*_TRANSACTION_BACKUP.*` باستخدام `workspace-rollback`.
