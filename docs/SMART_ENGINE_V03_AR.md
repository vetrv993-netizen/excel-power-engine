# Smart Edit Engine — v0.3

## فكرة القرار

القرار يتم على مستويين:
1. **نوع العملية**: قيمة، صيغة، تنسيق، إعادة حساب، تحليل، إنشاء تقرير، VBA، أو تعديل OOXML.
2. **ملف Excel نفسه**: XLSX/XLSM، VBA، x14، Controls، Slicers، Pivot، VML، External Links، وغيرها.

### القواعد الأساسية

| العملية | شرط الملف | المحرك |
|---|---|---|
| inspect | أي OOXML | OOXML Inspector |
| read_fast / analyze | قابل للقراءة | Fast Reader / Pandas |
| create / report | ملف جديد | XlsxWriter |
| vba_inspect | يحتوي VBA | oletools |
| cell_edit / formula_edit | XLSX بسيط بلا امتدادات حساسة | OpenPyXL |
| cell_edit / formula_edit | XLSM أو x14 أو امتدادات حساسة | OOXML Surgical |
| recalculate | Excel متاح على Windows | Excel Native |
| recalculate | لا يوجد Excel | fallback إلى OpenPyXL/OOXML مع تحذير |
| explicit XML edit | أي OOXML | OOXML Surgical |

**ملاحظة:** القرار ليس حكمًا على المكتبة، بل مطابقة للمخاطر والقدرات. OpenPyXL ممتاز للتحرير الهيكلي، لكنه ليس الاختيار الآمن الافتراضي عندما توجد امتدادات لا يستطيع الحفاظ عليها بالكامل.

## مخرجات التخطيط
`plan-edit` لا يلمس الملف. يعرض:
- operation
- selected_engine
- reason
- safety_level
- sensitive_features
- native_excel_available
- requested_recalculate

## مخرجات التنفيذ
`smart-edit` يعرض تقريرًا موحدًا حتى لو تغير المحرك تحت الغطاء.
