# Excel Power Engine v0.6

## ذكاء الخلية
الأمر:
```powershell
excel-power cell-info <file> --sheet "الشهر الثاني" --cell V266 --trace
```
يعرض نوع الخلية، الصيغة، القيمة المخزنة، الخطأ إن وجد، وسلسلة الاعتماد المباشرة.

## فحص الصيغ
```powershell
excel-power formula-scan <file.xlsm>
```
يكشف #REF! داخل نصوص الصيغ والقيم الخطأ المخزنة.

## معاينة الإدخال الشامل
البيانات يمكن أن تكون CSV/TSV أو XLSX/XLSM.
```powershell
excel-power bulk-preview <file> --sheet "الشهر الخامس" --start-cell F11 --data-file "زيارات.csv"
```

## تنفيذ الإدخال الشامل
```powershell
excel-power bulk-edit <file> --sheet "الشهر الخامس" --start-cell F11 --data-file "زيارات.csv" --output "نسخة_الادخال.xlsm"
```

المحرك يعرض التغييرات قبل التنفيذ، ثم يكتب عبر SafeEdit/OOXML في ملفات XLSM الحساسة ويحافظ على VBA وx14.

## نطاقات متعددة
يمكن توزيع الإدخال على أكثر من بلوك دون لمس الصفوف الفاصلة. مثال:
```powershell
-ranges "F11:I70;F76:I135;F141:I200;F206:I265;F271:I330"
```
ويتم توزيع الصفوف المدخلة بالتسلسل على هذه النطاقات. الخانات الفارغة لا تمسح البيانات افتراضيًا إلا مع خيار المسح.
