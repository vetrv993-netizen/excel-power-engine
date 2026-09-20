# Excel Power Engine v0.4.3 — Smart Recalculate + Excel Native Bridge

## الهدف
توسيع Smart Engine ليقرر متى تحتاج العملية إلى Excel الحقيقي لإعادة الحساب، مع الحفاظ على مسار OOXML الجراحي للعمليات التي لا تحتاج فتح Excel.

## الجديد
- `smart-recalc`: إعادة حساب ذكية عبر Excel Native عند توفر Excel على Windows.
- `smart-edit --recalculate`: إذا تطلب التعديل إعادة الحساب وكان Excel متاحًا، يستخدم Excel Native تلقائيًا.
- `open-excel`: فتح الملف بصريًا في Excel، مع إمكانية الانتقال مباشرة إلى ورقة/خلية للمراجعة البصرية.
- فحص بصمة VBA ووجود x14 قبل/بعد الحفظ Native.
- التقاط حالة الحساب قبل/بعد العملية.
- فحص Best-effort لأخطاء الصيغ الشائعة مثل `#REF!` و`#DIV/0!`.
- fallback واضح: إذا لم يتوفر Excel، لا يدّعي المحرك أنه أعاد الحساب فعليًا.

## أوامر Windows
```powershell
.\scripts\run_cli.ps1 -Command plan-edit -Path "$env:USERPROFILE\Desktop\file.xlsm" -Operation recalculate

.\scripts\run_cli.ps1 -Command smart-recalc `
  -Path "$env:USERPROFILE\Desktop\file.xlsm" `
  -Output "$env:USERPROFILE\Desktop\file_RECALCULATED.xlsm"

.\scripts\run_cli.ps1 -Command smart-edit `
  -Path "$env:USERPROFILE\Desktop\file.xlsm" `
  -Sheet "الشهر الخامس" `
  -Edit "AI11::اختبار" `
  -Output "$env:USERPROFILE\Desktop\file_EDITED.xlsm" `
  -Recalculate

.\scripts\run_cli.ps1 -Command open-excel `
  -Path "$env:USERPROFILE\Desktop\file_EDITED.xlsm" `
  -Sheet "الشهر الخامس" `
  -Cell "AI11"
```

## قاعدة التحقق البصري المعتمدة
بعد كل تعديل ناجح، يفتح الناتج في Excel وتتم مشاهدة الخلية/الخلايا المعدلة بصريًا داخل الملف. أدوات التفتيش التقنية تبقى طبقة حماية إضافية وليست بديلًا عن المشاهدة البصرية.
