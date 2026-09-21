# Excel Power Engine v0.11 — منصة تحرير Excel الذكية

v0.11 تبني فوق v0.10 المستقرة ولا تستبدل محرك العمليات أو Verification Gate.

## المسار الجديد

`UI → Intent → Semantic Workbook → Target Resolution → Risk → Preview → Verification Gate → Execution → Post Verification → History/Rollback`

## أهم الإضافات

- Command Intelligence لأوامر عربية طبيعية مثل: `لوّن عمود الطلاب بالأحمر`.
- Semantic Workbook يربط الكلمات بأسماء الأعمدة الفعلية والنطاقات المكتشفة من الملف.
- Confidence/Ambiguity: لا يتم التخمين عند وجود أكثر من هدف قريب.
- Dynamic UI controls: الأوراق والأعمدة والنطاقات المولدة من المصنف تظهر كقوائم منسدلة قابلة للتحرير اليدوي.
- Smart Inspector للخلية المحددة.
- Smart Suggestions مع Preview / Apply / Ignore.
- Operation History مع حفظ JSONL محليًا.
- Risk classification SAFE / WARNING / HIGH_RISK.
- عمليات إضافية للمحرك: عرض العمود، ارتفاع الصف، إخفاء/إظهار، ومحاذاة.
- واجهة v0.11 حديثة مع Command Box علوي، حالة تنفيذ، Panels للـInspector والاقتراحات والتاريخ، وثيمات Light/Dark/Auto.

## السلامة

جميع الأوامر التي تغير المصنف تستخدم Workspace Runner الحالي، وبالتالي تستفيد من النسخة الاحتياطية وTransaction Manifest وVerification Gate وRollback.
