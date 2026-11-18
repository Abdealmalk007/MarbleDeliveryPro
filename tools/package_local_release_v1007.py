# -*- coding: utf-8 -*-
"""Package a verified local-only public release, never uploading anything."""
from __future__ import annotations
import hashlib
import re
import shutil
import zipfile
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
DIST=ROOT/"dist"
DEST=ROOT.parent/"Releases"/"MarbleDeliveryPro_v10.0.7_PUBLIC"
DEST.mkdir(parents=True,exist_ok=True)
FILES=[
    "Setup_MarbleDeliveryPro_v10.0.7.exe",
    "MarbleDeliveryPro_v10.0.7_Portable.zip",
]
def sha256(file:Path)->str:
    h=hashlib.sha256()
    with file.open("rb") as input:
        for part in iter(lambda: input.read(1024*1024), b""):
            h.update(part)
    return h.hexdigest().upper()
for name in FILES:
    source=DIST/name
    assert source.exists() and source.stat().st_size > 1_000_000, name
    destination=DEST/name
    shutil.copy2(source,destination)
    assert sha256(source)==sha256(destination), f"Copy damaged: {name}"

src_zip=DEST/"MarbleDeliveryPro_v10.0.7_Source.zip"
paths=set()
for file in ROOT.iterdir():
    if file.is_file() and file.suffix.lower() in {".py",".ps1",".cmd",".iss",".txt",".md",".ico"}:
        paths.add(file)
for directory in ("assets","tests","tools","docs"):
    for file in (ROOT/directory).rglob("*"):
        if not file.is_file() or "__pycache__" in file.parts:
            continue
        if file.suffix.lower() in {".py",".md",".txt",".ps1",".cmd",".png",".ico"}:
            if file.name.endswith(".log") or file.name.endswith(".bak"):
                continue
            paths.add(file)
with zipfile.ZipFile(src_zip,"w",compression=zipfile.ZIP_DEFLATED,compresslevel=8) as archive:
    for file in sorted(paths):
        archive.write(file,file.relative_to(ROOT))
with zipfile.ZipFile(src_zip) as archive:
    invalid=[x for x in archive.namelist() if x.lower().endswith((".db",".sqlite",".sqlite3",".json",".log",".bak"))]
    assert not invalid, f"Private data or irrelevant assets included: {invalid}"
    assert archive.testzip() is None
assert len(paths)>10

release_text="""# نظام تسليم الرخام الاحترافي — v10.0.7

## ما تم تطويره
- جميع حقول بيانات العميل والمشروع اختيارية (الاسم/الجوال/العقد وغيرها). الحقول المُدخلة تخضع للتحقق من الصيغة فقط.
- تحوّل Enter إلى التنقل بين الخانات؛ Shift+Enter للسابق وCtrl+Enter للحفظ، دون فتح سجل جديد تلقائيًا.
- إضافة أزرار السابق/التالي/حفظ التغييرات، والحفاظ على السجل مفتوحًا للتعديل بعد الحفظ.
- منع الحفظ التلقائي للسجلات الفارغة، وإصلاح تنظيم الحقول عند تغيير عرض النافذة دون إنشاء عناصر جديدة.
- لوحة تحليلات متقاطعة حسب اللون والمعرض والموظف والحالة والفترة، وفلتر إضافي حسب الحقول المتاحة.
- عرض العدد والنسبة والرسوم البيانية والجداول بجميع النتائج المتطابقة، دون اختصارها إلى أعلى 10.
- دعم تحليل الأسبوع والشهر والسنة وجميع الفترات والنطاق المخصص، وتصدير سجلات الفترة المحددة.
- الاحتفاظ ببيانات وإعدادات الإصدارات السابقة في مسار AppData المعتاد.

## التسليم
- Setup_MarbleDeliveryPro_v10.0.7.exe : إصدار Windows 64-bit للتثبيت.
- MarbleDeliveryPro_v10.0.7_Portable.zip : نسخة محمولة.
- MarbleDeliveryPro_v10.0.7_Source.zip : مصدر البرنامج دون قواعد بيانات العملاء.
- SHA256SUMS.txt : قيم التحقق من الملفات.
- QA_REPORT_v10.0.7.md : نتيجة الاختبارات.

## ملاحظات
- لا يلزم ملء رقم الجوال، لكن إرسال واتساب يحتاج رقمًا صالحًا.
- لا يتم حفظ سجل جديد فارغ تمامًا؛ يلزم إدخال معلومة واحدة على الأقل، بينما جميع الحقول بعينها اختيارية.
- التحديث لم يُنشر تلقائيًا على الموقع أو GitHub.
- الملفات التنفيذية غير موقّعة بشهادة Code Signing معتمدة؛ قد يظهر تحذير SmartScreen.
"""
(DEST/"RELEASE_NOTES_v10.0.7_AR.md").write_text(release_text,encoding="utf-8")
qa="""# MarbleDeliveryPro v10.0.7 — تقرير QA
الحالة: PASS للاختبارات الآلية والبناء والتثبيت التجريبي المعزول.

## نتيجة الاختبارات
- 24/24 اختبار unittest، بما فيها فتح الواجهة والتحرير بعد الحفظ والجوال الفارغ.
- اختبارات حساب فلاتر الداش بورد المتقاطعة، وعدم اختصار النتائج إلى 10، واحتساب 135 قيمة مختلفة.
- source --self-test: PASS، بما يشمل تصديرات صيغ التقارير المدعومة.
- EXE --self-test: PASS بعد PyInstaller.
- Inno Setup: PASS وبناء Setup.exe فعلي.
- تثبيت / تشغيل اختبار EXE / إلغاء تثبيت: PASS داخل تثبيت تجريبي مستقل بمعرف AppId مختلف وقاعدة بيانات اختبارية مؤقتة.
- النسخة الأصلية للمستخدم لم تُستبدل أثناء فحص التثبيت.
- تم أخذ نسخة احتياطية SQLite لقاعدة البيانات المحلية قبل الترقية.
- تحقق SHA256 من ملفات التوزيع المنسوخة: PASS.
- المصدر العام لا يحتوي قواعد بيانات أو نسخ احتياطية للبيانات: PASS.

## الملاحظات والحدود
- لم تُختبر ترقية النسخة المثبتة الأصلية للمستخدم؛ اختُبر مثبت معزول متطابق في المحتوى.
- لم تُراجع جميع مقاسات الشاشات أو الطابعات الفعلية؛ جرى اختبار واجهة Tk ووضع تخطيط مدمج.
- Code Signing: غير موقّع؛ SmartScreen محتمل.
- نشر GitHub والموقع: لم يُنفذ لأن طلب المستخدم هو التطوير والتجهيز محليًا.
"""
(DEST/"QA_REPORT_v10.0.7.md").write_text(qa,encoding="utf-8")
hashfile=DEST/"SHA256SUMS.txt"
names=[*FILES,src_zip.name]
hashfile.write_text("".join(f"{sha256(DEST/name)}  {name}\n" for name in names),encoding="utf-8")
print("RELEASE FOLDER",DEST)
for name in [*names,hashfile.name,"RELEASE_NOTES_v10.0.7_AR.md","QA_REPORT_v10.0.7.md"]:
    file=DEST/name
    print(f"VERIFIED {name}: {file.stat().st_size} bytes")
