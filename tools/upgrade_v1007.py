# -*- coding: utf-8 -*-
"""Atomic, checked patch: MarbleDeliveryPro v10.0.6 to v10.0.7."""
from __future__ import annotations
from pathlib import Path
import re

ROOT=Path(__file__).resolve().parents[1]
app_files=[p for p in ROOT.glob("*.py") if "تسليم" in p.name]
assert len(app_files)==1,app_files
source=app_files[0]
baseline=source.read_text(encoding="utf-8-sig")
assert 'APP_VERSION = "10.0.6"' in baseline
(ROOT/"docs"/"SOURCE_BASELINE_v10.0.6.bak").write_text(baseline,encoding="utf-8")
body=baseline
revisions=[]
def replace(old,new,label,expected=1):
    global body
    found=body.count(old)
    if found!=expected:
        raise AssertionError(f"PATCH {label}: expected {expected} but saw {found}")
    body=body.replace(old,new)
    revisions.append(label)
def rx(pattern,replacement,label,expected=1,flags=0):
    global body
    body,count=re.subn(pattern,replacement,body,flags=flags)
    if count!=expected:
        raise AssertionError(f"REGEX {label}: expected {expected} got {count}")
    revisions.append(label)

replace('APP_VERSION = "10.0.6"','APP_VERSION = "10.0.7"',"Version")
replace('report_margin_mm: float = 6.0','report_margin_mm: float = 4.5',"New default margin")
replace('values.get("report_margin_mm", 6.0)','values.get("report_margin_mm", 4.5)',"Margin fallback")

anchor='def balanced_report_field_groups(fields: list[Any], max_columns: int = 3)'
assert body.count(anchor)==1
analysis='''def report_dimension_counts(records: list[dict[str, Any]], key: str) -> dict[str, int]:
    """Use exactly the exported rows and disclose missing values, never guess."""
    result = Counter()
    for row in records:
        value = clean_value(row.get("values", {}).get(key, "")) or "غير محدد"
        result[value] += 1
    assert sum(result.values()) == len(records)
    return dict(sorted(result.items(), key=lambda item: (-item[1], item[0])))


def report_dimension_summary(records: list[dict[str, Any]], key: str) -> dict[str, Any]:
    counts = report_dimension_counts(records, key)
    leader = next(((name, n) for name, n in counts.items() if name != "غير محدد"),
                  ("غير محدد", 0))
    return {"counts": counts, "missing": counts.get("غير محدد", 0),
            "distinct": len(counts) - int("غير محدد" in counts),
            "leader": leader, "total": len(records)}


REPORT_ANALYSIS_DIMENSIONS: tuple[tuple[str, str], ...] = (
    ("marble_color", "لون الرخام"), ("sink_type", "نوع المجلى"),
    ("showroom", "المعرض"), ("salesperson", "موظف المعرض"),
    ("marble_edge", "شفة الرخام"), ("delivery_status", "حالة التسليم"),
)


'''
body=body.replace(anchor,analysis+anchor)
revisions.append("Consistent dimension aggregation")
for key,old,new in (
 ("color",'"marble_color": (9.0, 14.0, 0.92)', '"marble_color": (11.0, 18.0, 1.06)'),
 ("sink",'"sink_type": (8.0, 12.0, 0.84)', '"sink_type": (10.0, 16.0, 1.05)'),
 ("showroom",'"showroom": (8.0, 12.5, 0.84)', '"showroom": (10.5, 17.0, 1.06)'),
 ("salesperson",'"salesperson": (12.0, 18.0, 1.02)', '"salesperson": (13.0, 20.0, 1.07)')
):
    replace(old,new,"Column width "+key)

replace('if f.key in ["showroom", "marble_color", "salesperson"]:',
        'if f.key in {key for key, _ in REPORT_ANALYSIS_DIMENSIONS}:',
        "HTML categories include sink, edge, status")
html_old='''                counts = {}
                for r in records:
                    val = clean_value(r.get("values", {}).get(f.key, "غير محدد")) or "غير محدد"
                    counts[val] = counts.get(val, 0) + 1
                if counts:
                    analytics[f.label] = counts'''
replace(html_old,
'''                counts = report_dimension_counts(records, f.key)
                if counts:
                    analytics[f.label] = counts''',
        "Shared HTML aggregation")
replace('''        html = f"""<!DOCTYPE html>''',
'''        print_page_size, _, _, print_orientation = self._resolve_report_page(fields)
        print_paper = str(s.report_page_size).upper()
        if print_paper == "CUSTOM":
            print_width, print_height = self._page_mm(print_page_size)
            print_paper_css = f"{print_width:.1f}mm {print_height:.1f}mm"
        else:
            print_paper_css = f"{'Letter' if print_paper == 'LETTER' else print_paper} {print_orientation}"
        print_margin = max(3.0, float(s.report_margin_mm))
        html = f"""<!DOCTYPE html>''',
        "HTML print format matches settings")
replace('@page {{ size: A4 landscape; margin: 3mm 2.5mm; }}',
        '@page {{ size: {print_paper_css}; margin: {print_margin:g}mm; }}',
        "HTML dynamic print geometry")

replace('''        use_portrait = count <= 4
        ws.sheet_properties.pageSetUpPr.fitToPage = True''',
        '''        _, _, _, direction = self._resolve_report_page(fields)
        use_portrait = direction == "portrait"
        ws.sheet_properties.pageSetUpPr.fitToPage = True''',
        "Excel landscape/portrait matches PDF")
replace('ws.page_setup.paperSize = ws.PAPERSIZE_A4',
        '''ws.page_setup.paperSize = {
            "A3": ws.PAPERSIZE_A3, "LETTER": ws.PAPERSIZE_LETTER,
        }.get(str(self.settings.report_page_size).upper(), ws.PAPERSIZE_A4)''',
        "Excel real paper sizes")
replace('''        self._add_excel_summary(wb, ws, records, fields, palette)
        output.parent.mkdir(parents=True, exist_ok=True)''',
        '''        self._add_excel_summary(wb, ws, records, fields, palette)
        self._add_excel_full_analysis(wb, records, fields, palette)
        output.parent.mkdir(parents=True, exist_ok=True)''',
        "Extra Excel full breakdown sheet")
# Existing sheet used COUNTIF formula with a fake "غير محدد" category: Excel blanks were incorrectly zero.
rx(r'            counts: dict\[str, int\] = \{\}\n            for record in records:\n                value = clean_value\(record.get\("values", \{\}\).get\(key, ""\)\) or "غير محدد"\n                counts\[value\] = counts.get\(value, 0\) \+ 1\n            top_items = sorted\(counts.items\(\), key=lambda pair: \(-pair\[1\], pair\[0\]\)\)\[:8\]',
   '            counts = report_dimension_counts(records, key)\n            top_items = list(counts.items())[:8]',
   "Correct bounded chart summary source")
rx(r'            data_col_letter = get_column_letter\(field_columns\[key\]\)\n            for offset, \(category, _count\) in enumerate\(top_items, header_row \+ 1\):\n                summary.cell\(offset, start_col, category\)\n                summary.cell\(offset, start_col \+ 1, f\'=COUNTIF\([^\n]+\n',
   '            for offset, (category, actual_count) in enumerate(top_items, header_row + 1):\n                summary.cell(offset, start_col, category)\n                summary.cell(offset, start_col + 1, actual_count)\n',
   "Accurate Excel counts including blank categories")

marker='    def _add_excel_summary(self, workbook: Any, data_sheet: Any, records: list[dict[str, Any]],'
assert body.count(marker)==1
excel_method='''    def _add_excel_full_analysis(
        self, workbook: Any, records: list[dict[str, Any]],
        fields: list[FieldDefinition], palette: dict[str, str],
    ) -> None:
        """Every distinct field value, share, and explicit missing count."""
        sheet = workbook.create_sheet("تفاصيل التحليل")
        sheet.sheet_view.rightToLeft = True
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A3"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.append(["تحليل كامل مطابق للسجلات المصدّرة، دون اختصار"])
        sheet.append(["الحقل", "القيمة الفعلية", "عدد السجلات", "النسبة", "الموجود", "غير المحدد"])
        sheet.merge_cells("A1:F1")
        sheet["A1"].font = Font(name=self.settings.font_family, size=15, bold=True,
                                color=palette["title_text"].lstrip("#"))
        sheet["A1"].fill = PatternFill("solid", fgColor=palette["title_fill"].lstrip("#"))
        sheet["A1"].alignment = Alignment(horizontal="center")
        for cell in sheet[2]:
            cell.font = Font(name=self.settings.font_family, size=11, bold=True,
                             color=palette["header_text"].lstrip("#"))
            cell.fill = PatternFill("solid", fgColor=palette["header_fill"].lstrip("#"))
            cell.alignment = Alignment(horizontal="center")
        present = {field.key: field.label for field in fields}
        total = len(records)
        for key, label in REPORT_ANALYSIS_DIMENSIONS:
            if key not in present:
                continue
            overview = report_dimension_summary(records, key)
            for value, count in overview["counts"].items():
                sheet.append([present[key] or label, value, count,
                              count / total if total else 0,
                              total - overview["missing"], overview["missing"]])
                sheet.cell(sheet.max_row, 4).number_format = "0.00%"
        sheet.auto_filter.ref = f"A2:F{max(2, sheet.max_row)}"
        for col, width in {"A": 25, "B": 35, "C": 19, "D": 12, "E": 18, "F": 18}.items():
            sheet.column_dimensions[col].width = width
        sheet.print_title_rows = "1:2"
        sheet.print_area = f"A1:F{max(2, sheet.max_row)}"


'''
body=body.replace(marker,excel_method+marker)
revisions.append("Complete Excel sheet including all values")

replace('topMargin=39*mm, bottomMargin=24*mm,',
        'topMargin=34*mm, bottomMargin=17*mm,',
        "More printable PDF height")
replace('cell_padding = max(4.4, font_size * 0.50)',
        'cell_padding = max(3.0, font_size * 0.40)',
        "Readable compact table cells")
replace('table_height_limit = doc.height - meta_height - 3*mm - 14',
        'table_height_limit = doc.height - meta_height - 2*mm - 11',
        "More rows per PDF page")
replace('story.extend([meta, Spacer(1, 3*mm), table])',
        'story.extend([meta, Spacer(1, 2*mm), table])',
        "Smaller table gap")
replace('"PDF v10.0.4 page %s/%s:',
        '"PDF v10.0.7 page %s/%s:',
        "Correct logging version")

# PDF summary is based on the whole export, not the current page or stale analytics.
pdf_meta_old='''        def build_meta(page_records: list[dict[str, Any]]) -> Table:
            page_w_mm, page_h_mm = self._page_mm(page_size)
            meta = Table(
                [[
                    para(f"السجلات: {metrics['records']}", right_style),
                    para(f"العقود: {metrics['contracts']}", right_style),
                    para(f"العقود المتكررة: {metrics['multi_kitchen_contracts']}", right_style),
                    para(f"هذه الصفحة: {len(page_records)} | {now:%Y-%m-%d}", right_style),
                ]],
                colWidths=[available_width/4] * 4,
            )'''
pdf_meta_new='''        dimension_summaries = {
            key: report_dimension_summary(records, key)
            for key, _ in REPORT_ANALYSIS_DIMENSIONS
        }

        def build_meta(page_records: list[dict[str, Any]]) -> Table:
            def top_label(key: str, caption: str) -> str:
                summary = dimension_summaries[key]
                name, count = summary["leader"]
                return f"{caption}: {name} ({count}) | غير محدد {summary['missing']}"
            meta = Table(
                [
                    [
                        para(f"السجلات: {metrics['records']}", right_style),
                        para(f"العقود: {metrics['contracts']}", right_style),
                        para(f"العقود المتكررة: {metrics['multi_kitchen_contracts']}", right_style),
                        para(f"هذه الصفحة: {len(page_records)} | {now:%Y-%m-%d}", right_style),
                    ],
                    [
                        para(top_label("marble_color", "اللون"), right_style),
                        para(top_label("sink_type", "المجلى"), right_style),
                        para(top_label("salesperson", "الموظف"), right_style),
                        para(top_label("showroom", "المعرض"), right_style),
                    ],
                ],
                colWidths=[available_width / 4] * 4,
            )'''
replace(pdf_meta_old,pdf_meta_new,"PDF accurate categorical meta")
replace('("TOPPADDING", (0, 0), (-1, -1), 5.0),','("TOPPADDING", (0, 0), (-1, -1), 3.2),',"Compact summary top")
replace('("BOTTOMPADDING", (0, 0), (-1, -1), 5.0),','("BOTTOMPADDING", (0, 0), (-1, -1), 3.2),',"Compact summary bottom")

# Header/footer positions kept strictly outside the frame (top margin=34, bottom=17).
for key,old,new in [
 ("PDF title background","canvas.rect(0, page_height - 22*mm, page_width, 22*mm, stroke=0, fill=1)","canvas.rect(0, page_height - 17*mm, page_width, 17*mm, stroke=0, fill=1)"),
 ("PDF subtitle background","canvas.rect(0, page_height - 35*mm, page_width, 13*mm, stroke=0, fill=1)","canvas.rect(0, page_height - 30*mm, page_width, 13*mm, stroke=0, fill=1)"),
 ("PDF title bar","canvas.rect(0, page_height - 36*mm, page_width, 1.3*mm, stroke=0, fill=1)","canvas.rect(0, page_height - 30.6*mm, page_width, 0.8*mm, stroke=0, fill=1)"),
 ("PDF title font",'title_size = 24 if str(self.settings.report_page_size).upper() == "A4" else 27','title_size = 19 if str(self.settings.report_page_size).upper() == "A4" else 23'),
 ("PDF title baseline","canvas.drawCentredString(page_width/2, page_height - 10*mm, shape_arabic(self.settings.report_title))","canvas.drawCentredString(page_width/2, page_height - 8*mm, shape_arabic(self.settings.report_title))"),
 ("PDF en baseline","canvas.drawCentredString(page_width/2, page_height - 18*mm, self.settings.report_title_en)","canvas.drawCentredString(page_width/2, page_height - 15*mm, self.settings.report_title_en)"),
 ("PDF subtitle baseline","page_width/2, page_height - 29*mm,","page_width/2, page_height - 24*mm,"),
 ("PDF logo","self.settings.logo_path, side_margin, page_height - 31*mm,","self.settings.logo_path, side_margin, page_height - 26*mm,"),
 ("PDF logo max height","height=min(27*mm, logo_width * 0.72),","height=min(20*mm, logo_width * 0.72),"),
 ("PDF QR","canvas, page_width - side_margin - qr_size, page_height - 22*mm)","canvas, page_width - side_margin - qr_size, page_height - 20*mm)"),
 ("PDF recipient","page_width - side_margin, 20*mm,","page_width - side_margin, 13*mm,"),
 ("PDF signature",'canvas.drawCentredString(page_width/2, 20*mm,','canvas.drawCentredString(page_width/2, 13*mm,'),
 ("PDF date","canvas.drawString(side_margin, 20*mm,","canvas.drawString(side_margin, 13*mm,"),
 ("PDF first footer line","canvas.line(side_margin, 16*mm, page_width - side_margin, 16*mm)","canvas.line(side_margin, 10*mm, page_width - side_margin, 10*mm)"),
 ("PDF second footer line","canvas.line(side_margin, 11*mm, page_width - side_margin, 11*mm)","canvas.line(side_margin, 7*mm, page_width - side_margin, 7*mm)"),
 ("PDF footer note","canvas.drawRightString(page_width - side_margin, 6*mm,","canvas.drawRightString(page_width - side_margin, 3*mm,"),
 ("PDF page count","canvas.drawString(side_margin, 6*mm,","canvas.drawString(side_margin, 3*mm,"),
 ("PDF footer ref","page_width/2, 6*mm,","page_width/2, 3*mm,"),
]:
    replace(old,new,key)

# Stop preview from changing report archive and make all previews nonblocking.
replace('*, after_success=None, show_success=True):',
        '*, after_success=None, show_success=True, archive=True):',
        "Archive policy flag")
replace('''                DatabaseManager.archive_report(
                    output_path, file_type, settings_snapshot, recipient, records_snapshot, fields_snapshot
                )''',
        '''                if archive:
                    DatabaseManager.archive_report(
                        output_path, file_type, settings_snapshot, recipient, records_snapshot, fields_snapshot
                    )''',
        "Preview is read-only")
replace('''after_success=open_with_default_app, show_success=False,
        )''',
        '''after_success=open_with_default_app, show_success=False, archive=False,
        )''',
        "Preview main window never archives")
preview_old='''        preview_path = PREVIEW_DIR / "معاينة_تقرير_تسليم_الرخام.pdf"
        progress = ProgressDialog(self.parent_app, "جاري تجهيز المعاينة...")
        try:
            ReportExporter(self.settings).export_pdf(
                records, self.parent_app.recipient_var.get(), preview_path, fields
            )
            open_with_default_app(preview_path)
            self.parent_app.status_var.set("تم فتح معاينة PDF")
        except Exception as exc:
            logging.exception("فشلت معاينة التقرير")
            messagebox.showerror("تعذر فتح المعاينة", str(exc), parent=self)
        finally:
            progress.close()'''
preview_new='''        preview_path = PREVIEW_DIR / f"معاينة_تقرير_تسليم_الرخام_{dt.datetime.now():%Y%m%d_%H%M%S_%f}.pdf"
        self.parent_app._start_export_task(
            "pdf", records, self.parent_app.recipient_var.get(), fields, preview_path,
            after_success=open_with_default_app, show_success=False, archive=False,
        )'''
replace(preview_old,preview_new,"Export options preview asynchronously")

compile(body,str(source),"exec")
source.write_text(body,encoding="utf-8")
script=ROOT/"Build_All.ps1"
raw=script.read_text(encoding="utf-8-sig")
assert raw.count("10.0.6")==1
script.write_text(raw.replace("10.0.6","10.0.7"),encoding="utf-8-sig")
setup=next(ROOT.glob("*.iss"))
raw=setup.read_text(encoding="utf-8-sig")
assert raw.count("10.0.6")>=3
setup.write_text(raw.replace("10.0.6","10.0.7"),encoding="utf-8-sig")
for path in (ROOT/"tests").glob("test*.py"):
    raw=path.read_text(encoding="utf-8-sig")
    path.write_text(raw.replace('"10.0.6"','"10.0.7"').replace("'10.0.6'","'10.0.7'"),
                    encoding="utf-8")
print(f"PATCH SUCCESS {len(revisions)} validated modifications")
for index,label in enumerate(revisions,1):
    print(f"{index:02d}. {label}")
