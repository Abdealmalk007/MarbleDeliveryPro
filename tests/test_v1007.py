# -*- coding: utf-8 -*-
"""Data correctness and report export regressions for v10.0.7."""
from __future__ import annotations
import tempfile
import unittest
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook

from test_v1005 import app_module as app


def fake_records(n=27):
    values = []
    for index in range(n):
        values.append({
            "id":f"sample-{index:04d}",
            "created_at":"2026-11-18T12:00:00",
            "updated_at":"2026-11-18T12:00:00",
            "values":{
                "customer_name":f"عميل {index}","contract_number":f"C-{index}",
                "marble_color":("أبيض" if index%3==0 else ("أسود" if index%3==1 else "")),
                "sink_type":("مزدوج" if index%4==0 else ("مفرد" if index%4==1 else "")),
                "salesperson": ("أحمد" if index%5==0 else ("محمد" if index%5==1 else "")),
                "showroom":("المعرض الرئيسي" if index%2==0 else "المعرض الثاني"),
                "marble_edge":"مستقيمة" if index%2==0 else "دائرية",
                "delivery_status":"تم التسليم" if index%2==0 else "جديد",
            }
        })
    return values


class ReportDataAccuracy(unittest.TestCase):
    def test_version_and_reconciliation_each_dimension(self):
        self.assertEqual(app.APP_VERSION,"10.0.7")
        items=fake_records(137)
        for key,label in app.REPORT_ANALYSIS_DIMENSIONS:
            with self.subTest(label=label):
                summary=app.report_dimension_summary(items,key)
                self.assertEqual(summary["total"],137)
                self.assertEqual(sum(summary["counts"].values()),137)
                self.assertEqual(summary["missing"],sum(
                    1 for record in items if not record["values"].get(key)
                ))
                self.assertEqual(summary["counts"],app.report_dimension_counts(items,key))

    def test_excel_full_categories_with_no_top_n_cutoff(self):
        items=fake_records(135)
        for i,record in enumerate(items):
            record["values"]["marble_color"]=f"لون فريد {i}"
        for i in (1,7,11):
            items[i]["values"]["marble_color"]=""
        settings=app.AppSettings()
        with tempfile.TemporaryDirectory() as dirname:
            output=Path(dirname)/"report.xlsx"
            app.ReportExporter(settings).export_excel(items,"",output)
            wb=load_workbook(output,data_only=False)
            self.assertIn("تفاصيل التحليل",wb.sheetnames)
            self.assertGreaterEqual(len(wb["تفاصيل التحليل"]["A"]),133)
            rows=list(wb["تفاصيل التحليل"].values)[2:]
            colors=[row for row in rows if row[0]=="لون الرخام"]
            self.assertEqual(len(colors),133)
            self.assertEqual(sum(r[2] for r in colors),135)
            self.assertIn(("لون الرخام","غير محدد",3),[(r[0],r[1],r[2]) for r in colors])
            self.assertEqual(sum(r[3] for r in colors),1.0)
            # KPI sheet uses numbers and not COUNTIF against human-readable placeholders.
            for row in wb.worksheets[0].iter_rows():
                for cell in row:
                    self.assertFalse(isinstance(cell.value,str) and "COUNTIF" in cell.value)

    def test_pdf_uses_real_rows_and_wide_geometry(self):
        items=fake_records(25)
        settings=app.AppSettings()
        settings.report_orientation="landscape"
        with tempfile.TemporaryDirectory() as dirname:
            report=Path(dirname)/"landscape.pdf"
            app.ReportExporter(settings).export_pdf(items,"",report)
            self.assertTrue(report.exists())
            self.assertGreater(report.stat().st_size,1000)
            self.assertEqual(report.read_bytes()[:4],b"%PDF")
            # Dedicated PDF layout has increased usable height vs v10.0.6.
            self.assertEqual(settings.report_margin_mm,4.5)

    def test_html_has_all_chart_dimensions_and_print_settings(self):
        items=fake_records(19)
        settings=app.AppSettings()
        settings.report_page_size="A3"
        settings.report_orientation="landscape"
        with tempfile.TemporaryDirectory() as dirname:
            file=Path(dirname)/"report.html"
            app.ReportExporter(settings).export_html(items,"",file)
            html=file.read_text(encoding="utf-8")
            self.assertIn("A3 landscape",html)
            self.assertIn("4.5mm",html)
            for text in ("نوع المجلى","لون الرخام","موظف المعرض","المعرض"):
                self.assertIn(text,html)
            self.assertIn("المعرض الرئيسي",html)
            self.assertIn("مزدوج",html)


if __name__=="__main__":
    unittest.main()
