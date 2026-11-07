"""Isolated smoke/regression checks; never reads or alters installed user data."""
from __future__ import annotations
import importlib.util
import os
import tempfile
import unittest
from pathlib import Path
from uuid import uuid4

_ROOT = Path(__file__).resolve().parents[1]
_SOURCE = _ROOT / 'برنامج تسليم الرخام الاحترافي.py'
_TEST_DATA = tempfile.TemporaryDirectory(prefix='marble-v1004-qa-')
os.environ['MARBLE_DELIVERY_DATA_DIR'] = _TEST_DATA.name
_spec = importlib.util.spec_from_file_location('marble_delivery_qa_v1004', _SOURCE)
app_module = importlib.util.module_from_spec(_spec)
import sys
sys.modules[_spec.name] = app_module
_spec.loader.exec_module(app_module)


class SourceRegression(unittest.TestCase):
    def test_source_version_and_schema(self):
        self.assertEqual(app_module.APP_VERSION, '10.0.5')
        self.assertEqual(app_module.SCHEMA_VERSION, 14)
        with app_module.get_db() as conn:
            self.assertEqual(conn.execute('PRAGMA quick_check').fetchone()[0], 'ok')
            self.assertTrue(conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='records'").fetchone())

    def test_hidden_delivery_date_compatible(self):
        settings = app_module.AppSettings()
        self.assertNotIn('scheduled_delivery_date', [f.key for f in settings.report_fields])
        self.assertIn('scheduled_delivery_date', [f.key for f in settings.fields])

    def test_numeric_arabic_and_date_sorting(self):
        k=app_module.report_table_sort_value
        self.assertLess(k('2'), k('15'))
        self.assertLess(k('٢٠'), k('٣٠'))
        self.assertLess(k('2026-02-01'), k('2026-10-03'))
        self.assertEqual(k('١٢'),k('12'))
        self.assertLess(k('أحمد'),k(''))

    def test_password_verification_not_changed(self):
        salt, digest = app_module.password_hash('Qa-Test-Strong-Password!')
        self.assertTrue(app_module.password_matches('Qa-Test-Strong-Password!', salt, digest))
        self.assertFalse(app_module.password_matches('wrong', salt, digest))

    def test_database_record_roundtrip(self):
        now=app_module.dt.datetime.now().isoformat(timespec='seconds')
        record={'id':str(uuid4()),'created_at':now,'updated_at':now,'values':{
            'contract_number':'QA-001','customer_name':'عميل فحص','mobile':'0500000000',
            'delivery_status':'جديد','scheduled_delivery_date':'2026-12-01',
        }}
        app_module.DatabaseManager.save_record(record)
        records=app_module.DatabaseManager.load_records()
        copy=next(v for v in records if v['id']==record['id'])
        self.assertEqual(copy['values']['scheduled_delivery_date'], '2026-12-01')


class TkSmoke(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest('GUI display unavailable; run tests with Xvfb on Linux or desktop Windows')
        self.root.withdraw()
        self.app = app_module.MarbleDeliveryApp(self.root)

    def tearDown(self):
        if hasattr(self, 'root'):
            try: self.root.destroy()
            except Exception: pass

    def test_action_toolbar_and_tree_sort_ascending_descending(self):
        import tkinter as tk
        # Locate buttons recursively, even when packed in nested Tk frames.
        def all_buttons(widget):
            for child in widget.winfo_children():
                if isinstance(child, tk.Button): yield child
                yield from all_buttons(child)
        texts = {b.cget('text') for b in all_buttons(self.root)}
        for label in ('+ سجل جديد','حفظ السجل','استيراد Excel','تصدير PDF','تصدير Excel','نسخة احتياطية','تحديث'):
            self.assertIn(label, texts)
        self.app.records = [
            {'id':'qa-2','values':{'contract_number':'20','customer_name':'ج'}} ,
            {'id':'qa-1','values':{'contract_number':'3','customer_name':'ب'}} ,
        ]
        self.app.current_report_record_ids = None
        self.app._apply_filters()
        self.app._sort_table_by('contract_number')
        self.assertEqual(list(self.app.tree.get_children()), ['qa-1','qa-2'])
        self.app._sort_table_by('contract_number')
        self.assertEqual(list(self.app.tree.get_children()), ['qa-2','qa-1'])
        self.assertEqual(self.app.filtered_records[0]['id'], 'qa-2')
        self.app.tree.selection_set('qa-1')
        self.app._update_selection_count()
        self.assertEqual(self.app.selection_count_var.get(),'المحدد: 1')


if __name__=='__main__':
    unittest.main(verbosity=2)
