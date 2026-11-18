# -*- coding: utf-8 -*-
"""Acceptance tests for v10.0.6 optional-input/editing/dashboard release."""
from __future__ import annotations
import copy
import datetime as dt
import unittest
from unittest.mock import patch
from uuid import uuid4

from test_v1005 import app_module as app


def row(identifier: str, *, color: str = "", showroom: str = "", employee: str = "",
        status: str = "جديد", created: str = "2026-11-18") -> dict:
    return {"id": identifier, "created_at": created + "T12:00:00",
            "updated_at": created + "T12:00:00",
            "values": {"marble_color": color, "showroom": showroom,
                       "salesperson": employee, "delivery_status": status}}


class DataAndAnalysisTests(unittest.TestCase):
    def test_all_fields_optional_for_new_and_legacy_settings(self):
        self.assertEqual(app.APP_VERSION, "10.0.7")
        self.assertTrue(all(not f.required for f in app.default_fields()))
        legacy = app.AppSettings().to_dict()
        for field in legacy["fields"]:
            field["required"] = True
        legacy["fields"].append({
            "key": "legacy_custom", "label": "مخصص", "required": True,
            "visible": True, "kind": "text", "system": False,
        })
        migrated = app.AppSettings.from_dict(legacy)
        self.assertTrue(all(not f.required for f in migrated.fields))
        self.assertIn("legacy_custom", [f.key for f in migrated.fields])

    def test_partial_record_roundtrip_with_empty_phone(self):
        key = "qa-optional-" + uuid4().hex
        record = {"id": key, "created_at": "2026-11-18T11:00:00",
                  "updated_at": "2026-11-18T11:00:00",
                  "values": {"customer_name": "اسم عميل فقط", "mobile": "",
                             "showroom": "", "contract_number": ""}}
        app.DatabaseManager.save_record(record)
        actual = next(item for item in app.DatabaseManager.load_records() if item["id"] == key)
        self.assertEqual(actual["values"]["customer_name"], "اسم عميل فقط")
        self.assertEqual(actual["values"]["mobile"], "")

    def test_dashboard_cross_filters_all_values_including_unset(self):
        rows = [
            row("one", color="أبيض", showroom="الشرق", employee="سعيد", status="تم التسليم"),
            row("two", color="أسود", showroom="الشرق", employee="سعيد"),
            row("three", color="أبيض", showroom="الغرب", employee="خالد"),
            row("four", showroom="الغرب", employee="سعيد"),
        ]
        f = app.filter_dashboard_records
        self.assertEqual(len(f(rows)), 4)
        self.assertEqual([r["id"] for r in f(rows, filters={"marble_color": "أبيض"})], ["one", "three"])
        self.assertEqual([r["id"] for r in f(rows, filters={"marble_color": "أبيض", "showroom": "الشرق"})], ["one"])
        self.assertEqual([r["id"] for r in f(rows, filters={"marble_color": "غير محدد"})], ["four"])
        self.assertEqual(len(f(rows, filters={"salesperson": "سعيد"})), 3)
        self.assertEqual(len(f(rows, filters={"salesperson": "الكل"})), 4)
        self.assertEqual(len(f(rows, filters={"delivery_status": "تم التسليم"})), 1)

    def test_dashboard_period_dates_and_dedup(self):
        rows = [
            row("nov", color="أبيض", created="2026-11-18"),
            row("oct", color="أسود", created="2026-10-21"),
            row("nov", color="أبيض", created="2026-11-18"),
            row("no-date", created=""),
        ]
        f = app.filter_dashboard_records
        period = f(rows, start=dt.date(2026, 11, 1), end=dt.date(2026, 11, 30),
                   basis="تاريخ إنشاء السجل")
        self.assertEqual([item["id"] for item in period], ["nov"])
        self.assertEqual(len(f(rows)), 3)

    def test_dashboard_never_limits_to_top_ten(self):
        rows = [row(str(i), color=f"لون {i}", showroom="معرض",
                    employee=f"موظف {i}") for i in range(135)]
        selected = app.filter_dashboard_records(rows, filters={"showroom": "معرض"})
        self.assertEqual(len(selected), 135)
        from collections import Counter
        self.assertEqual(len(app.OperationsDashboardDialog._counter(selected, "marble_color")), 135)
        self.assertEqual(sum(app.OperationsDashboardDialog._counter(selected, "marble_color").values()), 135)


class InteractionAcceptance(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("Tk unavailable")
        self.root.withdraw()
        self.gui = app.MarbleDeliveryApp(self.root)
        self.gui.show_all_records()

    def tearDown(self):
        if hasattr(self, "root"):
            try:
                self.gui._cancel_record_autosave()
                if getattr(self.gui, "_session_after_id", None):
                    self.root.after_cancel(self.gui._session_after_id)
                self.root.destroy()
            except Exception:
                pass

    def test_enter_moves_focus_without_saving_or_clearing(self):
        entry = self.gui.entries["customer_name"]
        entry.delete(0, "end")
        entry.insert(0, "اختبار انتقال")
        target = []
        with patch.object(self.gui, "_focus_entry_index", side_effect=target.append):
            pos = self.gui._entry_order.index("customer_name")
            self.assertEqual(self.gui._handle_enter(None, pos), "break")
        self.assertEqual(target, [pos + 1])
        self.assertEqual(entry.get(), "اختبار انتقال")

    def test_no_blank_autosave_record(self):
        start_count = len(self.gui.records)
        self.gui.clear_form()
        self.assertFalse(self.gui._autosave_current_record())
        self.assertEqual(len(self.gui.records), start_count)

    def test_create_edit_same_record_without_phone(self):
        self.gui.clear_form()
        entry = self.gui.entries["customer_name"]
        entry.insert(0, "العميل قبل التعديل")
        self.assertTrue(self.gui.save_record())
        record_id = self.gui.selected_record_id
        self.assertIsNotNone(record_id)
        self.assertEqual(entry.get(), "العميل قبل التعديل")
        entry.delete(0, "end")
        entry.insert(0, "العميل بعد التعديل")
        self.assertTrue(self.gui.update_record())
        updated = self.gui._record_by_id(record_id)
        self.assertEqual(updated["values"]["customer_name"], "العميل بعد التعديل")
        self.assertEqual(updated["values"]["mobile"], "")
        self.assertEqual(sum(r["id"] == record_id for r in self.gui.records), 1)

    def test_dashboard_visual_filters(self):
        records_backup = self.gui.records
        self.gui.records = [
            row("vis-one", color="أبيض", showroom="الشرق", employee="سعيد"),
            row("vis-two", color="أسود", showroom="الشرق", employee="خالد"),
            row("vis-three", color="أبيض", showroom="الغرب", employee="سعيد"),
        ]
        try:
            dash = app.OperationsDashboardDialog(self.gui)
            dash.period_var.set("كل الفترات")
            dash.dashboard_filter_vars["marble_color"].set("أبيض")
            dash.dashboard_filter_vars["salesperson"].set("سعيد")
            dash._refresh()
            self.assertEqual(len(dash.filtered_records), 2)
            self.assertGreaterEqual(len(dash.notebook.tabs()), 5)
            dash.extra_field_var.set("المعرض")
            dash.extra_value_var.set("الغرب")
            dash._refresh()
            self.assertEqual([r["id"] for r in dash.filtered_records], ["vis-three"])
            dash.destroy()
        finally:
            self.gui.records = records_backup

    def test_responsive_form_keeps_entries(self):
        before = {key: id(widget) for key, widget in self.gui.entries.items()}
        self.gui._apply_compact_layout()
        self.assertEqual({key: id(widget) for key, widget in self.gui.entries.items()}, before)
        self.assertGreaterEqual(self.gui._form_field_columns, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
