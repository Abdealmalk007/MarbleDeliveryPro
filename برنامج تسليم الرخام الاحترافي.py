# -*- coding: utf-8 -*-
"""
نظام تسليم الرخام الاحترافي - الإصدار 10.0.4
إصدار مؤسسي شامل: واجهة عربية RTL مرنة، مساحة تقرير قابلة للتركيز، سلة محذوفات، أدلة عملاء ومعارض،
قوالب تقارير، مستخدمون وصلاحيات، واستيراد/تصدير وطباعة احترافية عالية الدقة PDF/PNG/JPG/Excel.
"""
from __future__ import annotations

import copy
import calendar
import csv
import datetime as dt
import html as html_lib
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
import webbrowser
import difflib
import hashlib
import io
import zipfile
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from contextlib import contextmanager
from dataclasses import asdict, dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any, Callable, Optional, List, Tuple
from xml.sax.saxutils import escape

import tkinter as tk
from tkinter import colorchooser, filedialog, font as tkfont, messagebox, ttk

# ============================================================================
# المكتبات الخارجية (مع التحقق من وجودها)
# ============================================================================
REPORTLAB_IMPORT_ERROR = ""
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.drawing.image import Image as ExcelImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, colors as openpyxl_colors
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.pagebreak import Break
except ImportError:
    Workbook = load_workbook = ExcelImage = BarChart = Reference = DataLabelList = Break = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A3, A4, LETTER, landscape, portrait
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.graphics import renderPDF
    from reportlab.graphics.barcode.qr import QrCodeWidget
    from reportlab.graphics.shapes import Drawing
    from reportlab.pdfgen.canvas import Canvas as PDFCanvas
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
except ImportError as exc:
    REPORTLAB_IMPORT_ERROR = str(exc)
    colors = None

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError:
    Image = ImageDraw = ImageFont = ImageTk = None

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except ImportError:
    arabic_reshaper = None
    get_display = None

# ============================================================================
# قاعدة بيانات SQLite
# ============================================================================
import sqlite3
import hmac

def _user_data_dir() -> Path:
    override = os.environ.get("MARBLE_DELIVERY_DATA_DIR", "").strip()
    if override:
        path = Path(override).expanduser()
    elif sys.platform.startswith("win"):
        path = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "MarbleDeliveryPro"
    elif sys.platform.startswith("darwin"):
        path = Path.home() / "Library" / "Application Support" / "MarbleDeliveryPro"
    else:
        path = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share")) / "marble-delivery-pro"
    path.mkdir(parents=True, exist_ok=True)
    return path


APP_DATA_DIR = _user_data_dir()
BACKUP_DIR = APP_DATA_DIR / "backups"
PREVIEW_DIR = APP_DATA_DIR / "previews"
ATTACHMENTS_DIR = APP_DATA_DIR / "attachments"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
DB_FILE = APP_DATA_DIR / "marble_delivery_v7.db"
LOG_FILE = APP_DATA_DIR / "marble_delivery.log"
LEGACY_DB_FILES = [
    APP_DATA_DIR / "marble_delivery_v6.db",
    Path(__file__).resolve().parent / "marble_delivery_v6.db",
    Path(__file__).resolve().parent / "marble_delivery_v5.db",
]

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)

if not DB_FILE.exists():
    for legacy_db in LEGACY_DB_FILES:
        if legacy_db.exists() and legacy_db.resolve() != DB_FILE.resolve():
            try:
                shutil.copy2(legacy_db, DB_FILE)
                logging.info("تم ترحيل قاعدة البيانات القديمة من %s", legacy_db)
                break
            except OSError:
                logging.exception("تعذر ترحيل قاعدة البيانات القديمة من %s", legacy_db)

@contextmanager
def get_db():
    """اتصال قصير العمر مع تأكيد/تراجع المعاملة وإغلاق المقبض دائمًا."""
    conn = sqlite3.connect(str(DB_FILE), detect_types=sqlite3.PARSE_DECLTYPES, timeout=20)
    try:
        conn.text_factory = str
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA encoding = 'UTF-8'")
        conn.execute("PRAGMA busy_timeout = 20000")
        conn.execute("PRAGMA journal_mode = WAL")
        with conn:
            yield conn
    finally:
        conn.close()

def init_db():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id TEXT PRIMARY KEY,
                created_at TEXT,
                updated_at TEXT,
                material_estimate TEXT,
                contract_number TEXT,
                customer_name TEXT,
                mobile TEXT,
                marble_color TEXT,
                sink_type TEXT,
                marble_edge TEXT,
                showroom TEXT,
                salesperson TEXT,
                custom_fields TEXT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                actor TEXT NOT NULL,
                action TEXT NOT NULL,
                record_id TEXT,
                details TEXT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS attachments (
                id TEXT PRIMARY KEY,
                record_id TEXT NOT NULL,
                original_name TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                added_at TEXT NOT NULL,
                FOREIGN KEY(record_id) REFERENCES records(id) ON DELETE CASCADE
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS session_state (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                payload TEXT NOT NULL DEFAULT '{}',
                updated_at TEXT NOT NULL,
                clean_shutdown INTEGER NOT NULL DEFAULT 1
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS report_archive (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                report_format TEXT NOT NULL,
                file_name TEXT NOT NULL,
                report_title TEXT NOT NULL,
                company_name TEXT NOT NULL,
                recipient TEXT,
                record_count INTEGER NOT NULL DEFAULT 0,
                selected_fields TEXT NOT NULL DEFAULT '[]',
                search_text TEXT NOT NULL DEFAULT '',
                sha256 TEXT NOT NULL,
                file_data BLOB NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS integration_outbox (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                action TEXT NOT NULL,
                record_id TEXT,
                payload TEXT NOT NULL DEFAULT '{}',
                status TEXT NOT NULL DEFAULT 'pending',
                attempts INTEGER NOT NULL DEFAULT 0,
                last_error TEXT NOT NULL DEFAULT '',
                sent_at TEXT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS deleted_records (
                id TEXT PRIMARY KEY,
                deleted_at TEXT NOT NULL,
                payload TEXT NOT NULL,
                attachment_manifest TEXT NOT NULL DEFAULT '[]'
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS directories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_type TEXT NOT NULL,
                name TEXT NOT NULL,
                mobile TEXT NOT NULL DEFAULT '',
                details TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(entity_type, name)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS report_templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                payload TEXT NOT NULL DEFAULT '{}'
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id TEXT PRIMARY KEY,
                username TEXT NOT NULL UNIQUE,
                display_name TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'موظف',
                permissions TEXT NOT NULL DEFAULT '[]',
                salt TEXT NOT NULL DEFAULT '',
                password_hash TEXT NOT NULL DEFAULT '',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        now_text = dt.datetime.now().isoformat(timespec="seconds")
        cursor.execute(
            """INSERT OR IGNORE INTO users
               (id, username, display_name, role, permissions, salt, password_hash, active, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, '', '', 1, ?, ?)""",
            ("builtin-admin", "admin", "مدير النظام", "مدير", json.dumps(["*"], ensure_ascii=False), now_text, now_text),
        )
        for col in ["contract_number", "mobile", "customer_name", "showroom", "salesperson", "marble_color", "sink_type", "created_at", "updated_at"]:
            cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_{col} ON records({col})")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_created_at ON audit_log(created_at)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_record ON audit_log(record_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_attachment_record ON attachments(record_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_report_archive_created ON report_archive(created_at)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_report_archive_format ON report_archive(report_format)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_integration_status_created ON integration_outbox(status, created_at)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_deleted_records_date ON deleted_records(deleted_at)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_directories_type_name ON directories(entity_type, name)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
        conn.commit()

DB_HEALTH_ERROR = ""
try:
    init_db()
    with get_db() as _health_conn:
        _health_result = _health_conn.execute("PRAGMA quick_check").fetchone()
        if not _health_result or str(_health_result[0]).lower() != "ok":
            DB_HEALTH_ERROR = str(_health_result[0] if _health_result else "فشل فحص قاعدة البيانات")
except Exception as _db_health_exc:
    DB_HEALTH_ERROR = str(_db_health_exc)
    logging.exception("فشل فحص قاعدة البيانات عند بدء التشغيل")

# ============================================================================
# إعدادات التطبيق
# ============================================================================
@dataclass
class FieldDefinition:
    key: str
    label: str
    visible: bool = True
    required: bool = False
    width: int = 145
    kind: str = "text"
    system: bool = True
    options: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "FieldDefinition":
        return cls(
            key=str(raw.get("key") or f"custom_{uuid.uuid4().hex[:10]}"),
            label=str(raw.get("label") or "حقل جديد"),
            visible=bool(raw.get("visible", True)),
            required=bool(raw.get("required", False)),
            width=max(80, min(320, int(raw.get("width", 145)))),
            kind=str(raw.get("kind", "text")),
            system=bool(raw.get("system", False)),
            options=[clean_value(item) for item in raw.get("options", []) if clean_value(item)]
            if isinstance(raw.get("options"), list) else [],
        )

def default_fields() -> list[FieldDefinition]:
    return [
        FieldDefinition("material_estimate", "تقدير الخامات", width=150),
        FieldDefinition("contract_number", "رقم العقد", required=True, width=120),
        FieldDefinition("kitchen_number", "رقم/اسم المطبخ (اختياري)", visible=False, width=135),
        FieldDefinition("customer_name", "اسم العميل", required=True, width=190),
        FieldDefinition("mobile", "رقم الجوال", required=True, width=145, kind="phone"),
        FieldDefinition("marble_color", "لون الرخام", width=145),
        FieldDefinition("sink_type", "نوع المجلى", width=145),
        FieldDefinition("marble_edge", "شفة الرخام", width=145),
        FieldDefinition("showroom", "المعرض", width=145),
        FieldDefinition("salesperson", "موظف المعرض", width=175),
        FieldDefinition("delivery_status", "حالة التسليم", required=True, width=160, kind="choice",
                        options=["جديد", "يحتاج رفع مقاس", "نهائي لا يحتاج رفع مقاس",
                                 "جاهز للتسليم", "تم التسليم",
                                 "تعديل - يعتمد الحالي ويلغى السابق", "مؤجل", "ملغي"]),
        FieldDefinition("scheduled_delivery_date", "موعد التسليم", width=135, kind="date"),
        FieldDefinition("actual_delivery_date", "تاريخ التركيب الفعلي", width=155, kind="date"),
    ]

@dataclass
class AppSettings:
    company_name: str = ""
    report_title: str = "تقرير تسليم الرخام"
    report_title_en: str = "Marble Delivery Report"
    report_subtitle: str = "إدارة التسليم والمتابعة"
    footer_note: str = "نشكركم لاختياركم منتجاتنا"
    logo_path: str = ""
    font_family: str = "Segoe UI"
    base_font_size: int = 12
    heading_font_size: int = 22
    primary_color: str = "#17324D"
    accent_color: str = "#C9A84C"
    background_color: str = "#EEF2F5"
    surface_color: str = "#FFFFFF"
    text_color: str = "#1A1A2E"
    muted_color: str = "#6C757D"
    success_color: str = "#2B8C5E"
    danger_color: str = "#C13E3E"
    print_friendly_reports: bool = True
    fields: list[FieldDefinition] = field(default_factory=default_fields)
    export_preferences: dict = field(default_factory=dict)
    hidden_table_columns: list[str] = field(default_factory=list)
    compact_kitchen_display_migrated: bool = True
    sync_folder: str = ""
    central_api_url: str = ""
    erp_api_url: str = ""
    notification_webhook_url: str = ""
    integration_api_token: str = ""
    integration_timeout_seconds: int = 20
    auto_contract_prefix: str = "MD"
    overdue_warning: bool = True
    logo_width_mm: int = 35
    png_scale_percent: int = 100  # توافق مع الإصدارات السابقة
    report_page_size: str = "A4"
    report_orientation: str = "auto"
    report_margin_mm: float = 6.0
    pdf_base_font_size: float = 10.0
    pdf_min_font_size: float = 7.5
    png_dpi: int = 300
    jpg_quality: int = 95
    custom_page_width_mm: float = 210.0
    custom_page_height_mm: float = 297.0
    default_export_folder: str = ""
    backup_keep_count: int = 20
    autosave_seconds: int = 4
    require_login: bool = False
    compact_form_default: bool = False

    # ألوان العناوين الجديدة (ذهبي على أخضر فاتح)
    title_bg_color: str = "#C8E6C9"
    title_fg_color: str = "#C9A84C"
    subtitle_bg_color: str = "#C8E6C9"
    subtitle_fg_color: str = "#C9A84C"

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "AppSettings":
        defaults = cls()
        allowed = {
            key for key in asdict(defaults).keys()
            if key not in ("fields", "export_preferences", "hidden_table_columns")
        }
        values = {key: raw[key] for key in allowed if key in raw}
        values["base_font_size"] = max(10, min(22, int(values.get("base_font_size", 12))))
        values["heading_font_size"] = max(16, min(34, int(values.get("heading_font_size", 22))))
        values["logo_width_mm"] = max(15, min(60, int(values.get("logo_width_mm", 35))))
        values["png_scale_percent"] = max(75, min(200, int(values.get("png_scale_percent", 100))))
        values["report_page_size"] = str(values.get("report_page_size", "A4")).upper()
        if values["report_page_size"] not in {"A4", "A3", "LETTER", "CUSTOM"}:
            values["report_page_size"] = "A4"
        values["report_orientation"] = str(values.get("report_orientation", "auto")).lower()
        if values["report_orientation"] not in {"auto", "portrait", "landscape"}:
            values["report_orientation"] = "auto"
        values["report_margin_mm"] = max(3.0, min(20.0, float(values.get("report_margin_mm", 6.0))))
        values["pdf_base_font_size"] = max(8.0, min(14.0, float(values.get("pdf_base_font_size", 10.0))))
        values["pdf_min_font_size"] = max(6.5, min(values["pdf_base_font_size"], float(values.get("pdf_min_font_size", 7.5))))
        values["png_dpi"] = max(150, min(600, int(values.get("png_dpi", 300))))
        values["jpg_quality"] = max(70, min(100, int(values.get("jpg_quality", 95))))
        values["custom_page_width_mm"] = max(80.0, min(1000.0, float(values.get("custom_page_width_mm", 210.0))))
        values["custom_page_height_mm"] = max(80.0, min(1000.0, float(values.get("custom_page_height_mm", 297.0))))
        values["backup_keep_count"] = max(5, min(100, int(values.get("backup_keep_count", 20))))
        values["autosave_seconds"] = max(2, min(60, int(values.get("autosave_seconds", 4))))
        values["integration_timeout_seconds"] = max(
            5, min(120, int(values.get("integration_timeout_seconds", 20)))
        )
        raw_fields = raw.get("fields")
        values["fields"] = (
            [FieldDefinition.from_dict(item) for item in raw_fields if isinstance(item, dict)]
            if isinstance(raw_fields, list) and raw_fields
            else default_fields()
        )
        existing_keys = {item.key for item in values["fields"]}
        for default_field in default_fields():
            if default_field.key not in existing_keys:
                if default_field.key == "kitchen_number":
                    contract_index = next(
                        (index for index, item in enumerate(values["fields"]) if item.key == "contract_number"),
                        len(values["fields"]) - 1,
                    )
                    values["fields"].insert(contract_index + 1, default_field)
                else:
                    values["fields"].append(default_field)
                existing_keys.add(default_field.key)
        # ترقية الحقول النظامية مع الحفاظ على الحقول المخصصة وإعدادات المستخدم.
        default_by_key = {item.key: item for item in default_fields()}
        for current in values["fields"]:
            default_field = default_by_key.get(current.key)
            if current.key == "delivery_status" and default_field:
                current.label = default_field.label
                current.kind = "choice"
                current.options = list(dict.fromkeys([*default_field.options, *current.options]))
                current.width = max(current.width, default_field.width)
            elif current.key in {"scheduled_delivery_date", "actual_delivery_date"} and default_field:
                current.label = default_field.label
                current.kind = "date"
                current.width = max(current.width, default_field.width)
        if not any(item.visible for item in values["fields"]):
            values["fields"][0].visible = True
        values["export_preferences"] = raw.get("export_preferences", {})
        valid_field_keys = {item.key for item in values["fields"]}
        hidden_table_columns = raw.get("hidden_table_columns", defaults.hidden_table_columns)
        values["hidden_table_columns"] = (
            [
                str(key) for key in hidden_table_columns
                if isinstance(key, str) and key in valid_field_keys
            ]
            if isinstance(hidden_table_columns, list) else []
        )
        # ترقية سابقة: أُخفي عمود رقم/اسم المطبخ مرة واحدة من الجدول لتقليل الزحام.
        # يبقى الحقل وبياناته كاملين، ويمكن للمستخدم إظهاره مجددًا من «اختيار الأعمدة».
        if not bool(raw.get("compact_kitchen_display_migrated", False)):
            values["hidden_table_columns"] = list(dict.fromkeys([
                *values["hidden_table_columns"], "kitchen_number",
            ]))
            for current in values["fields"]:
                if current.key == "kitchen_number":
                    current.visible = False
            values["compact_kitchen_display_migrated"] = True
        for key in ["title_bg_color", "title_fg_color", "subtitle_bg_color", "subtitle_fg_color"]:
            if key in raw:
                values[key] = raw[key]
        return cls(**values)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @property
    def visible_fields(self) -> list[FieldDefinition]:
        return [f for f in self.fields if f.visible] or self.fields[:1]

    @property
    def table_fields(self) -> list[FieldDefinition]:
        hidden = set(self.hidden_table_columns)
        fields = [f for f in self.visible_fields if f.key not in hidden]
        return fields or self.visible_fields[:1]

    @property
    def report_fields(self) -> list[FieldDefinition]:
        """حقول التقرير الافتراضية. موعد التسليم محفوظ في البيانات لكنه غير قابل للتصدير."""
        excluded = {"kitchen_number", "scheduled_delivery_date"}
        fields = [f for f in self.visible_fields if f.key not in excluded]
        return fields or [f for f in self.visible_fields if f.key != "scheduled_delivery_date"][:1]

# ============================================================================
# دوال مساعدة
# ============================================================================
APP_NAME = "نظام تسليم الرخام الاحترافي"
APP_VERSION = "10.0.5"
SCHEMA_VERSION = 14
STATUS_OPTIONS = (
    "جديد", "يحتاج رفع مقاس", "نهائي لا يحتاج رفع مقاس",
    "جاهز للتسليم", "تم التسليم", "تعديل - يعتمد الحالي ويلغى السابق", "مؤجل", "ملغي",
)
FINISHED_STATUSES = {"تم التسليم", "ملغي"}
REPORT_STATUS_FILLS = {
    "جاهز للتسليم": "DCEEFF",
    "تم التسليم": "DDF1E5",
    "تعديل - يعتمد الحالي ويلغى السابق": "FFE8A3",
    "مؤجل": "FFF1CC",
    "ملغي": "FDE3E3",
}
REPORT_MAX_FONT_SIZE = 14
REPORT_MIN_FONT_SIZE = 7.0
REPORT_FONT_SIZE = REPORT_MAX_FONT_SIZE  # توافق مع التخطيط القديم غير المستخدم
EXPORT_METHOD_NAMES = {
    "xlsx": "export_excel",
    "pdf": "export_pdf",
    "png": "export_png",
    "jpg": "export_jpg",
    "csv": "export_csv",
    "html": "export_html",
}



def report_table_sort_value(value: Any) -> tuple[int, Any]:
    """مفتاح فرز متجانس: الأرقام والتواريخ أولاً، ثم النصوص العربية والإنجليزية."""
    text = clean_value(value)
    if not text:
        return (3, "")
    digits = text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789"))
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", digits):
        return (0, digits)
    numeric = digits.replace(",", "").strip()
    try:
        if re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", numeric):
            return (1, float(numeric))
    except ValueError:
        pass
    return (2, normalize_search_text(text))


def balanced_report_field_groups(fields: list[Any], max_columns: int = 3) -> list[list[Any]]:
    """يقسم الحقول بالتساوي دون ترك مجموعة أخيرة ضيقة قدر الإمكان."""
    items = list(fields)
    if not items:
        return []
    group_count = max(1, math.ceil(len(items) / max(1, max_columns)))
    base_size, remainder = divmod(len(items), group_count)
    groups = []
    start = 0
    for index in range(group_count):
        size = base_size + (1 if index < remainder else 0)
        groups.append(items[start:start + size])
        start += size
    return groups


def balanced_report_page_chunks(
    records: list[Any], target_rows: int = 13, minimum_rows: int = 10,
) -> list[list[Any]]:
    """يقسم السجلات رأسياً مع إبقاء عشر سجلات على الأقل في الصفحة متى أمكن."""
    items = list(records)
    if not items:
        return [[]]
    minimum_rows = max(1, int(minimum_rows))
    target_rows = max(minimum_rows, int(target_rows))
    page_count = max(1, math.ceil(len(items) / target_rows))
    while page_count > 1 and len(items) // page_count < minimum_rows:
        page_count -= 1
    base_size, remainder = divmod(len(items), page_count)
    chunks: list[list[Any]] = []
    cursor = 0
    for page_index in range(page_count):
        chunk_size = base_size + (1 if page_index < remainder else 0)
        chunks.append(items[cursor:cursor + chunk_size])
        cursor += chunk_size
    return chunks


def report_rows_per_page(fields: list[Any]) -> int:
    """تقدير للطباعة عبر HTML فقط؛ PDF/PNG يحددان السعة الفعلية للورقة ديناميكياً."""
    count = len(fields)
    if count <= 4:
        return 28
    if count <= 8:
        return 22
    return 18


def largest_fitting_page_chunk(
    records: list[Any], fit_chunk: Callable[[list[Any]], Any],
) -> tuple[list[Any], Any]:
    """يعيد أكبر عدد متتالٍ من السجلات يمكن أن يتسع فعلياً في الصفحة.

    البحث ثنائي لأن قابلية الملاءمة أحادية الاتجاه: إذا لم تتسع N سجلات فلن
    تتسع N+1. بهذه الطريقة لا يوجد حد ثابت مثل 7 صفوف، وتُملأ كل صفحة
    حتى أقصى سعة ممكنة قبل الانتقال تلقائياً إلى الصفحة التالية.
    """
    items = list(records)
    if not items:
        fitted = fit_chunk([])
        return [], fitted

    low, high = 1, len(items)
    best_count = 0
    best_fit = None
    while low <= high:
        middle = (low + high) // 2
        fitted = fit_chunk(items[:middle])
        if fitted is not None:
            best_count = middle
            best_fit = fitted
            low = middle + 1
        else:
            high = middle - 1

    if best_count <= 0:
        return [], None
    return items[:best_count], best_fit


def report_column_weights(fields: list[Any], records: list[dict[str, Any]]) -> list[float]:
    """يوزع عرض A4 حسب المحتوى المعتاد مع أولوية واضحة لاسم العميل."""
    layout_hints: dict[str, tuple[float, float, float]] = {
        "customer_name": (21.0, 31.0, 1.30),
        "salesperson": (12.0, 18.0, 1.02),
        "contract_number": (13.0, 18.0, 1.08),
        "kitchen_number": (10.0, 14.0, 0.95),
        "mobile": (10.5, 14.0, 0.88),
        "scheduled_delivery_date": (11.0, 15.0, 0.92),
        "actual_delivery_date": (12.0, 16.5, 0.94),
        "delivery_status": (14.0, 19.0, 1.08),
        "material_estimate": (8.0, 12.0, 0.82),
        "marble_color": (9.0, 14.0, 0.92),
        "sink_type": (8.0, 12.0, 0.84),
        "marble_edge": (7.5, 11.0, 0.80),
        "showroom": (8.0, 12.5, 0.84),
    }
    weights: list[float] = []
    for item in fields:
        label = clean_value(getattr(item, "label", ""))
        key = clean_value(getattr(item, "key", ""))
        values = [
            clean_value(record.get("values", {}).get(key, ""))
            for record in records[:200]
        ]
        nonempty_lengths = sorted(len(value) for value in values if value)
        if nonempty_lengths:
            percentile_index = min(
                len(nonempty_lengths) - 1,
                max(0, math.ceil(len(nonempty_lengths) * 0.85) - 1),
            )
            typical_length = nonempty_lengths[percentile_index]
        else:
            typical_length = 0
        label_weight = len(label) * 0.52 + 2.5
        content_weight = typical_length * 0.40 + 3.0
        configured_width = float(getattr(item, "width", 145) or 145)
        configured_weight = max(6.0, min(20.0, configured_width / 14.0))
        weight = max(label_weight, content_weight, configured_weight * 0.72)
        minimum, maximum, multiplier = layout_hints.get(key, (7.0, 20.0, 1.0))
        weight = max(minimum, min(maximum, weight * multiplier))
        populated_ratio = sum(bool(value) for value in values) / max(1, len(values))
        if key != "customer_name" and values and populated_ratio < 0.30:
            weight = max(minimum, weight * 0.88)
        weights.append(weight)
    return weights


def record_date_for_basis(record: dict[str, Any], basis: str = "التاريخ الأنسب") -> Optional[dt.date]:
    """يعيد تاريخ السجل وفق اختيار المستخدم، مع مسار تلقائي متوافق مع الإصدارات السابقة."""
    values = record.get("values", {})
    candidates_by_basis = {
        "موعد التسليم": [values.get("scheduled_delivery_date", "")],
        "التركيب الفعلي": [values.get("actual_delivery_date", "")],
        "تاريخ إنشاء السجل": [record.get("created_at", "")],
        "تاريخ آخر تعديل": [record.get("updated_at", "")],
    }
    candidates = candidates_by_basis.get(basis, [
        values.get("actual_delivery_date", ""),
        values.get("scheduled_delivery_date", ""),
        record.get("created_at", ""),
    ])
    for candidate in candidates:
        text = clean_value(candidate)
        if not text:
            continue
        try:
            return dt.datetime.strptime(text[:10], "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
    return None

def clean_value(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def report_record_metrics(records: list[dict[str, Any]]) -> dict[str, int]:
    """مؤشرات واضحة تفصل بين عدد العقود وعدد المطابخ/السجلات."""
    contracts = [
        clean_value(record.get("values", {}).get("contract_number", ""))
        for record in records
    ]
    contracts = [item for item in contracts if item]
    counts = Counter(contracts)
    return {
        "records": len(records),
        "contracts": len(counts),
        "multi_kitchen_contracts": sum(1 for count in counts.values() if count > 1),
        "extra_kitchens": sum(max(0, count - 1) for count in counts.values()),
    }


def next_kitchen_number(
    records: list[dict[str, Any]], contract_number: Any, *, exclude_id: str | None = None,
) -> str:
    """يرجع أول رقم مطبخ متاح داخل العقد؛ والسجل القديم بلا رقم يُعامل كمطبخ 1."""
    contract = clean_value(contract_number)
    if not contract:
        return "1"
    used: set[int] = set()
    for record in records:
        if exclude_id and str(record.get("id", "")) == str(exclude_id):
            continue
        values = record.get("values", {}) if isinstance(record, dict) else {}
        if clean_value(values.get("contract_number", "")) != contract:
            continue
        kitchen = clean_value(values.get("kitchen_number", ""))
        if not kitchen:
            used.add(1)
            continue
        match = re.search(r"\d+", kitchen)
        if match:
            used.add(max(1, int(match.group(0))))
        else:
            used.add(1)
    candidate = 1
    while candidate in used:
        candidate += 1
    return str(candidate)


def ensure_kitchen_identity(
    values: dict[str, Any], records: list[dict[str, Any]], *, exclude_id: str | None = None,
) -> str:
    """يملأ رقم المطبخ عند تكرار العقد ويحل تضارب الرقم من دون تغيير رقم العقد."""
    contract = clean_value(values.get("contract_number", ""))
    if not contract:
        return clean_value(values.get("kitchen_number", ""))
    matching = [
        record for record in records
        if str(record.get("id", "")) != str(exclude_id or "")
        and clean_value(record.get("values", {}).get("contract_number", "")) == contract
    ]
    kitchen = clean_value(values.get("kitchen_number", ""))
    if not matching:
        if not kitchen:
            values["kitchen_number"] = "1"
        return clean_value(values.get("kitchen_number", ""))
    occupied = {
        clean_value(record.get("values", {}).get("kitchen_number", "")) or "1"
        for record in matching
    }
    if not kitchen or kitchen in occupied:
        kitchen = next_kitchen_number(records, contract, exclude_id=exclude_id)
        values["kitchen_number"] = kitchen
    return kitchen


def normalize_search_text(value: Any) -> str:
    text = clean_value(value).casefold()
    text = re.sub(r"[\u064b-\u065f\u0670\u0640]", "", text)
    text = text.translate(str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه"}))
    return clean_value(text)


def arabic_date_text(value: Optional[dt.date] = None) -> str:
    day_value = value or dt.date.today()
    weekdays = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
    months = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
              "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
    return f"{weekdays[day_value.weekday()]}، {day_value.day} {months[day_value.month - 1]} {day_value.year}"


def record_is_overdue(record: dict[str, Any], today: Optional[dt.date] = None) -> bool:
    values = record.get("values", {}) if isinstance(record, dict) else {}
    status = clean_value(values.get("delivery_status", "جديد")) or "جديد"
    due_text = clean_value(values.get("scheduled_delivery_date", ""))
    if not due_text or status in FINISHED_STATUSES:
        return False
    try:
        due_date = dt.datetime.strptime(due_text, "%Y-%m-%d").date()
    except ValueError:
        return False
    return due_date < (today or dt.date.today())


def integration_snapshot_payload(
    records: list[dict[str, Any]], settings: AppSettings, *, package_type: str,
    events: Optional[list[dict[str, Any]]] = None,
) -> dict[str, Any]:
    """حزمة تبادل مستقرة يمكن إرسالها إلى API مركزي أو ERP."""
    return {
        "package_type": clean_value(package_type) or "records_snapshot",
        "app_name": APP_NAME,
        "app_version": APP_VERSION,
        "schema_version": SCHEMA_VERSION,
        "generated_at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "company": clean_value(settings.company_name),
        "fields": [asdict(item) for item in settings.fields],
        "record_count": len(records),
        "records": copy.deepcopy(list(records)),
        "events": copy.deepcopy(list(events or [])),
    }


def notification_summary_payload(records: list[dict[str, Any]], settings: AppSettings) -> dict[str, Any]:
    today = dt.date.today()
    overdue = [record for record in records if record_is_overdue(record, today)]
    due_today = []
    delivered_today = []
    for record in records:
        values = record.get("values", {}) if isinstance(record, dict) else {}
        if clean_value(values.get("scheduled_delivery_date")) == today.isoformat():
            due_today.append(record)
        if clean_value(values.get("actual_delivery_date")) == today.isoformat():
            delivered_today.append(record)
    return {
        "package_type": "delivery_notification_summary",
        "app_version": APP_VERSION,
        "generated_at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "company": clean_value(settings.company_name),
        "date": today.isoformat(),
        "counts": {
            "total_records": len(records),
            "overdue": len(overdue),
            "due_today": len(due_today),
            "installed_today": len(delivered_today),
        },
        "overdue_records": overdue,
        "due_today_records": due_today,
        "installed_today_records": delivered_today,
    }


def post_json_payload(url: str, payload: dict[str, Any], token: str = "", timeout: int = 20) -> dict[str, Any]:
    endpoint = clean_value(url)
    parsed = urllib.parse.urlparse(endpoint)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("أدخل رابط HTTP أو HTTPS صحيحًا لخدمة التكامل.")
    headers = {
        "Content-Type": "application/json; charset=utf-8",
        "Accept": "application/json",
        "User-Agent": f"MarbleDeliveryPro/{APP_VERSION}",
    }
    if clean_value(token):
        headers["Authorization"] = f"Bearer {clean_value(token)}"
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=max(5, min(120, int(timeout)))) as response:
            raw = response.read(1_000_000)
            status = int(getattr(response, "status", 200) or 200)
    except urllib.error.HTTPError as exc:
        detail = exc.read(2000).decode("utf-8", errors="replace")
        raise RuntimeError(f"رفضت الخدمة الطلب برمز {exc.code}: {clean_value(detail)[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"تعذر الاتصال بخدمة التكامل: {clean_value(exc.reason)}") from exc
    if not 200 <= status < 300:
        raise RuntimeError(f"أعادت خدمة التكامل رمزًا غير ناجح: {status}")
    if not raw:
        return {"ok": True, "status": status}
    try:
        decoded = json.loads(raw.decode("utf-8"))
        return decoded if isinstance(decoded, dict) else {"ok": True, "status": status, "response": decoded}
    except (UnicodeDecodeError, json.JSONDecodeError):
        return {"ok": True, "status": status, "response_text": raw.decode("utf-8", errors="replace")[:1000]}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_backup_archive(path: Path) -> dict[str, Any]:
    """يفحص النسخة دون استعادتها أو تعديل قاعدة البيانات الحالية."""
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError("ملف النسخة الاحتياطية غير موجود.")
    if source.suffix.lower() == ".json":
        data = json.loads(source.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("records"), list):
            raise ValueError("ملف JSON لا يحتوي بنية نسخة احتياطية صحيحة.")
        return {
            "valid": True,
            "format": "json",
            "records": len(data.get("records", [])),
            "attachments": 0,
            "reports": 0,
            "manifest_files": 0,
            "legacy_without_manifest": True,
        }
    with zipfile.ZipFile(source, "r") as archive:
        bad_file = archive.testzip()
        if bad_file:
            raise ValueError(f"يوجد ملف تالف داخل النسخة: {bad_file}")
        names = set(archive.namelist())
        for name in names:
            pure = PurePosixPath(name)
            if pure.is_absolute() or ".." in pure.parts:
                raise ValueError(f"مسار غير آمن داخل النسخة: {name}")
        if "backup.json" not in names:
            raise ValueError("النسخة لا تحتوي ملف backup.json.")
        data = json.loads(archive.read("backup.json").decode("utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("records"), list):
            raise ValueError("بنية بيانات النسخة الاحتياطية غير صحيحة.")
        manifest = data.get("file_manifest", {})
        if manifest and not isinstance(manifest, dict):
            raise ValueError("دليل التحقق داخل النسخة غير صحيح.")
        for name, expected in manifest.items() if isinstance(manifest, dict) else []:
            if name not in names:
                raise ValueError(f"ملف مفقود من النسخة: {name}")
            digest = hashlib.sha256()
            size = 0
            with archive.open(name, "r") as stream:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(chunk)
                    size += len(chunk)
            if not isinstance(expected, dict):
                raise ValueError(f"بيانات تحقق غير صحيحة للملف: {name}")
            if digest.hexdigest().lower() != clean_value(expected.get("sha256")).lower():
                raise ValueError(f"بصمة الملف لا تتطابق: {name}")
            if int(expected.get("size", -1)) != size:
                raise ValueError(f"حجم الملف لا يتطابق: {name}")
        return {
            "valid": True,
            "format": "zip",
            "records": len(data.get("records", [])),
            "attachments": len(data.get("attachments", [])),
            "reports": len(data.get("reports", [])),
            "manifest_files": len(manifest) if isinstance(manifest, dict) else 0,
            "legacy_without_manifest": not bool(manifest),
        }


def safe_filename(value: str) -> str:
    name = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", Path(value).name).strip(" .")
    return name[:160] or "attachment"


def report_reference(records: list[dict[str, Any]], generated_at: dt.datetime) -> str:
    seed = "|".join([
        generated_at.isoformat(timespec="seconds"),
        str(len(records)),
        *[str(item.get("id", "")) for item in records[:50]],
    ])
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()[:12].upper()


def html_escape(value: Any) -> str:
    return html_lib.escape(clean_value(value), quote=True)

def is_hex_color(value: str) -> bool:
    return bool(re.fullmatch(r"#[0-9A-Fa-f]{6}", value.strip()))


def open_with_default_app(path: Path) -> None:
    resolved = path.resolve()
    if sys.platform.startswith("win"):
        os.startfile(str(resolved))
    elif sys.platform.startswith("darwin"):
        subprocess.Popen(["open", str(resolved)])
    else:
        opener = shutil.which("xdg-open")
        if not opener:
            raise RuntimeError("لم يتم العثور على أداة فتح الملفات الافتراضية.")
        subprocess.Popen([opener, str(resolved)])


def print_pdf_file(path: Path) -> None:
    resolved = path.resolve()
    if sys.platform.startswith("win"):
        os.startfile(str(resolved), "print")
        return
    command = "lpr" if sys.platform.startswith("darwin") else (shutil.which("lp") or shutil.which("lpr"))
    if not command:
        raise RuntimeError("لم يتم العثور على أمر طباعة في النظام. افتح المعاينة واطبع من قارئ PDF.")
    subprocess.Popen([command, str(resolved)])


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name

def shape_arabic(value: Any) -> str:
    text = str(value or "")
    if arabic_reshaper is not None and get_display is not None:
        return get_display(arabic_reshaper.reshape(text))
    return text

def resolve_font_paths(family: str) -> tuple[str | None, str | None]:
    candidates = []
    if sys.platform.startswith("win"):
        fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
        mapping = {
            "Segoe UI": ("segoeui.ttf", "segoeuib.ttf"),
            "Tahoma": ("tahoma.ttf", "tahomabd.ttf"),
            "Arial": ("arial.ttf", "arialbd.ttf"),
            "Traditional Arabic": ("trado.ttf", "tradbdo.ttf"),
            "Calibri": ("calibri.ttf", "calibrib.ttf"),
        }
        regular, bold = mapping.get(family, mapping["Segoe UI"])
        candidates.append((str(fonts / regular), str(fonts / bold)))
    candidates.extend([
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf"),
    ])
    for regular, bold in candidates:
        if Path(regular).exists():
            return regular, bold if Path(bold).exists() else regular
    return None, None

def get_smart_mapping(source_keys: list, target_fields: list, threshold=0.6) -> dict:
    """مطابقة عربية/إنجليزية مرنة لربط الحقول القديمة تلقائيًا."""
    def normalize(value: Any) -> str:
        text = clean_value(value).casefold()
        text = re.sub(r"[\u064b-\u065f\u0670\u0640]", "", text)
        text = text.translate(str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه"}))
        return re.sub(r"[^0-9a-z\u0600-\u06ff]", "", text)

    mapping = {}
    for sk in source_keys:
        source = normalize(sk)
        exact = next((f.key for f in target_fields if source in {normalize(f.key), normalize(f.label)}), None)
        if exact:
            mapping[sk] = exact
            continue
        scored = [
            (max(difflib.SequenceMatcher(None, source, normalize(f.label)).ratio(),
                 difflib.SequenceMatcher(None, source, normalize(f.key)).ratio()), f.key)
            for f in target_fields
        ]
        if scored:
            score, matched_key = max(scored)
            if score >= threshold:
                mapping[sk] = matched_key
    return mapping

# ============================================================================
# مدير قاعدة البيانات
# ============================================================================

CURRENT_USER: dict[str, Any] = {
    "id": "builtin-admin", "username": "admin", "display_name": "مدير النظام",
    "role": "مدير", "permissions": ["*"], "active": True,
}
ROLE_PERMISSIONS = {
    "مدير": {"*"},
    "مشرف": {"records.view", "records.create", "records.edit", "records.delete", "records.restore", "export", "import", "attachments", "reports", "backup"},
    "موظف": {"records.view", "records.create", "records.edit", "export", "attachments", "reports"},
    "قراءة فقط": {"records.view", "export", "reports"},
}

def password_hash(password: str, salt: str = "") -> tuple[str, str]:
    salt = salt or os.urandom(16).hex()
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), 180_000)
    return salt, digest.hex()

def password_matches(password: str, salt: str, expected: str) -> bool:
    """Never authenticate empty or unconfigured accounts."""
    if not password or not salt or not expected:
        return False
    try:
        _salt, digest = password_hash(password, salt)
        return hmac.compare_digest(digest, expected)
    except (ValueError, TypeError):
        return False

class DatabaseManager:
    @staticmethod
    def log_action(action: str, record_id: str = "", details: str = "") -> None:
        try:
            with get_db() as conn:
                conn.execute(
                    "INSERT INTO audit_log (created_at, actor, action, record_id, details) VALUES (?, ?, ?, ?, ?)",
                    (
                        dt.datetime.now().isoformat(timespec="seconds"),
                        clean_value(CURRENT_USER.get("display_name", "تطبيق محلي")) or "تطبيق محلي",
                        clean_value(action),
                        clean_value(record_id),
                        clean_value(details)[:1000],
                    ),
                )
                conn.execute(
                    "DELETE FROM audit_log WHERE id NOT IN (SELECT id FROM audit_log ORDER BY id DESC LIMIT 5000)"
                )
        except Exception:
            logging.exception("تعذر تسجيل النشاط")

    @staticmethod
    def recent_activity(limit: int = 500) -> list[dict[str, Any]]:
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT id, created_at, actor, action, record_id, details FROM audit_log ORDER BY id DESC LIMIT ?",
                (max(1, min(5000, int(limit))),),
            ).fetchall()
        return [dict(row) for row in rows]

    @staticmethod
    def enqueue_integration_event(action: str, record_id: str = "", payload: Optional[dict[str, Any]] = None) -> str:
        event_id = uuid.uuid4().hex
        with get_db() as conn:
            if clean_value(action) == "record_upsert" and clean_value(record_id):
                conn.execute(
                    "DELETE FROM integration_outbox WHERE action = 'record_upsert' "
                    "AND record_id = ? AND status IN ('pending', 'failed')",
                    (clean_value(record_id),),
                )
            conn.execute(
                """
                INSERT INTO integration_outbox (
                    id, created_at, action, record_id, payload, status, attempts, last_error
                ) VALUES (?, ?, ?, ?, ?, 'pending', 0, '')
                """,
                (
                    event_id,
                    dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                    clean_value(action),
                    clean_value(record_id),
                    json.dumps(payload if isinstance(payload, dict) else {}, ensure_ascii=False),
                ),
            )
            conn.execute(
                "DELETE FROM integration_outbox WHERE status = 'sent' AND id NOT IN "
                "(SELECT id FROM integration_outbox WHERE status = 'sent' ORDER BY sent_at DESC LIMIT 2000)"
            )
        return event_id

    @staticmethod
    def pending_integration_events(limit: int = 1000) -> list[dict[str, Any]]:
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """
                SELECT id, created_at, action, record_id, payload, attempts
                FROM integration_outbox
                WHERE status IN ('pending', 'failed')
                ORDER BY created_at ASC LIMIT ?
                """,
                (max(1, min(5000, int(limit))),),
            ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            try:
                item["payload"] = json.loads(item.get("payload") or "{}")
            except (TypeError, json.JSONDecodeError):
                item["payload"] = {}
            result.append(item)
        return result

    @staticmethod
    def integration_outbox_counts() -> dict[str, int]:
        counts = {"pending": 0, "failed": 0, "sent": 0}
        with get_db() as conn:
            for status, count in conn.execute(
                "SELECT status, COUNT(*) FROM integration_outbox GROUP BY status"
            ).fetchall():
                counts[clean_value(status)] = int(count or 0)
        return counts

    @staticmethod
    def mark_integration_events(event_ids: list[str], *, sent: bool, error: str = "") -> None:
        ids = [clean_value(item) for item in event_ids if clean_value(item)]
        if not ids:
            return
        placeholders = ",".join("?" for _ in ids)
        with get_db() as conn:
            if sent:
                conn.execute(
                    f"UPDATE integration_outbox SET status = 'sent', attempts = attempts + 1, "
                    f"last_error = '', sent_at = ? WHERE id IN ({placeholders})",
                    [dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"), *ids],
                )
            else:
                conn.execute(
                    f"UPDATE integration_outbox SET status = 'failed', attempts = attempts + 1, "
                    f"last_error = ? WHERE id IN ({placeholders})",
                    [clean_value(error)[:1000], *ids],
                )

    @staticmethod
    def load_session_state() -> dict[str, Any]:
        try:
            with get_db() as conn:
                conn.row_factory = sqlite3.Row
                row = conn.execute(
                    "SELECT payload, updated_at, clean_shutdown FROM session_state WHERE id = 1"
                ).fetchone()
            if not row:
                return {"payload": {}, "updated_at": "", "clean_shutdown": True}
            try:
                payload = json.loads(row["payload"] or "{}")
            except (TypeError, json.JSONDecodeError):
                payload = {}
            return {
                "payload": payload if isinstance(payload, dict) else {},
                "updated_at": clean_value(row["updated_at"]),
                "clean_shutdown": bool(row["clean_shutdown"]),
            }
        except Exception:
            logging.exception("تعذر تحميل حالة الجلسة")
            return {"payload": {}, "updated_at": "", "clean_shutdown": True}

    @staticmethod
    def save_session_state(payload: dict[str, Any], clean_shutdown: bool = False) -> None:
        try:
            with get_db() as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO session_state (id, payload, updated_at, clean_shutdown)
                    VALUES (1, ?, ?, ?)
                    """,
                    (
                        json.dumps(payload if isinstance(payload, dict) else {}, ensure_ascii=False),
                        dt.datetime.now().isoformat(timespec="seconds"),
                        1 if clean_shutdown else 0,
                    ),
                )
        except Exception:
            logging.exception("تعذر حفظ حالة الجلسة")

    @staticmethod
    def archive_report(output: Path, report_format: str, settings: AppSettings,
                       recipient: str, records: list[dict[str, Any]],
                       fields: list[FieldDefinition]) -> str:
        path = Path(output)
        data = path.read_bytes()
        archive_id = uuid.uuid4().hex
        field_labels = [clean_value(item.label) for item in fields]
        searchable_values = []
        for record in records:
            values = record.get("values", {}) if isinstance(record, dict) else {}
            searchable_values.extend(clean_value(value) for value in values.values())
        search_text = normalize_search_text(" ".join([
            settings.company_name, settings.report_title, settings.report_title_en,
            recipient, path.name, report_format, *field_labels, *searchable_values,
        ]))[:1_000_000]
        digest = hashlib.sha256(data).hexdigest()
        with get_db() as conn:
            conn.execute(
                """
                INSERT INTO report_archive (
                    id, created_at, report_format, file_name, report_title, company_name,
                    recipient, record_count, selected_fields, search_text, sha256, file_data
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    archive_id, dt.datetime.now().isoformat(timespec="seconds"),
                    clean_value(report_format).lower(), safe_filename(path.name),
                    clean_value(settings.report_title), clean_value(settings.company_name),
                    clean_value(recipient), len(records),
                    json.dumps(field_labels, ensure_ascii=False), search_text, digest,
                    sqlite3.Binary(data),
                ),
            )
        DatabaseManager.log_action(
            "أرشفة تقرير", details=f"{path.name} | {report_format.upper()} | {len(records)} سجل"
        )
        return archive_id

    @staticmethod
    def search_report_archive(query: str = "", limit: int = 500) -> list[dict[str, Any]]:
        normalized = normalize_search_text(query)
        where = "WHERE search_text LIKE ?" if normalized else ""
        params: list[Any] = [f"%{normalized}%"] if normalized else []
        params.append(max(1, min(5000, int(limit))))
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                f"""
                SELECT id, created_at, report_format, file_name, report_title, company_name,
                       recipient, record_count, selected_fields, search_text, sha256,
                       length(file_data) AS file_size
                FROM report_archive {where}
                ORDER BY created_at DESC LIMIT ?
                """,
                params,
            ).fetchall()
        return [dict(row) for row in rows]

    @staticmethod
    def get_archived_report(report_id: str) -> Optional[dict[str, Any]]:
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT id, report_format, file_name, sha256, file_data FROM report_archive WHERE id = ?",
                (clean_value(report_id),),
            ).fetchone()
        return dict(row) if row else None

    @staticmethod
    def report_archive_for_backup() -> list[dict[str, Any]]:
        rows = DatabaseManager.search_report_archive("", limit=5000)
        return [
            {
                **{key: value for key, value in row.items() if key != "file_size"},
                "archive_name": f"reports/{row['id']}/{safe_filename(row['file_name'])}",
            }
            for row in rows
        ]

    @staticmethod
    def restore_report_archive(items: list[dict[str, Any]], archive: zipfile.ZipFile) -> int:
        restored = 0
        with get_db() as conn:
            for item in items if isinstance(items, list) else []:
                if not isinstance(item, dict):
                    continue
                archive_name = clean_value(item.get("archive_name", ""))
                if not archive_name:
                    continue
                try:
                    data = archive.read(archive_name)
                except (KeyError, OSError):
                    continue
                digest = hashlib.sha256(data).hexdigest()
                fields = item.get("selected_fields", "[]")
                if not isinstance(fields, str):
                    fields = json.dumps(fields, ensure_ascii=False)
                searchable = clean_value(item.get("search_text", "")) or normalize_search_text(" ".join([
                    clean_value(item.get("company_name", "")), clean_value(item.get("report_title", "")),
                    clean_value(item.get("recipient", "")), clean_value(item.get("file_name", "")), fields,
                ]))
                conn.execute(
                    """
                    INSERT OR REPLACE INTO report_archive (
                        id, created_at, report_format, file_name, report_title, company_name,
                        recipient, record_count, selected_fields, search_text, sha256, file_data
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        clean_value(item.get("id")) or uuid.uuid4().hex,
                        clean_value(item.get("created_at")) or dt.datetime.now().isoformat(timespec="seconds"),
                        clean_value(item.get("report_format")), safe_filename(clean_value(item.get("file_name", "report.bin"))),
                        clean_value(item.get("report_title")), clean_value(item.get("company_name")),
                        clean_value(item.get("recipient")), int(item.get("record_count", 0) or 0),
                        fields, searchable, digest, sqlite3.Binary(data),
                    ),
                )
                restored += 1
        return restored

    @staticmethod
    def load_settings() -> AppSettings:
        try:
            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT value FROM settings WHERE key = 'app_settings'")
                row = cursor.fetchone()
                if row:
                    raw = json.loads(row[0])
                    return AppSettings.from_dict(raw)
        except Exception:
            logging.exception("فشل تحميل الإعدادات")
        return AppSettings()

    @staticmethod
    def save_settings(settings: AppSettings) -> None:
        try:
            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                    ("app_settings", json.dumps(settings.to_dict(), ensure_ascii=False))
                )
                conn.commit()
            DatabaseManager.log_action("تحديث الإعدادات", details="تم حفظ إعدادات الهوية والمظهر والحقول")
        except Exception as e:
            logging.exception("فشل حفظ الإعدادات")
            raise

    @staticmethod
    def load_records() -> list[dict[str, Any]]:
        records = []
        try:
            with get_db() as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM records ORDER BY created_at DESC")
                rows = cursor.fetchall()
                for row in rows:
                    rec = dict(row)
                    if rec.get("custom_fields"):
                        try:
                            custom = json.loads(rec["custom_fields"])
                            rec["values"] = {k: v for k, v in custom.items()}
                        except:
                            rec["values"] = {}
                    else:
                        rec["values"] = {}
                    for key in ["material_estimate", "contract_number", "customer_name", "mobile",
                                "marble_color", "sink_type", "marble_edge", "showroom", "salesperson"]:
                        if key in rec and rec[key] is not None:
                            rec["values"][key] = rec[key]
                    records.append({
                        "id": rec["id"],
                        "created_at": rec["created_at"],
                        "updated_at": rec["updated_at"],
                        "values": rec["values"]
                    })
        except Exception as e:
            logging.exception("فشل تحميل السجلات")
        # ترقية عرض السجلات القديمة: العقد الأول مطبخ 1 وأي تكرار لاحق يأخذ الرقم التالي.
        migrated_records: list[dict[str, Any]] = []
        for record in reversed(records):
            ensure_kitchen_identity(record.get("values", {}), migrated_records)
            migrated_records.append(record)
        return records

    @staticmethod
    def save_record(record: dict[str, Any]) -> None:
        values = record.get("values", {})
        if not clean_value(values.get("delivery_status", "")):
            values["delivery_status"] = "جديد"
        base_fields = ["material_estimate", "contract_number", "customer_name", "mobile",
                       "marble_color", "sink_type", "marble_edge", "showroom", "salesperson"]
        base_values = {k: values.get(k, "") for k in base_fields}
        custom = {k: v for k, v in values.items() if k not in base_fields}
        custom_json = json.dumps(custom, ensure_ascii=False) if custom else None

        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO records (
                    id, created_at, updated_at,
                    material_estimate, contract_number, customer_name, mobile,
                    marble_color, sink_type, marble_edge, showroom, salesperson,
                    custom_fields
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                record["id"],
                record.get("created_at", dt.datetime.now().isoformat(timespec="seconds")),
                record.get("updated_at", dt.datetime.now().isoformat(timespec="seconds")),
                base_values["material_estimate"],
                base_values["contract_number"],
                base_values["customer_name"],
                base_values["mobile"],
                base_values["marble_color"],
                base_values["sink_type"],
                base_values["marble_edge"],
                base_values["showroom"],
                base_values["salesperson"],
                custom_json
            ))
            conn.commit()
        action = "إضافة سجل" if record.get("created_at") == record.get("updated_at") else "تحديث سجل"
        DatabaseManager.log_action(action, record.get("id", ""), values.get("contract_number", ""))
        DatabaseManager.enqueue_integration_event(
            "record_upsert", record.get("id", ""), {"record": copy.deepcopy(record)}
        )

    @staticmethod
    def save_records_bulk(records: list[dict[str, Any]]) -> None:
        if not records:
            return
        base_fields = ["material_estimate", "contract_number", "customer_name", "mobile",
                       "marble_color", "sink_type", "marble_edge", "showroom", "salesperson"]
        rows = []
        now = dt.datetime.now().isoformat(timespec="seconds")
        for record in records:
            values = record.get("values", {}) if isinstance(record.get("values", {}), dict) else {}
            if not clean_value(values.get("delivery_status", "")):
                values["delivery_status"] = "جديد"
            custom = {k: clean_value(v) for k, v in values.items() if k not in base_fields}
            rows.append((
                str(record.get("id") or uuid.uuid4().hex),
                str(record.get("created_at") or now),
                str(record.get("updated_at") or now),
                *[clean_value(values.get(key, "")) for key in base_fields],
                json.dumps(custom, ensure_ascii=False) if custom else None,
            ))
        with get_db() as conn:
            conn.executemany("""
                INSERT OR REPLACE INTO records (
                    id, created_at, updated_at,
                    material_estimate, contract_number, customer_name, mobile,
                    marble_color, sink_type, marble_edge, showroom, salesperson,
                    custom_fields
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, rows)
            conn.commit()
        DatabaseManager.log_action("استيراد سجلات", details=f"عدد السجلات: {len(records)}")
        DatabaseManager.enqueue_integration_event(
            "records_bulk_import", payload={"record_ids": [row[0] for row in rows], "count": len(rows)}
        )

    @staticmethod
    def contract_numbers() -> set[str]:
        with get_db() as conn:
            rows = conn.execute(
                "SELECT contract_number FROM records WHERE TRIM(COALESCE(contract_number, '')) <> ''"
            ).fetchall()
        return {clean_value(row[0]) for row in rows if clean_value(row[0])}

    @staticmethod
    def next_contract_number(prefix: str = "MD") -> str:
        safe_prefix = re.sub(r"[^A-Za-z0-9_-]", "", clean_value(prefix).upper()) or "MD"
        year = dt.date.today().year
        pattern = re.compile(rf"^{re.escape(safe_prefix)}-{year}-(\d+)$", re.IGNORECASE)
        highest = 0
        for number in DatabaseManager.contract_numbers():
            match = pattern.match(number)
            if match:
                highest = max(highest, int(match.group(1)))
        return f"{safe_prefix}-{year}-{highest + 1:04d}"

    @staticmethod
    def delete_records(ids: list[str]) -> None:
        """نقل السجلات إلى سلة المحذوفات مع الاحتفاظ بالمرفقات حتى التفريغ النهائي."""
        if not ids:
            return
        with get_db() as conn:
            cursor = conn.cursor()
            for record_id in ids:
                row = cursor.execute(
                    """SELECT id, created_at, updated_at, material_estimate, contract_number, customer_name, mobile,
                              marble_color, sink_type, marble_edge, showroom, salesperson, custom_fields
                       FROM records WHERE id = ?""", (record_id,)
                ).fetchone()
                if not row:
                    continue
                try:
                    custom = json.loads(row[12] or "{}")
                except Exception:
                    custom = {}
                values = {
                    "material_estimate": row[3] or "", "contract_number": row[4] or "", "customer_name": row[5] or "",
                    "mobile": row[6] or "", "marble_color": row[7] or "", "sink_type": row[8] or "",
                    "marble_edge": row[9] or "", "showroom": row[10] or "", "salesperson": row[11] or "", **custom,
                }
                payload = {"id": row[0], "created_at": row[1], "updated_at": row[2], "values": values}
                attachments = [
                    {"id": a[0], "original_name": a[1], "stored_path": a[2], "added_at": a[3]}
                    for a in cursor.execute(
                        "SELECT id, original_name, stored_path, added_at FROM attachments WHERE record_id = ?", (record_id,)
                    ).fetchall()
                ]
                cursor.execute(
                    "INSERT OR REPLACE INTO deleted_records(id, deleted_at, payload, attachment_manifest) VALUES (?, ?, ?, ?)",
                    (record_id, dt.datetime.now().isoformat(timespec="seconds"), json.dumps(payload, ensure_ascii=False), json.dumps(attachments, ensure_ascii=False)),
                )
                cursor.execute("DELETE FROM records WHERE id = ?", (record_id,))
            conn.commit()
        DatabaseManager.log_action("نقل إلى سلة المحذوفات", details=f"عدد السجلات: {len(ids)}")
        DatabaseManager.enqueue_integration_event(
            "records_soft_delete", payload={"record_ids": list(ids), "count": len(ids)}
        )

    @staticmethod
    def search_records(query: str, filters: Optional[dict] = None) -> list[dict[str, Any]]:
        base_fields = ["material_estimate", "contract_number", "customer_name", "mobile",
                       "marble_color", "sink_type", "marble_edge", "showroom", "salesperson"]
        conditions = []
        params = []
        if query:
            like_conditions = []
            for field in base_fields:
                like_conditions.append(f"{field} LIKE ?")
                params.append(f"%{query}%")
            like_conditions.append("custom_fields LIKE ?")
            params.append(f"%{query}%")
            conditions.append("(" + " OR ".join(like_conditions) + ")")

        if filters:
            for field, value in filters.items():
                if field in base_fields and value:
                    conditions.append(f"{field} = ?")
                    params.append(value)
                elif value:
                    json_path = f"$.{field}"
                    if field == "delivery_status" and value == "جديد":
                        conditions.append("(json_extract(COALESCE(custom_fields, '{}'), ?) = ? OR json_extract(COALESCE(custom_fields, '{}'), ?) IS NULL OR json_extract(COALESCE(custom_fields, '{}'), ?) = '')")
                        params.extend([json_path, value, json_path, json_path])
                    else:
                        conditions.append("json_extract(COALESCE(custom_fields, '{}'), ?) = ?")
                        params.extend([json_path, value])

        where_clause = " AND ".join(conditions) if conditions else "1=1"
        sql = f"SELECT * FROM records WHERE {where_clause} ORDER BY created_at DESC"

        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(sql, params)
            rows = cursor.fetchall()
            records = []
            for row in rows:
                rec = dict(row)
                if rec.get("custom_fields"):
                    try:
                        custom = json.loads(rec["custom_fields"])
                        rec["values"] = {k: v for k, v in custom.items()}
                    except:
                        rec["values"] = {}
                else:
                    rec["values"] = {}
                for key in base_fields:
                    if key in rec and rec[key] is not None:
                        rec["values"][key] = rec[key]
                records.append({
                    "id": rec["id"],
                    "created_at": rec["created_at"],
                    "updated_at": rec["updated_at"],
                    "values": rec["values"]
                })
            return records

    @staticmethod
    def list_attachments(record_id: str) -> list[dict[str, Any]]:
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT id, record_id, original_name, stored_path, added_at FROM attachments WHERE record_id = ? ORDER BY added_at DESC",
                (record_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    @staticmethod
    def all_attachments() -> list[dict[str, Any]]:
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT id, record_id, original_name, stored_path, added_at FROM attachments ORDER BY added_at"
            ).fetchall()
        return [dict(row) for row in rows]

    @staticmethod
    def attachment_record_date(record_id: str) -> str:
        """يربط مجلد المرفقات بتاريخ التركيب ثم موعد التسليم ثم تاريخ إنشاء السجل."""
        with get_db() as conn:
            row = conn.execute(
                "SELECT created_at, custom_fields FROM records WHERE id = ?", (record_id,),
            ).fetchone()
        if not row:
            raise ValueError("السجل المرتبط بالمرفق غير موجود.")
        values: dict[str, Any] = {}
        try:
            values = json.loads(row[1] or "{}") if row[1] else {}
        except (TypeError, json.JSONDecodeError):
            values = {}
        candidate = (
            clean_value(values.get("actual_delivery_date", ""))
            or clean_value(values.get("scheduled_delivery_date", ""))
            or clean_value(row[0])[:10]
            or dt.date.today().isoformat()
        )
        try:
            return dt.datetime.strptime(candidate[:10], "%Y-%m-%d").date().isoformat()
        except ValueError:
            return dt.date.today().isoformat()

    @staticmethod
    def restore_attachment_rows(items: list[dict[str, Any]], archive: zipfile.ZipFile) -> int:
        restored = 0
        valid_record_ids = {record["id"] for record in DatabaseManager.load_records()}
        for item in items:
            record_id = clean_value(item.get("record_id", ""))
            archive_name = clean_value(item.get("archive_name", ""))
            if not record_id or record_id not in valid_record_ids or not archive_name.startswith("attachments/"):
                continue
            try:
                info = archive.getinfo(archive_name)
            except KeyError:
                continue
            if info.file_size > 50 * 1024 * 1024:
                continue
            attachment_id = clean_value(item.get("id", "")) or uuid.uuid4().hex
            original_name = clean_value(item.get("original_name", "")) or Path(archive_name).name
            record_date = DatabaseManager.attachment_record_date(record_id)
            target_dir = ATTACHMENTS_DIR / record_date / record_id
            target_dir.mkdir(parents=True, exist_ok=True)
            target = target_dir / f"{attachment_id[:10]}_{safe_filename(original_name)}"
            with archive.open(info, "r") as source, target.open("wb") as destination:
                shutil.copyfileobj(source, destination)
            added_at = clean_value(item.get("added_at", "")) or dt.datetime.now().isoformat(timespec="seconds")
            with get_db() as conn:
                conn.execute(
                    "INSERT OR REPLACE INTO attachments (id, record_id, original_name, stored_path, added_at) VALUES (?, ?, ?, ?, ?)",
                    (attachment_id, record_id, original_name, str(target), added_at),
                )
            restored += 1
        return restored

    @staticmethod
    def add_attachment(record_id: str, source_path: Path) -> dict[str, Any]:
        source_path = source_path.resolve()
        if not source_path.is_file():
            raise FileNotFoundError("ملف المرفق غير موجود.")
        if source_path.stat().st_size > 50 * 1024 * 1024:
            raise ValueError("حجم المرفق يتجاوز 50 ميجابايت.")
        attachment_id = uuid.uuid4().hex
        record_date = DatabaseManager.attachment_record_date(record_id)
        target_dir = ATTACHMENTS_DIR / record_date / record_id
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / f"{attachment_id[:10]}_{safe_filename(source_path.name)}"
        shutil.copy2(source_path, target)
        added_at = dt.datetime.now().isoformat(timespec="seconds")
        with get_db() as conn:
            conn.execute(
                "INSERT INTO attachments (id, record_id, original_name, stored_path, added_at) VALUES (?, ?, ?, ?, ?)",
                (attachment_id, record_id, source_path.name, str(target), added_at),
            )
        DatabaseManager.log_action("إضافة مرفق", record_id, source_path.name)
        return {"id": attachment_id, "record_id": record_id, "original_name": source_path.name,
                "stored_path": str(target), "added_at": added_at}

    @staticmethod
    def delete_attachment(attachment_id: str) -> None:
        with get_db() as conn:
            row = conn.execute(
                "SELECT record_id, original_name, stored_path FROM attachments WHERE id = ?", (attachment_id,)
            ).fetchone()
            if not row:
                return
            conn.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
        resolved = Path(row[2]).resolve()
        if resolved.is_relative_to(ATTACHMENTS_DIR.resolve()):
            resolved.unlink(missing_ok=True)
        DatabaseManager.log_action("حذف مرفق", row[0], row[1])

    @staticmethod
    def get_records_in_date_range(start_date: str = "", end_date: str = "") -> list[dict[str, Any]]:
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            conditions = []
            params = []
            if start_date:
                conditions.append("date(created_at) >= ?")
                params.append(start_date)
            if end_date:
                conditions.append("date(created_at) <= ?")
                params.append(end_date)
            where_clause = " AND ".join(conditions) if conditions else "1=1"
            cursor.execute(
                f"SELECT * FROM records WHERE {where_clause} ORDER BY created_at DESC",
                params,
            )
            rows = cursor.fetchall()
            records = []
            base_fields = ["material_estimate", "contract_number", "customer_name", "mobile",
                           "marble_color", "sink_type", "marble_edge", "showroom", "salesperson"]
            for row in rows:
                rec = dict(row)
                if rec.get("custom_fields"):
                    try:
                        custom = json.loads(rec["custom_fields"])
                        rec["values"] = {k: v for k, v in custom.items()}
                    except:
                        rec["values"] = {}
                else:
                    rec["values"] = {}
                for key in base_fields:
                    if key in rec and rec[key] is not None:
                        rec["values"][key] = rec[key]
                records.append({
                    "id": rec["id"],
                    "created_at": rec["created_at"],
                    "updated_at": rec["updated_at"],
                    "values": rec["values"]
                })
            return records

    @staticmethod
    def export_to_csv(records: list[dict[str, Any]], fields: list[FieldDefinition], output_path: Path) -> None:
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([f.label for f in fields])
            for rec in records:
                row = [rec.get("values", {}).get(f.key, "") for f in fields]
                writer.writerow(row)


    @staticmethod
    def list_deleted_records(limit: int = 1000) -> list[dict[str, Any]]:
        with get_db() as conn:
            rows = conn.execute(
                "SELECT id, deleted_at, payload, attachment_manifest FROM deleted_records ORDER BY deleted_at DESC LIMIT ?",
                (max(1, min(5000, int(limit))),),
            ).fetchall()
        items = []
        for row in rows:
            try:
                payload = json.loads(row[2] or "{}")
            except Exception:
                payload = {}
            items.append({"id": row[0], "deleted_at": row[1], "record": payload})
        return items

    @staticmethod
    def restore_deleted_records(ids: list[str]) -> int:
        restored = 0
        if not ids:
            return 0
        with get_db() as conn:
            cursor = conn.cursor()
            for record_id in ids:
                row = cursor.execute(
                    "SELECT payload, attachment_manifest FROM deleted_records WHERE id = ?", (record_id,)
                ).fetchone()
                if not row:
                    continue
                try:
                    record = json.loads(row[0] or "{}")
                    attachments = json.loads(row[1] or "[]")
                except Exception:
                    continue
                values = record.get("values", {}) if isinstance(record, dict) else {}
                custom_values = dict(values)
                base_keys = ["material_estimate", "contract_number", "customer_name", "mobile", "marble_color", "sink_type", "marble_edge", "showroom", "salesperson"]
                custom_fields = {k: v for k, v in custom_values.items() if k not in base_keys}
                cursor.execute(
                    """INSERT OR REPLACE INTO records
                       (id, created_at, updated_at, material_estimate, contract_number, customer_name, mobile,
                        marble_color, sink_type, marble_edge, showroom, salesperson, custom_fields)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        clean_value(record.get("id", record_id)) or record_id,
                        clean_value(record.get("created_at", dt.datetime.now().isoformat(timespec="seconds"))),
                        dt.datetime.now().isoformat(timespec="seconds"),
                        *[clean_value(values.get(key, "")) for key in base_keys],
                        json.dumps(custom_fields, ensure_ascii=False),
                    ),
                )
                for item in attachments if isinstance(attachments, list) else []:
                    stored = clean_value(item.get("stored_path", ""))
                    if not stored or not Path(stored).is_file():
                        continue
                    cursor.execute(
                        """INSERT OR IGNORE INTO attachments
                           (id, record_id, original_name, stored_path, added_at) VALUES (?, ?, ?, ?, ?)""",
                        (
                            clean_value(item.get("id", "")) or str(uuid.uuid4()), record_id,
                            clean_value(item.get("original_name", Path(stored).name)), stored,
                            clean_value(item.get("added_at", dt.datetime.now().isoformat(timespec="seconds"))),
                        ),
                    )
                cursor.execute("DELETE FROM deleted_records WHERE id = ?", (record_id,))
                restored += 1
            conn.commit()
        if restored:
            DatabaseManager.log_action("استعادة من سلة المحذوفات", details=f"عدد السجلات: {restored}")
        return restored

    @staticmethod
    def purge_deleted_records(ids: list[str]) -> int:
        if not ids:
            return 0
        removed = 0
        with get_db() as conn:
            cursor = conn.cursor()
            for record_id in ids:
                row = cursor.execute("SELECT attachment_manifest FROM deleted_records WHERE id = ?", (record_id,)).fetchone()
                if not row:
                    continue
                try:
                    attachments = json.loads(row[0] or "[]")
                except Exception:
                    attachments = []
                for item in attachments if isinstance(attachments, list) else []:
                    stored = clean_value(item.get("stored_path", ""))
                    if stored:
                        try:
                            path = Path(stored).resolve()
                            if path.is_relative_to(ATTACHMENTS_DIR.resolve()):
                                path.unlink(missing_ok=True)
                        except OSError:
                            logging.exception("تعذر حذف مرفق نهائياً من السلة")
                cursor.execute("DELETE FROM deleted_records WHERE id = ?", (record_id,))
                removed += 1
            conn.commit()
        if removed:
            DatabaseManager.log_action("تفريغ سلة المحذوفات", details=f"عدد السجلات: {removed}")
        return removed

    @staticmethod
    def list_directory(entity_type: str) -> list[dict[str, Any]]:
        with get_db() as conn:
            rows = conn.execute(
                "SELECT id, entity_type, name, mobile, details, created_at, updated_at FROM directories WHERE entity_type = ? ORDER BY name COLLATE NOCASE",
                (clean_value(entity_type),),
            ).fetchall()
        return [
            {"id": row[0], "entity_type": row[1], "name": row[2], "mobile": row[3], "details": row[4], "created_at": row[5], "updated_at": row[6]}
            for row in rows
        ]

    @staticmethod
    def save_directory(entity_type: str, name: str, mobile: str = "", details: str = "") -> None:
        entity_type, name = clean_value(entity_type), clean_value(name)
        if not entity_type or not name:
            raise ValueError("نوع الدليل والاسم مطلوبان.")
        now = dt.datetime.now().isoformat(timespec="seconds")
        with get_db() as conn:
            conn.execute(
                """INSERT INTO directories(entity_type, name, mobile, details, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT(entity_type, name) DO UPDATE SET mobile=excluded.mobile, details=excluded.details, updated_at=excluded.updated_at""",
                (entity_type, name, clean_value(mobile), clean_value(details), now, now),
            )
        DatabaseManager.log_action("تحديث دليل البيانات", details=f"{entity_type}: {name}")

    @staticmethod
    def delete_directory(entity_type: str, name: str) -> None:
        with get_db() as conn:
            conn.execute("DELETE FROM directories WHERE entity_type = ? AND name = ?", (clean_value(entity_type), clean_value(name)))

    @staticmethod
    def learn_directories_from_records() -> int:
        records = DatabaseManager.load_records()
        added = 0
        for record in records:
            values = record.get("values", {})
            for entity_type, key in (("customer", "customer_name"), ("showroom", "showroom"), ("salesperson", "salesperson")):
                name = clean_value(values.get(key, ""))
                if not name:
                    continue
                mobile = clean_value(values.get("mobile", "")) if entity_type == "customer" else ""
                DatabaseManager.save_directory(entity_type, name, mobile, "")
                added += 1
        return added

    @staticmethod
    def save_report_template(name: str, settings: AppSettings) -> None:
        name = clean_value(name)
        if not name:
            raise ValueError("اسم القالب مطلوب.")
        payload = {
            "report_page_size": settings.report_page_size,
            "report_orientation": settings.report_orientation,
            "report_margin_mm": settings.report_margin_mm,
            "pdf_base_font_size": settings.pdf_base_font_size,
            "pdf_min_font_size": settings.pdf_min_font_size,
            "png_dpi": settings.png_dpi,
            "jpg_quality": settings.jpg_quality,
            "custom_page_width_mm": settings.custom_page_width_mm,
            "custom_page_height_mm": settings.custom_page_height_mm,
            "logo_width_mm": settings.logo_width_mm,
            "print_friendly_reports": settings.print_friendly_reports,
            "title_bg_color": settings.title_bg_color,
            "title_fg_color": settings.title_fg_color,
            "subtitle_bg_color": settings.subtitle_bg_color,
            "subtitle_fg_color": settings.subtitle_fg_color,
            "export_preferences": copy.deepcopy(settings.export_preferences),
        }
        now = dt.datetime.now().isoformat(timespec="seconds")
        with get_db() as conn:
            conn.execute(
                """INSERT INTO report_templates(id, name, created_at, updated_at, payload) VALUES (?, ?, ?, ?, ?)
                   ON CONFLICT(name) DO UPDATE SET updated_at=excluded.updated_at, payload=excluded.payload""",
                (str(uuid.uuid4()), name, now, now, json.dumps(payload, ensure_ascii=False)),
            )
        DatabaseManager.log_action("حفظ قالب تقرير", details=name)

    @staticmethod
    def list_report_templates() -> list[dict[str, Any]]:
        with get_db() as conn:
            rows = conn.execute("SELECT id, name, created_at, updated_at, payload FROM report_templates ORDER BY name").fetchall()
        result = []
        for row in rows:
            try:
                payload = json.loads(row[4] or "{}")
            except Exception:
                payload = {}
            result.append({"id": row[0], "name": row[1], "created_at": row[2], "updated_at": row[3], "payload": payload})
        return result

    @staticmethod
    def delete_report_template(template_id: str) -> None:
        with get_db() as conn:
            conn.execute("DELETE FROM report_templates WHERE id = ?", (template_id,))

    @staticmethod
    def list_users(active_only: bool = False) -> list[dict[str, Any]]:
        query = "SELECT id, username, display_name, role, permissions, salt, password_hash, active, created_at, updated_at FROM users"
        params: tuple[Any, ...] = ()
        if active_only:
            query += " WHERE active = 1"
        query += " ORDER BY display_name, username"
        with get_db() as conn:
            rows = conn.execute(query, params).fetchall()
        result = []
        for row in rows:
            try:
                permissions = json.loads(row[4] or "[]")
            except Exception:
                permissions = []
            result.append({
                "id": row[0], "username": row[1], "display_name": row[2], "role": row[3],
                "permissions": permissions, "salt": row[5], "password_hash": row[6],
                "active": bool(row[7]), "created_at": row[8], "updated_at": row[9],
            })
        return result

    @staticmethod
    def save_user(username: str, display_name: str, role: str, password: str = "", *, user_id: str = "", active: bool = True) -> str:
        username = re.sub(r"[^A-Za-z0-9_.-]", "", clean_value(username)).lower()
        display_name = clean_value(display_name)
        if not username or not display_name:
            raise ValueError("اسم المستخدم والاسم الظاهر مطلوبان.")
        if role not in ROLE_PERMISSIONS:
            role = "موظف"
        now = dt.datetime.now().isoformat(timespec="seconds")
        existing = None
        if user_id:
            existing = next((u for u in DatabaseManager.list_users() if u["id"] == user_id), None)
        if password:
            salt, digest = password_hash(password)
        elif existing:
            salt, digest = existing.get("salt", ""), existing.get("password_hash", "")
        else:
            salt, digest = "", ""
        user_id = user_id or str(uuid.uuid4())
        permissions = sorted(ROLE_PERMISSIONS.get(role, set()))
        with get_db() as conn:
            conn.execute(
                """INSERT INTO users(id, username, display_name, role, permissions, salt, password_hash, active, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(id) DO UPDATE SET username=excluded.username, display_name=excluded.display_name,
                       role=excluded.role, permissions=excluded.permissions, salt=excluded.salt,
                       password_hash=excluded.password_hash, active=excluded.active, updated_at=excluded.updated_at""",
                (user_id, username, display_name, role, json.dumps(permissions, ensure_ascii=False), salt, digest, int(active),
                 existing.get("created_at", now) if existing else now, now),
            )
        return user_id

    @staticmethod
    def authenticate(username: str, password: str) -> Optional[dict[str, Any]]:
        username = clean_value(username).lower()
        user = next((u for u in DatabaseManager.list_users(active_only=True) if u["username"].lower() == username), None)
        if user and password_matches(password, user.get("salt", ""), user.get("password_hash", "")):
            return user
        return None

    @staticmethod
    def delete_user(user_id: str) -> None:
        if user_id == "builtin-admin":
            raise ValueError("لا يمكن حذف حساب مدير النظام الأساسي؛ يمكن تعديل كلمة مروره.")
        with get_db() as conn:
            conn.execute("DELETE FROM users WHERE id = ?", (user_id,))

    @staticmethod
    def repair_database() -> str:
        with get_db() as conn:
            before = conn.execute("PRAGMA quick_check").fetchone()
            conn.execute("PRAGMA optimize")
            conn.execute("REINDEX")
            after = conn.execute("PRAGMA quick_check").fetchone()
        return f"قبل الإصلاح: {before[0] if before else 'غير معروف'} | بعد الإصلاح: {after[0] if after else 'غير معروف'}"

@dataclass
class SpreadsheetImportPreview:
    source_path: Path
    sheet_name: str
    header_row: int
    headers: list[str]
    data_rows: list[list[str]]
    suggested_mapping: dict[int, str]


class SpreadsheetImportReader:
    """قراءة مرنة لملفات XLSX/XLSM/XLS/CSV مع اكتشاف العناوين تلقائيًا."""

    FIELD_ALIASES: dict[str, list[str]] = {
        "material_estimate": ["تقدير الخامات", "تقدير خامات", "الخامات", "كمية الخامات", "الكمية", "materials"],
        "contract_number": ["رقم العقد", "رقم عقد", "العقد", "contract", "contract no", "contract number"],
        "kitchen_number": ["رقم المطبخ", "اسم المطبخ", "المطبخ", "الوحدة", "kitchen", "kitchen no", "kitchen name"],
        "customer_name": ["اسم العميل", "العميل", "اسم الزبون", "الزبون", "customer", "client"],
        "mobile": ["رقم الجوال", "الجوال", "رقم الهاتف", "الهاتف", "الموبايل", "mobile", "phone"],
        "marble_color": ["لون الرخام", "اللون", "كود اللون", "marble color", "color"],
        "sink_type": ["نوع المجلى", "المجلى", "نوع الحوض", "الحوض", "sink type", "sink"],
        "marble_edge": ["شفة الرخام", "الشفة", "حافة الرخام", "الحافة", "marble edge", "edge"],
        "showroom": ["المعرض", "اسم المعرض", "الفرع", "showroom", "branch"],
        "salesperson": ["موظف المعرض", "الموظف", "مندوب المبيعات", "البائع", "salesperson", "sales rep"],
        "delivery_status": ["حالة التسليم", "الحالة", "status", "delivery status"],
        "scheduled_delivery_date": ["موعد التسليم", "تاريخ التسليم المتوقع", "scheduled delivery", "due date"],
        "actual_delivery_date": ["تاريخ التركيب الفعلي", "تاريخ التسليم الفعلي", "تم التسليم بتاريخ",
                                 "actual delivery date", "installation date"],
    }
    _similarity_cache: dict[tuple[str, str], float] = {}

    def __init__(self, settings: AppSettings):
        self.settings = settings
        self.fields = settings.fields
        self.aliases: dict[str, set[str]] = {}
        for item in self.fields:
            candidates = [item.key, item.label, *self.FIELD_ALIASES.get(item.key, [])]
            self.aliases[item.key] = {self.normalize_header(value) for value in candidates if value}

    @staticmethod
    def normalize_header(value: Any) -> str:
        text = clean_value(value).casefold()
        text = re.sub(r"[\u064b-\u065f\u0670\u0640]", "", text)
        text = text.translate(str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه"}))
        return re.sub(r"[^0-9a-z\u0600-\u06ff]", "", text)

    @staticmethod
    def _plain_number(value: Any, number_format: str = "") -> str:
        if isinstance(value, bool):
            return "نعم" if value else "لا"
        if isinstance(value, dt.datetime):
            return value.strftime("%Y-%m-%d %H:%M")
        if isinstance(value, dt.date):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, (int, float)):
            fmt = str(number_format or "").strip()
            if re.fullmatch(r"0+", fmt):
                return str(int(value)).zfill(len(fmt))
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
        return clean_value(value)

    def suggest_mapping(self, headers: list[str]) -> dict[int, str]:
        mapping: dict[int, str] = {}
        used: set[str] = set()
        for index, header in enumerate(headers):
            normalized = self.normalize_header(header)
            if not normalized:
                continue
            exact_key = next(
                (key for key, aliases in self.aliases.items() if key not in used and normalized in aliases),
                None,
            )
            if exact_key:
                mapping[index] = exact_key
                used.add(exact_key)
                continue
            candidates: list[tuple[float, str]] = []
            for key, aliases in self.aliases.items():
                if key in used:
                    continue
                ratios = []
                for alias in aliases:
                    cache_key = (normalized, alias)
                    ratio = self._similarity_cache.get(cache_key)
                    if ratio is None:
                        ratio = difflib.SequenceMatcher(None, normalized, alias).ratio()
                        self._similarity_cache[cache_key] = ratio
                    ratios.append(ratio)
                ratio = max(ratios, default=0.0)
                candidates.append((ratio, key))
            if candidates:
                score, key = max(candidates)
                if score >= 0.72:
                    mapping[index] = key
                    used.add(key)
        return mapping

    def _best_header(self, sheets: list[tuple[str, list[list[str]]]]) -> tuple[str, int, list[str], dict[int, str]]:
        best: tuple[int, int, str, int, list[str], dict[int, str]] | None = None
        for sheet_name, rows in sheets:
            for row_index, row in enumerate(rows[:30]):
                headers = [clean_value(value) for value in row[:80]]
                nonempty = sum(bool(value) for value in headers)
                if not nonempty:
                    continue
                mapping = self.suggest_mapping(headers)
                score = len(mapping) * 100 + min(nonempty, 40)
                candidate = (score, nonempty, sheet_name, row_index, headers, mapping)
                if best is None or candidate[:2] > best[:2]:
                    best = candidate
        if best is None:
            raise ValueError("الملف لا يحتوي على صفوف قابلة للقراءة.")
        return best[2], best[3], best[4], best[5]

    @staticmethod
    def _trim_data_rows(rows: list[list[str]]) -> list[list[str]]:
        result: list[list[str]] = []
        empty_run = 0
        started = False
        for row in rows:
            normalized = [clean_value(value) for value in row]
            if not any(normalized):
                if started:
                    empty_run += 1
                    if empty_run >= 2:
                        break
                continue
            started = True
            empty_run = 0
            result.append(normalized)
        return result

    def _read_xlsx(self, path: Path) -> SpreadsheetImportPreview:
        if load_workbook is None:
            raise RuntimeError("مكتبة openpyxl غير مثبتة. شغّل ملف تثبيت المتطلبات.")
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            scan_sheets: list[tuple[str, list[list[str]]]] = []
            for worksheet in workbook.worksheets:
                scan_rows = []
                for cells in worksheet.iter_rows(min_row=1, max_row=min(30, worksheet.max_row or 1), max_col=min(80, worksheet.max_column or 1)):
                    scan_rows.append([self._plain_number(cell.value, getattr(cell, "number_format", "")) for cell in cells])
                scan_sheets.append((worksheet.title, scan_rows))
            sheet_name, header_index, headers, mapping = self._best_header(scan_sheets)
            worksheet = workbook[sheet_name]
            max_col = max(1, min(80, max(len(headers), worksheet.max_column or 1)))
            data_rows = [
                [self._plain_number(cell.value, getattr(cell, "number_format", "")) for cell in cells]
                for cells in worksheet.iter_rows(min_row=header_index + 2, max_col=max_col)
            ]
            return SpreadsheetImportPreview(path, sheet_name, header_index + 1, headers, self._trim_data_rows(data_rows), mapping)
        finally:
            workbook.close()

    def _read_xls(self, path: Path) -> SpreadsheetImportPreview:
        if xlrd is None:
            raise RuntimeError("قراءة ملفات XLS القديمة تحتاج مكتبة xlrd. شغّل ملف تثبيت المتطلبات ثم أعد المحاولة.")
        book = xlrd.open_workbook(str(path), on_demand=True, formatting_info=True)
        try:
            scan_sheets: list[tuple[str, list[list[str]]]] = []
            for sheet in book.sheets():
                rows = [[self._plain_number(sheet.cell_value(r, c)) for c in range(min(sheet.ncols, 80))]
                        for r in range(min(sheet.nrows, 30))]
                scan_sheets.append((sheet.name, rows))
            sheet_name, header_index, headers, mapping = self._best_header(scan_sheets)
            sheet = book.sheet_by_name(sheet_name)
            data_rows: list[list[str]] = []
            for row_index in range(header_index + 1, sheet.nrows):
                row: list[str] = []
                for col_index in range(min(sheet.ncols, 80)):
                    cell = sheet.cell(row_index, col_index)
                    number_format = ""
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        value = dt.datetime(*xlrd.xldate_as_tuple(cell.value, book.datemode))
                    else:
                        value = cell.value
                    if cell.ctype == xlrd.XL_CELL_NUMBER:
                        try:
                            xf = book.xf_list[cell.xf_index]
                            number_format = book.format_map[xf.format_key].format_str
                        except (AttributeError, IndexError, KeyError):
                            number_format = ""
                    row.append(self._plain_number(value, number_format))
                data_rows.append(row)
            return SpreadsheetImportPreview(path, sheet_name, header_index + 1, headers, self._trim_data_rows(data_rows), mapping)
        finally:
            book.release_resources()

    def _read_csv(self, path: Path) -> SpreadsheetImportPreview:
        rows: list[list[str]] | None = None
        for encoding in ("utf-8-sig", "utf-8", "cp1256", "windows-1256"):
            try:
                with open(path, newline="", encoding=encoding) as stream:
                    rows = [[clean_value(value) for value in row] for row in csv.reader(stream)]
                break
            except UnicodeDecodeError:
                continue
        if rows is None:
            raise ValueError("تعذر تحديد ترميز ملف CSV.")
        sheet_name, header_index, headers, mapping = self._best_header([(path.stem, rows[:30])])
        return SpreadsheetImportPreview(path, sheet_name, header_index + 1, headers, self._trim_data_rows(rows[header_index + 1:]), mapping)

    def read(self, path: Path) -> SpreadsheetImportPreview:
        suffix = path.suffix.lower()
        if suffix in {".xlsx", ".xlsm"}:
            return self._read_xlsx(path)
        if suffix == ".xls":
            return self._read_xls(path)
        if suffix == ".csv":
            return self._read_csv(path)
        raise ValueError("الصيغ المدعومة للاستيراد: XLSX وXLSM وXLS وCSV.")

    def records_from_preview(self, preview: SpreadsheetImportPreview, mapping: dict[int, str]) -> list[dict[str, Any]]:
        now = dt.datetime.now().isoformat(timespec="seconds")
        records: list[dict[str, Any]] = []
        for row in preview.data_rows:
            values = {item.key: "" for item in self.fields}
            for column_index, field_key in mapping.items():
                if column_index < len(row):
                    values[field_key] = clean_value(row[column_index])
            if not any(values.values()):
                continue
            records.append({
                "id": uuid.uuid4().hex,
                "created_at": now,
                "updated_at": now,
                "values": values,
            })
        return records

# ============================================================================
# نافذة التقدم (شريط التقدم)
# ============================================================================
class ProgressDialog:
    def __init__(self, parent_app: "MarbleDeliveryApp", title="جارٍ التنفيذ..."):
        self.parent_app = parent_app
        parent = parent_app.root
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("440x125")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)
        self.window.configure(bg=parent_app.settings.surface_color)
        self.label = tk.Label(
            self.window, text=title, bg=parent_app.settings.surface_color,
            fg=parent_app.settings.text_color, font=(parent_app.settings.font_family, 11, "bold"),
        )
        self.label.pack(pady=(18, 10))
        self.progress = ttk.Progressbar(self.window, length=380, mode="indeterminate")
        self.progress.pack(pady=6)
        self.progress.start(10)
        self.window.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self.window.winfo_width()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - self.window.winfo_height()) // 2)
        self.window.geometry(f"+{x}+{y}")
        self.window.update()

    def update_text(self, text):
        self.label.config(text=text)
        self.window.update_idletasks()

    def close(self):
        if self.window.winfo_exists():
            self.progress.stop()
            self.window.grab_release()
            self.window.destroy()


class AttachmentsDialog(tk.Toplevel):
    def __init__(self, parent_app: "MarbleDeliveryApp", record_id: str):
        super().__init__(parent_app.root)
        self.parent_app = parent_app
        self.record_id = record_id
        self.items: dict[str, dict[str, Any]] = {}
        record = parent_app._record_by_id(record_id) or {}
        record_values = record.get("values", {})
        self.record_date = DatabaseManager.attachment_record_date(record_id)
        contract = clean_value(record_values.get("contract_number", "")) or "بلا رقم عقد"
        customer = clean_value(record_values.get("customer_name", "")) or "بلا اسم عميل"
        self.title(f"مرفقات السجل - {contract}")
        self.geometry("780x470")
        self.minsize(650, 400)
        self.transient(parent_app.root)
        self.grab_set()
        self.configure(bg=parent_app.settings.background_color)
        toolbar = tk.Frame(self, bg=parent_app.settings.background_color)
        toolbar.pack(fill="x", padx=14, pady=12)
        parent_app.make_button(toolbar, "إضافة ملفات", self._add, parent_app.settings.success_color, width=12, pady=6).pack(side="right", padx=4)
        parent_app.make_button(toolbar, "فتح", self._open, parent_app.settings.primary_color, width=10, pady=6).pack(side="right", padx=4)
        parent_app.make_button(toolbar, "حذف", self._delete, parent_app.settings.danger_color, width=10, pady=6).pack(side="right", padx=4)
        parent_app.make_button(toolbar, "إغلاق", self.destroy, parent_app.settings.muted_color, width=10, pady=6).pack(side="left", padx=4)
        tk.Label(
            self,
            text=f"العميل: {customer}   |   العقد: {contract}   |   تاريخ السجل: {self.record_date}",
            bg=parent_app.settings.surface_color, fg=parent_app.settings.primary_color,
            anchor="e", padx=14, pady=8,
            font=(parent_app.settings.font_family, 11, "bold"),
        ).pack(fill="x", padx=14, pady=(0, 8))
        columns = ("name", "date", "size")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="browse", style="Modern.Treeview")
        self.tree.heading("name", text="اسم الملف")
        self.tree.heading("date", text="تاريخ الإضافة")
        self.tree.heading("size", text="الحجم")
        self.tree.column("name", width=380, anchor="e")
        self.tree.column("date", width=180, anchor="center")
        self.tree.column("size", width=110, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.tree.bind("<Double-1>", lambda _event: self._open())
        self._refresh()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        self.items = {item["id"]: item for item in DatabaseManager.list_attachments(self.record_id)}
        for attachment_id, item in self.items.items():
            path = Path(item["stored_path"])
            size = path.stat().st_size if path.exists() else 0
            size_text = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
            self.tree.insert("", "end", iid=attachment_id, values=(item["original_name"], item["added_at"].replace("T", " "), size_text))

    def _add(self):
        paths = filedialog.askopenfilenames(parent=self, title="اختيار المرفقات", filetypes=[("جميع الملفات", "*.*")])
        if not paths:
            return
        added = 0
        errors = []
        for path in paths:
            try:
                DatabaseManager.add_attachment(self.record_id, Path(path))
                added += 1
            except Exception as exc:
                errors.append(f"{Path(path).name}: {exc}")
        self._refresh()
        self.parent_app.status_var.set(f"تمت إضافة {added} مرفق")
        if errors:
            messagebox.showwarning("بعض المرفقات لم تُضف", "\n".join(errors[:8]), parent=self)

    def _selected(self) -> Optional[dict[str, Any]]:
        selected = self.tree.selection()
        return self.items.get(selected[0]) if selected else None

    def _open(self):
        item = self._selected()
        if not item:
            messagebox.showinfo("المرفقات", "حدد ملفاً أولاً.", parent=self)
            return
        path = Path(item["stored_path"])
        if not path.exists():
            messagebox.showerror("ملف مفقود", "ملف المرفق غير موجود في مجلد البيانات.", parent=self)
            return
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        else:
            webbrowser.open(path.as_uri())

    def _delete(self):
        item = self._selected()
        if not item:
            return
        if messagebox.askyesno("حذف المرفق", f"حذف «{item['original_name']}» نهائياً؟", parent=self):
            DatabaseManager.delete_attachment(item["id"])
            self._refresh()


class ActivityLogDialog(tk.Toplevel):
    def __init__(self, parent_app: "MarbleDeliveryApp"):
        super().__init__(parent_app.root)
        self.title("سجل النشاط والتدقيق")
        self.geometry("980x560")
        self.minsize(760, 460)
        self.transient(parent_app.root)
        self.configure(bg=parent_app.settings.background_color)
        columns = ("date", "actor", "action", "record", "details")
        tree = ttk.Treeview(self, columns=columns, show="headings", style="Modern.Treeview")
        for key, label, width in [
            ("date", "التاريخ والوقت", 170), ("actor", "المصدر", 120),
            ("action", "الإجراء", 150), ("record", "معرّف السجل", 170), ("details", "التفاصيل", 300),
        ]:
            tree.heading(key, text=label)
            tree.column(key, width=width, anchor="center" if key != "details" else "e")
        tree.pack(fill="both", expand=True, padx=14, pady=14)
        for item in DatabaseManager.recent_activity():
            tree.insert("", "end", values=(item["created_at"].replace("T", " "), item["actor"], item["action"], item["record_id"], item["details"]))


class ReportArchiveDialog(tk.Toplevel):
    def __init__(self, parent_app: "MarbleDeliveryApp"):
        super().__init__(parent_app.root)
        self.parent_app = parent_app
        self.title("أرشيف التقارير والبحث")
        self.geometry("1080x620")
        self.minsize(820, 500)
        self.transient(parent_app.root)
        self.configure(bg=parent_app.settings.background_color)
        self.query_var = tk.StringVar()
        self.item_ids: dict[str, str] = {}

        toolbar = tk.Frame(self, bg=parent_app.settings.surface_color)
        toolbar.pack(fill="x", padx=14, pady=(14, 6))
        tk.Label(
            toolbar, text="ابحث باسم العميل أو العقد أو الشركة أو اسم التقرير أو أي قيمة محفوظة:",
            bg=parent_app.settings.surface_color, fg=parent_app.settings.text_color,
            font=(parent_app.settings.font_family, 11, "bold"),
        ).pack(side="right", padx=8)
        search = tk.Entry(
            toolbar, textvariable=self.query_var, justify="right",
            font=(parent_app.settings.font_family, 12), relief="solid", bd=1,
        )
        search.pack(side="right", fill="x", expand=True, ipady=5, padx=6)
        parent_app.make_button(toolbar, "بحث", self._refresh, parent_app.settings.primary_color, width=9).pack(side="right")
        parent_app.make_button(toolbar, "عرض الكل", self._show_all, "#546E7A", width=9).pack(side="left")

        columns = ("date", "format", "file", "title", "recipient", "count", "size")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", style="Modern.Treeview")
        for key, label, width in [
            ("date", "التاريخ والوقت", 170), ("format", "الصيغة", 80),
            ("file", "اسم الملف", 230), ("title", "عنوان التقرير", 190),
            ("recipient", "المستلم", 130), ("count", "السجلات", 80), ("size", "الحجم", 90),
        ]:
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, anchor="center" if key not in {"file", "title"} else "e")
        self.tree.pack(fill="both", expand=True, padx=14, pady=6)
        self.tree.bind("<Double-1>", lambda _event: self._open_selected())
        search.bind("<Return>", lambda _event: self._refresh())

        buttons = tk.Frame(self, bg=parent_app.settings.background_color)
        buttons.pack(fill="x", padx=14, pady=(6, 14))
        parent_app.make_button(buttons, "فتح التقرير", self._open_selected, parent_app.settings.success_color, width=14).pack(side="right", padx=4)
        parent_app.make_button(buttons, "حفظ نسخة", self._save_selected, "#7655A6", width=14).pack(side="right", padx=4)
        tk.Label(
            buttons, text="تُحفظ التقارير كاملة داخل قاعدة بيانات البرنامج.",
            bg=parent_app.settings.background_color, fg=parent_app.settings.muted_color,
            font=(parent_app.settings.font_family, 10),
        ).pack(side="left", padx=6)
        self._refresh()
        search.focus_set()

    @staticmethod
    def _size_text(value: Any) -> str:
        size = max(0, int(value or 0))
        if size >= 1024 * 1024:
            return f"{size / (1024 * 1024):.1f} MB"
        return f"{max(1, math.ceil(size / 1024))} KB"

    def _show_all(self):
        self.query_var.set("")
        self._refresh()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        self.item_ids.clear()
        for item in DatabaseManager.search_report_archive(self.query_var.get()):
            iid = self.tree.insert("", "end", values=(
                clean_value(item.get("created_at", "")).replace("T", " "),
                clean_value(item.get("report_format", "")).upper(),
                item.get("file_name", ""), item.get("report_title", ""),
                item.get("recipient", ""), item.get("record_count", 0),
                self._size_text(item.get("file_size", 0)),
            ))
            self.item_ids[iid] = item["id"]

    def _selected_report(self) -> Optional[dict[str, Any]]:
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("أرشيف التقارير", "حدد تقريرًا أولاً.", parent=self)
            return None
        return DatabaseManager.get_archived_report(self.item_ids.get(selected[0], ""))

    def _open_selected(self):
        item = self._selected_report()
        if not item:
            return
        file_name = safe_filename(clean_value(item.get("file_name", "report.bin")))
        output = PREVIEW_DIR / f"أرشيف_{item['id'][:8]}_{file_name}"
        output.write_bytes(bytes(item["file_data"]))
        open_with_default_app(output)

    def _save_selected(self):
        item = self._selected_report()
        if not item:
            return
        file_name = safe_filename(clean_value(item.get("file_name", "report.bin")))
        suffix = Path(file_name).suffix or f".{clean_value(item.get('report_format', 'bin'))}"
        path = filedialog.asksaveasfilename(
            parent=self, title="حفظ نسخة من التقرير", initialfile=file_name,
            defaultextension=suffix, filetypes=[("ملف التقرير", f"*{suffix}"), ("جميع الملفات", "*.*")],
        )
        if path:
            Path(path).write_bytes(bytes(item["file_data"]))
            messagebox.showinfo("تم", f"تم حفظ النسخة في:\n{path}", parent=self)


class IntegrationCenterDialog(tk.Toplevel):
    """إعداد وتشغيل التكاملات المؤسسية الاختيارية دون فرض مزود محدد."""

    def __init__(self, parent_app):
        super().__init__(parent_app.root)
        self.parent_app = parent_app
        self.settings = parent_app.settings
        self.title("مركز التكامل والاستمرارية")
        self.geometry("820x620")
        self.minsize(760, 560)
        self.transient(parent_app.root)
        self.grab_set()
        self.configure(bg=self.settings.background_color)
        self.status_var = tk.StringVar(value="جاهز - يمكن استخدام التصدير والفحص دون اتصال خارجي")
        self.variables = {
            "central_api_url": tk.StringVar(value=self.settings.central_api_url),
            "erp_api_url": tk.StringVar(value=self.settings.erp_api_url),
            "notification_webhook_url": tk.StringVar(value=self.settings.notification_webhook_url),
            "integration_api_token": tk.StringVar(value=self.settings.integration_api_token),
            "integration_timeout_seconds": tk.IntVar(value=self.settings.integration_timeout_seconds),
        }
        self._build()
        self._refresh_queue_status()

    def _build(self):
        s = self.settings
        header = tk.Frame(self, bg=s.primary_color)
        header.pack(fill="x")
        tk.Label(
            header, text="مركز التكامل والاستمرارية", bg=s.primary_color, fg="white",
            font=(s.font_family, 19, "bold"), pady=15,
        ).pack()
        tk.Label(
            header,
            text="مزامنة مركزية وERP وتنبيهات اختيارية، مع فحص النسخ الاحتياطية قبل الاستعادة",
            bg=s.primary_color, fg="#DDE6EE", font=(s.font_family, 10), pady=4,
        ).pack()

        body = tk.Frame(self, bg=s.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
        body.pack(fill="both", expand=True, padx=16, pady=14)
        body.grid_columnconfigure(0, weight=1)
        labels = [
            ("رابط النظام المركزي API", "central_api_url", False),
            ("رابط تكامل ERP", "erp_api_url", False),
            ("رابط استقبال التنبيهات", "notification_webhook_url", False),
            ("رمز الوصول المشترك", "integration_api_token", True),
            ("مهلة الاتصال بالثواني", "integration_timeout_seconds", False),
        ]
        for row, (label, key, secret) in enumerate(labels):
            tk.Label(
                body, text=label, bg=s.surface_color, fg=s.text_color, anchor="e",
                font=(s.font_family, 10, "bold"),
            ).grid(row=row, column=1, sticky="e", padx=12, pady=8)
            entry = tk.Entry(
                body, textvariable=self.variables[key], justify="left" if "url" in key else "right",
                show="●" if secret else "", relief="solid", bd=1,
                font=("Segoe UI", 10) if "url" in key else (s.font_family, 10),
            )
            entry.grid(row=row, column=0, sticky="ew", padx=12, pady=8, ipady=5)

        self.queue_label = tk.Label(
            body, bg="#F2F5F7", fg=s.primary_color, anchor="e", justify="right",
            font=(s.font_family, 10, "bold"), padx=12, pady=10,
        )
        self.queue_label.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 6))

        buttons = tk.Frame(body, bg=s.surface_color)
        buttons.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=8)
        for col in range(3):
            buttons.grid_columnconfigure(col, weight=1)
        actions = [
            ("حفظ الإعدادات", self._save_settings, s.primary_color),
            ("مزامنة النظام المركزي", self._sync_central, s.success_color),
            ("إرسال حزمة ERP", self._send_erp, "#426B8C"),
            ("إرسال ملخص التنبيهات", self._send_notifications, "#7655A6"),
            ("تصدير حزمة تكامل JSON", self._export_package, "#8A6D2F"),
            ("فحص أحدث نسخة احتياطية", self._verify_latest_backup, s.danger_color),
            ("فتح رسالة واتساب للسجل", self.parent_app.open_selected_whatsapp, "#218C5B"),
        ]
        for index, (text, command, color) in enumerate(actions):
            self.parent_app.make_button(
                buttons, text, command, color, width=21, pady=7,
            ).grid(row=index // 3, column=index % 3, sticky="ew", padx=4, pady=4)

        tk.Label(
            body, textvariable=self.status_var, bg="#EEF3F6", fg=s.text_color,
            wraplength=720, justify="right", anchor="e", padx=12, pady=12,
            font=(s.font_family, 10),
        ).grid(row=7, column=0, columnspan=2, sticky="ew", padx=12, pady=8)

        tk.Label(
            body,
            text="ملاحظة: الروابط الخارجية لا تعمل إلا بعد تزويدها من الجهة التقنية. لا يرسل البرنامج أي بيانات تلقائيًا دون الضغط على زر الإرسال.",
            bg=s.surface_color, fg=s.muted_color, wraplength=720, justify="right",
            font=(s.font_family, 9),
        ).grid(row=8, column=0, columnspan=2, sticky="e", padx=12, pady=(4, 10))

    def _refresh_queue_status(self):
        counts = DatabaseManager.integration_outbox_counts()
        self.queue_label.configure(
            text=(
                f"طابور المزامنة: {counts.get('pending', 0)} بانتظار الإرسال | "
                f"{counts.get('failed', 0)} متعثر | {counts.get('sent', 0)} تم إرساله"
            )
        )

    def _save_settings(self, show_message: bool = True) -> bool:
        for key in ("central_api_url", "erp_api_url", "notification_webhook_url"):
            value = clean_value(self.variables[key].get())
            if value:
                parsed = urllib.parse.urlparse(value)
                if parsed.scheme not in {"http", "https"} or not parsed.netloc:
                    messagebox.showwarning("رابط غير صحيح", f"راجع قيمة: {key}", parent=self)
                    return False
            setattr(self.settings, key, value)
        self.settings.integration_api_token = clean_value(self.variables["integration_api_token"].get())
        try:
            timeout = int(self.variables["integration_timeout_seconds"].get())
        except (TypeError, ValueError, tk.TclError):
            timeout = 20
        self.settings.integration_timeout_seconds = max(5, min(120, timeout))
        self.variables["integration_timeout_seconds"].set(self.settings.integration_timeout_seconds)
        DatabaseManager.save_settings(self.settings)
        if show_message:
            self.status_var.set("تم حفظ إعدادات التكامل محليًا")
        return True

    def _run_async(self, working_text: str, task: Callable[[], str]):
        self.status_var.set(working_text)

        def worker():
            try:
                message = task()
                self.after(0, lambda: self._finish_async(message, ""))
            except Exception as exc:
                logging.exception("فشل إجراء في مركز التكامل")
                self.after(0, lambda text=str(exc): self._finish_async("", text))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_async(self, message: str, error: str):
        self._refresh_queue_status()
        if error:
            self.status_var.set(f"تعذر إكمال العملية: {error}")
            messagebox.showerror("مركز التكامل", error, parent=self)
        else:
            self.status_var.set(message)
            messagebox.showinfo("مركز التكامل", message, parent=self)

    def _sync_central(self):
        if not self._save_settings(False):
            return
        if not self.settings.central_api_url:
            messagebox.showinfo("النظام المركزي", "أدخل رابط النظام المركزي أولًا.", parent=self)
            return
        events = DatabaseManager.pending_integration_events()
        event_ids = [item["id"] for item in events]
        payload = integration_snapshot_payload(
            DatabaseManager.load_records(), self.settings,
            package_type="central_sync", events=events,
        )

        def task():
            try:
                response = post_json_payload(
                    self.settings.central_api_url, payload, self.settings.integration_api_token,
                    self.settings.integration_timeout_seconds,
                )
                DatabaseManager.mark_integration_events(event_ids, sent=True)
                DatabaseManager.log_action(
                    "مزامنة مركزية", details=f"السجلات: {payload['record_count']} | الأحداث: {len(events)}"
                )
                return f"اكتملت المزامنة المركزية: {payload['record_count']} سجل و{len(events)} تغيير. استجابة الخدمة: {clean_value(response.get('message', 'ناجحة')) or 'ناجحة'}"
            except Exception as exc:
                DatabaseManager.mark_integration_events(event_ids, sent=False, error=str(exc))
                raise

        self._run_async("جاري إرسال السجلات والتغييرات إلى النظام المركزي...", task)

    def _send_erp(self):
        if not self._save_settings(False):
            return
        if not self.settings.erp_api_url:
            messagebox.showinfo("تكامل ERP", "أدخل رابط تكامل ERP أولًا.", parent=self)
            return
        payload = integration_snapshot_payload(
            DatabaseManager.load_records(), self.settings, package_type="erp_snapshot"
        )
        self._run_async(
            "جاري إرسال حزمة ERP...",
            lambda: (
                post_json_payload(
                    self.settings.erp_api_url, payload, self.settings.integration_api_token,
                    self.settings.integration_timeout_seconds,
                ),
                DatabaseManager.log_action("إرسال ERP", details=f"السجلات: {payload['record_count']}"),
                f"تم إرسال {payload['record_count']} سجل إلى رابط ERP بنجاح.",
            )[-1],
        )

    def _send_notifications(self):
        if not self._save_settings(False):
            return
        if not self.settings.notification_webhook_url:
            messagebox.showinfo("التنبيهات", "أدخل رابط استقبال التنبيهات أولًا.", parent=self)
            return
        payload = notification_summary_payload(DatabaseManager.load_records(), self.settings)
        self._run_async(
            "جاري إرسال ملخص التسليمات والتنبيهات...",
            lambda: (
                post_json_payload(
                    self.settings.notification_webhook_url, payload, self.settings.integration_api_token,
                    self.settings.integration_timeout_seconds,
                ),
                DatabaseManager.log_action("إرسال ملخص تنبيهات", details=json.dumps(payload["counts"], ensure_ascii=False)),
                "تم إرسال ملخص التنبيهات بنجاح: "
                f"{payload['counts']['overdue']} متأخر، {payload['counts']['due_today']} مستحق اليوم، "
                f"{payload['counts']['installed_today']} تركيب اليوم.",
            )[-1],
        )

    def _export_package(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="حفظ حزمة التكامل",
            initialfile=f"حزمة_تكامل_الرخام_{dt.datetime.now():%Y%m%d_%H%M}.json",
            defaultextension=".json", filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        payload = integration_snapshot_payload(
            DatabaseManager.load_records(), self.settings,
            package_type="portable_integration_snapshot",
            events=DatabaseManager.pending_integration_events(),
        )
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        DatabaseManager.log_action("تصدير حزمة تكامل", details=Path(path).name)
        self.status_var.set(f"تم تصدير حزمة تضم {payload['record_count']} سجل")

    def _verify_latest_backup(self):
        candidates = sorted(BACKUP_DIR.glob("*.zip"), key=lambda item: item.stat().st_mtime, reverse=True)
        if candidates:
            path = candidates[0]
        else:
            selected = filedialog.askopenfilename(
                parent=self, title="اختيار نسخة لفحصها",
                filetypes=[("نسخة ZIP", "*.zip"), ("JSON", "*.json")],
            )
            if not selected:
                return
            path = Path(selected)

        def task():
            result = validate_backup_archive(path)
            DatabaseManager.log_action("فحص نسخة احتياطية", details=f"{path.name} | سليمة")
            suffix = " (نسخة قديمة بلا دليل بصمات)" if result["legacy_without_manifest"] else ""
            return (
                f"النسخة سليمة{suffix}. السجلات: {result['records']}، "
                f"المرفقات: {result['attachments']}، التقارير: {result['reports']}، "
                f"الملفات المتحقق منها: {result['manifest_files']}."
            )

        self._run_async(f"جاري فحص {path.name} دون استعادتها...", task)


class OperationsDashboardDialog(tk.Toplevel):
    def __init__(self, parent_app: "MarbleDeliveryApp"):
        super().__init__(parent_app.root)
        self.parent_app = parent_app
        self.settings = parent_app.settings
        self.filtered_records: list[dict[str, Any]] = []
        self.title("لوحة تحليلات المبيعات والتسليم")
        self.geometry("1180x760")
        self.minsize(980, 620)
        self.transient(parent_app.root)
        self.configure(bg=parent_app.settings.background_color)
        header = tk.Frame(self, bg=self.settings.primary_color)
        header.pack(fill="x")
        tk.Label(
            header, text="تحليلات المبيعات والتسليم", bg=self.settings.primary_color,
            fg="white", font=(self.settings.font_family, 20, "bold"),
        ).pack(side="right", padx=18, pady=13)
        tk.Label(
            header, text="المؤشرات محسوبة بعدد سجلات التسليم في الفترة المختارة",
            bg=self.settings.primary_color, fg="#DDE6EE",
            font=(self.settings.font_family, 10),
        ).pack(side="right", padx=6)

        controls = tk.Frame(self, bg=self.settings.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
        controls.pack(fill="x", padx=14, pady=10)
        period_row = tk.Frame(controls, bg=self.settings.surface_color)
        period_row.pack(fill="x", padx=8, pady=(7, 3))
        self.period_var = tk.StringVar(value="هذا الشهر")
        tk.Label(period_row, text="الفترة:", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, 10, "bold")).pack(side="right", padx=(8, 4), pady=8)
        period_combo = ttk.Combobox(
            period_row, textvariable=self.period_var,
            values=("هذا الأسبوع", "هذا الشهر", "هذه السنة", "كل الفترات", "حسب الطلب"),
            state="readonly", width=14,
        )
        period_combo.pack(side="right", padx=4)
        period_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh())

        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        for label, variable in (("من", self.from_var), ("إلى", self.to_var)):
            tk.Label(period_row, text=label, bg=self.settings.surface_color, fg=self.settings.text_color).pack(side="right", padx=(8, 2))
            tk.Entry(period_row, textvariable=variable, justify="center", width=11).pack(side="right", ipady=3)
            tk.Button(
                period_row, text="تقويم", command=lambda var=variable, title=label: self._pick_period_date(var, title),
                bg="#E8EEF2", fg=self.settings.primary_color, relief="flat", cursor="hand2",
                font=(self.settings.font_family, 8, "bold"),
            ).pack(side="right", padx=2)

        field_labels = [item.label for item in self.settings.visible_fields]
        marble_label = next(
            (item.label for item in self.settings.visible_fields if item.key == "marble_color"),
            field_labels[0] if field_labels else "",
        )
        self.date_basis_var = tk.StringVar(value="التاريخ الأنسب")
        tk.Label(period_row, text="أساس التاريخ:", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, 10, "bold")).pack(side="right", padx=(14, 4))
        date_basis_combo = ttk.Combobox(
            period_row, textvariable=self.date_basis_var,
            values=("التاريخ الأنسب", "التركيب الفعلي", "تاريخ إنشاء السجل", "تاريخ آخر تعديل"),
            state="readonly", width=18,
        )
        date_basis_combo.pack(side="right", padx=4)
        date_basis_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh())

        action_row = tk.Frame(controls, bg=self.settings.surface_color)
        action_row.pack(fill="x", padx=8, pady=(3, 7))
        self.analysis_field_var = tk.StringVar(value=marble_label)
        tk.Label(action_row, text="حلّل النتائج حسب أي حقل:", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, 10, "bold")).pack(side="right", padx=(8, 4))
        analysis_combo = ttk.Combobox(
            action_row, textvariable=self.analysis_field_var, values=field_labels,
            state="readonly", width=24,
        )
        analysis_combo.pack(side="right", padx=4)
        analysis_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh())

        parent_app.make_button(action_row, "تحديث", self._refresh, self.settings.primary_color, width=8, pady=5).pack(side="left", padx=4)
        parent_app.make_button(action_row, "PDF للفترة", lambda: self._export_period("pdf"), self.settings.danger_color, width=10, pady=5).pack(side="left", padx=4)
        parent_app.make_button(action_row, "Excel للفترة", lambda: self._export_period("xlsx"), self.settings.success_color, width=10, pady=5).pack(side="left", padx=4)

        self.period_label_var = tk.StringVar()
        tk.Label(
            self, textvariable=self.period_label_var, bg=self.settings.background_color,
            fg=self.settings.muted_color, anchor="e", font=(self.settings.font_family, 10, "bold"),
        ).pack(fill="x", padx=18)
        self.summary = tk.Frame(self, bg=self.settings.background_color)
        self.summary.pack(fill="x", padx=14, pady=(6, 10))
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self._refresh()

    def _record_date(self, record: dict[str, Any]) -> Optional[dt.date]:
        return record_date_for_basis(record, self.date_basis_var.get())

    def _period_bounds(self) -> tuple[Optional[dt.date], Optional[dt.date], str]:
        today = dt.date.today()
        period = self.period_var.get()
        if period == "هذا الأسبوع":
            start = today - dt.timedelta(days=today.weekday())
            end = start + dt.timedelta(days=6)
        elif period == "هذا الشهر":
            start = today.replace(day=1)
            end = (start.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
        elif period == "هذه السنة":
            start, end = today.replace(month=1, day=1), today.replace(month=12, day=31)
        elif period == "حسب الطلب":
            try:
                start = dt.datetime.strptime(self.from_var.get(), "%Y-%m-%d").date() if self.from_var.get() else None
                end = dt.datetime.strptime(self.to_var.get(), "%Y-%m-%d").date() if self.to_var.get() else None
            except ValueError:
                raise ValueError("استخدم صيغة التاريخ YYYY-MM-DD في النطاق المخصص.")
            if start and end and start > end:
                raise ValueError("تاريخ البداية يجب ألا يكون بعد تاريخ النهاية.")
        else:
            start = end = None
        if start or end:
            label = f"الفترة: {start.isoformat() if start else 'البداية'} إلى {end.isoformat() if end else 'اليوم'}"
        else:
            label = "الفترة: جميع السجلات"
        return start, end, label

    def _pick_period_date(self, variable: tk.StringVar, label: str):
        picker = DatePickerDialog(self, variable.get(), f"اختيار تاريخ {label}", self.settings)
        self.wait_window(picker)
        if picker.result:
            variable.set(picker.result)
            self.period_var.set("حسب الطلب")
            self._refresh()

    @staticmethod
    def _counter(records: list[dict[str, Any]], key: str) -> Counter:
        return Counter(
            clean_value(record.get("values", {}).get(key, "")) or "غير محدد"
            for record in records
        )

    @staticmethod
    def _top_text(counts: Counter) -> str:
        return counts.most_common(1)[0][0] if counts else "لا توجد بيانات"

    def _add_card(self, label: str, value: str, color: str):
        card = tk.Frame(self.summary, bg=self.settings.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
        card.pack(side="right", fill="x", expand=True, padx=3, ipady=4)
        tk.Label(card, text=value, bg=self.settings.surface_color, fg=color,
                 wraplength=150, font=(self.settings.font_family, 14, "bold")).pack(padx=4, pady=(3, 0))
        tk.Label(card, text=label, bg=self.settings.surface_color, fg=self.settings.muted_color,
                 font=(self.settings.font_family, 9)).pack(pady=(0, 3))

    def _add_counter_tab(self, field: FieldDefinition, records: list[dict[str, Any]]):
        counts = self._counter(records, field.key)
        frame = tk.Frame(self.notebook, bg=self.settings.surface_color)
        self.notebook.add(frame, text=field.label)
        tree = ttk.Treeview(
            frame, columns=("rank", "name", "count", "share", "bar"),
            show="headings", style="Modern.Treeview",
        )
        for key, label, width, anchor in [
            ("rank", "الترتيب", 80, "center"), ("name", field.label, 300, "e"),
            ("count", "عدد السجلات", 120, "center"), ("share", "النسبة", 100, "center"),
            ("bar", "مؤشر", 260, "w"),
        ]:
            tree.heading(key, text=label)
            tree.column(key, width=width, anchor=anchor)
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        total = max(1, sum(counts.values()))
        maximum = max(counts.values(), default=1)
        for rank, (name, count) in enumerate(counts.most_common(), 1):
            bar = "█" * max(1, round(count / maximum * 24))
            tree.insert("", "end", values=(rank, name, count, f"{count / total:.1%}", bar))
        tree.bind("<Double-1>", lambda _event, target=tree, item=field: self._apply_result_filter(target, item))

    def _apply_result_filter(self, tree: ttk.Treeview, field: FieldDefinition):
        selected = tree.selection()
        if not selected:
            return
        value = clean_value(tree.item(selected[0], "values")[1])
        self.parent_app.filter_field_var.set(field.label)
        self.parent_app._on_filter_field_changed()
        self.parent_app.filter_value_var.set("" if value == "غير محدد" else value)
        self.parent_app._apply_filters()
        self.parent_app.status_var.set(f"تم تطبيق البحث حسب {field.label}: {value}")

    def _refresh(self):
        try:
            start, end, label = self._period_bounds()
        except ValueError as exc:
            messagebox.showwarning("الفترة", str(exc), parent=self)
            return
        self.filtered_records = []
        for record in self.parent_app.records:
            record_date = self._record_date(record)
            if start and (record_date is None or record_date < start):
                continue
            if end and (record_date is None or record_date > end):
                continue
            self.filtered_records.append(record)
        self.period_label_var.set(f"{label} | عدد السجلات: {len(self.filtered_records)}")
        for child in self.summary.winfo_children():
            child.destroy()
        status_counts = self._counter(self.filtered_records, "delivery_status")
        color_counts = self._counter(self.filtered_records, "marble_color")
        employee_counts = self._counter(self.filtered_records, "salesperson")
        showroom_counts = self._counter(self.filtered_records, "showroom")
        sink_counts = self._counter(self.filtered_records, "sink_type")
        self._add_card("إجمالي الفترة", str(len(self.filtered_records)), self.settings.primary_color)
        self._add_card("تم التسليم", str(status_counts["تم التسليم"]), self.settings.success_color)
        self._add_card("اللون الأعلى", self._top_text(color_counts), "#8A6D2F")
        self._add_card("الموظف الأعلى", self._top_text(employee_counts), "#7655A6")
        self._add_card("الفرع الأعلى", self._top_text(showroom_counts), "#2E6F9E")
        self._add_card("المجلى الأعلى", self._top_text(sink_counts), "#426B8C")

        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)
        preferred_keys = ["marble_color", "salesperson", "showroom", "sink_type", "marble_edge", "delivery_status"]
        chosen_label = self.analysis_field_var.get()
        chosen_key = next((item.key for item in self.settings.visible_fields if item.label == chosen_label), "")
        if chosen_key:
            preferred_keys = [chosen_key, *[key for key in preferred_keys if key != chosen_key]]
        by_key = {item.key: item for item in self.settings.visible_fields}
        for key in preferred_keys:
            field = by_key.get(key)
            if field:
                self._add_counter_tab(field, self.filtered_records)
        if self.notebook.tabs():
            self.notebook.select(self.notebook.tabs()[0])

    def _export_period(self, file_type: str):
        self._refresh()
        if not self.filtered_records:
            messagebox.showinfo("التقارير", "لا توجد سجلات في الفترة المختارة.", parent=self)
            return
        self.parent_app._export(
            file_type, records=self.filtered_records,
            recipient=self.parent_app.recipient_var.get(),
        )

# ============================================================================
# نافذة الربط اليدوي للأعمدة
# ============================================================================
class ManualMappingDialog(tk.Toplevel):
    def __init__(self, parent: "MarbleDeliveryApp", headers: list[str], fields: list[FieldDefinition],
                 suggested: Optional[dict[int, str]] = None, sample_rows: Optional[list[list[str]]] = None):
        super().__init__(parent.root)
        self.parent_app = parent
        self.title("الربط اليدوي للأعمدة")
        self.geometry("760x600")
        self.minsize(680, 480)
        self.transient(parent.root)
        self.grab_set()
        self.configure(bg=parent.settings.surface_color)
        self.mapping: dict[int, str] = {}
        self.vars: dict[int, tk.StringVar] = {}
        self.fields_by_label = {item.label: item.key for item in fields}
        self.fields_by_key = {item.key: item.label for item in fields}
        suggested = suggested or {}
        sample_rows = sample_rows or []
        tk.Label(
            self,
            text="راجع ربط أعمدة الملف بحقول البرنامج. يمكنك تجاهل أي عمود غير مطلوب.",
            bg=parent.settings.surface_color, fg=parent.settings.primary_color,
            font=(parent.settings.font_family, 13, "bold"),
        ).pack(pady=(14, 4))
        tk.Label(
            self, text="تم وضع اقتراحات تلقائية ويمكن تعديلها قبل الاستيراد.",
            bg=parent.settings.surface_color, fg=parent.settings.muted_color,
            font=(parent.settings.font_family, 10),
        ).pack(pady=(0, 10))
        frame = tk.Frame(self, bg=parent.settings.surface_color)
        frame.pack(fill="both", expand=True, padx=20)
        canvas = tk.Canvas(frame, bg=parent.settings.surface_color)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=parent.settings.surface_color)
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        scrollable.grid_columnconfigure(0, weight=1)
        scrollable.grid_columnconfigure(1, weight=1)
        for i, header in enumerate(headers):
            if not clean_value(header) and not any(i < len(row) and clean_value(row[i]) for row in sample_rows[:3]):
                continue
            sample = next((clean_value(row[i]) for row in sample_rows[:5] if i < len(row) and clean_value(row[i])), "")
            source_text = clean_value(header) or f"عمود {i + 1} بلا عنوان"
            if sample:
                source_text += f"\nمثال: {sample[:55]}"
            tk.Label(
                scrollable, text=source_text, justify="right", anchor="e",
                bg=parent.settings.surface_color, fg=parent.settings.text_color,
                font=(parent.settings.font_family, 10, "bold"),
            ).grid(row=i, column=1, sticky="ew", padx=10, pady=7)
            default_label = self.fields_by_key.get(suggested.get(i, ""), "(تجاهل)")
            var = tk.StringVar(value=default_label)
            self.vars[i] = var
            ttk.Combobox(
                scrollable, textvariable=var, values=["(تجاهل)"] + [f.label for f in fields],
                state="readonly", width=28,
            ).grid(row=i, column=0, sticky="ew", padx=10, pady=7)

        btn_frame = tk.Frame(self, bg=parent.settings.surface_color)
        btn_frame.pack(fill="x", pady=10)
        tk.Button(btn_frame, text="تأكيد الربط", command=self._save,
                  bg=parent.settings.success_color, fg="white", relief="flat",
                  cursor="hand2", font=(parent.settings.font_family, 10, "bold"), padx=20, pady=5).pack(side="right", padx=6)
        tk.Button(btn_frame, text="إلغاء", command=self.destroy,
                  bg=parent.settings.danger_color, fg="white", relief="flat",
                  cursor="hand2", font=(parent.settings.font_family, 10, "bold"), padx=20, pady=5).pack(side="right")
        self.wait_window()

    def _save(self):
        used: set[str] = set()
        for column_index, var in self.vars.items():
            label = var.get()
            if label == "(تجاهل)":
                continue
            key = self.fields_by_label.get(label)
            if not key:
                continue
            if key in used:
                messagebox.showwarning("ربط مكرر", f"تم ربط الحقل «{label}» بأكثر من عمود.", parent=self)
                return
            self.mapping[column_index] = key
            used.add(key)
        if not self.mapping:
            messagebox.showwarning("لا يوجد ربط", "اختر حقلًا واحدًا على الأقل للاستيراد.", parent=self)
            return
        self.destroy()

# ============================================================================
# نافذة استيراد الملفات القديمة
# ============================================================================
class OldImportDialog(tk.Toplevel):
    def __init__(self, parent: "MarbleDeliveryApp", file_path: Path):
        super().__init__(parent.root)
        self.parent_app = parent
        self.file_path = file_path
        self.title("استيراد بيانات من نسخة قديمة")
        self.geometry("700x500")
        self.transient(parent.root)
        self.grab_set()
        self.configure(bg=parent.settings.background_color)
        self.settings = parent.settings
        self.old_data = None
        self.vars: dict[str, tk.StringVar] = {}
        self._load_data()
        self._build()
        self.wait_visibility()
        self.focus_force()

    def _load_data(self):
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self.old_data = json.load(f)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل قراءة الملف: {e}", parent=self)
            self.destroy()

    def _build(self):
        main = tk.Frame(self, bg=self.settings.background_color)
        main.pack(fill="both", expand=True, padx=20, pady=20)
        tk.Label(main, text="حدد كيفية ربط الحقول القديمة (تم اقتراح المطابقات تلقائياً بالذكاء الاصطناعي):",
                 bg=self.settings.background_color, font=(self.settings.font_family, 11, "bold")).pack(anchor="w", pady=(0, 10))
        old_records = self.old_data.get("records", [])
        if not old_records:
            messagebox.showwarning("تنبيه", "لا توجد سجلات.", parent=self)
            self.destroy()
            return
        first_record = old_records[0] if isinstance(old_records[0], dict) else {}
        first_values = first_record.get("values", first_record)
        old_keys = list(first_values.keys()) if isinstance(first_values, dict) else []
        if not old_keys:
            messagebox.showwarning("تنبيه", "تعذر اكتشاف حقول في النسخة القديمة.", parent=self)
            self.destroy()
            return
        smart_map = get_smart_mapping(old_keys, self.settings.visible_fields)
        field_labels = {item.key: item.label for item in self.settings.fields}

        frame = tk.Frame(main, bg=self.settings.surface_color, relief="solid", bd=1)
        frame.pack(fill="both", expand=True, pady=10)
        canvas = tk.Canvas(frame, bg=self.settings.surface_color)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=self.settings.surface_color)
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for i, old_key in enumerate(old_keys):
            tk.Label(scrollable, text=f"القديم: {old_key}", bg=self.settings.surface_color,
                     font=(self.settings.font_family, 10, "bold")).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            var = tk.StringVar(value=field_labels.get(smart_map.get(old_key, ""), "(تجاهل)"))
            self.vars[old_key] = var
            ttk.Combobox(scrollable, textvariable=var, values=["(تجاهل)"] + [f.label for f in self.settings.fields],
                         state="readonly", width=20).grid(row=i, column=1, sticky="w", padx=5, pady=5)

        btn_frame = tk.Frame(main, bg=self.settings.background_color)
        btn_frame.pack(fill="x", pady=10)
        tk.Button(btn_frame, text="استيراد", command=self._import,
                  bg=self.settings.success_color, fg="white", relief="flat",
                  cursor="hand2", font=(self.settings.font_family, 10, "bold"), padx=20, pady=5).pack(side="right", padx=5)
        tk.Button(btn_frame, text="إلغاء", command=self.destroy,
                  bg=self.settings.danger_color, fg="white", relief="flat",
                  cursor="hand2", font=(self.settings.font_family, 10, "bold"), padx=20, pady=5).pack(side="right", padx=5)

    def _import(self):
        fields_by_label = {item.label: item.key for item in self.settings.fields}
        mapping = {
            old_key: fields_by_label[var.get()]
            for old_key, var in self.vars.items()
            if var.get() in fields_by_label
        }
        if not mapping:
            messagebox.showwarning("تنبيه", "لم تختر أي حقل.", parent=self)
            return
        old_records = self.old_data.get("records", [])
        new_records = []
        progress = ProgressDialog(self.parent_app, "جاري استيراد النسخة القديمة...")
        for idx, old_rec in enumerate(old_records):
            if idx % 50 == 0:
                progress.update_text(f"تم استيراد {idx} من {len(old_records)}")
            if not isinstance(old_rec, dict):
                continue
            old_vals = old_rec.get("values", old_rec)
            if not isinstance(old_vals, dict):
                continue
            new_vals = {new_k: old_vals.get(old_k, "") for old_k, new_k in mapping.items()}
            ensure_kitchen_identity(new_vals, [*self.parent_app.records, *new_records])
            now = dt.datetime.now().isoformat(timespec="seconds")
            new_records.append({
                "id": uuid.uuid4().hex,
                "created_at": old_rec.get("created_at", now),
                "updated_at": now,
                "values": new_vals
            })
        progress.close()
        DatabaseManager.save_records_bulk(new_records)
        self.parent_app.records.extend(new_records)
        if self.parent_app.current_report_record_ids is not None:
            self.parent_app.current_report_record_ids.update(
                str(record["id"]) for record in new_records
            )
        self.parent_app._apply_filters()
        messagebox.showinfo(
            "اكتمل",
            f"تم استيراد {len(new_records)} سجل.\n"
            "تم قبول العقود المتكررة كسجلات مستقلة مع تمييزها داخليًا.",
            parent=self,
        )
        self.destroy()

# ============================================================================
# فئة Exporters المحسنة
# ============================================================================
class ReportExporter:
    def __init__(self, settings: AppSettings):
        self.settings = settings
        self.regular_font_path, self.bold_font_path = resolve_font_paths(settings.font_family)

    @staticmethod
    def _values(record: dict[str, Any], fields: list[FieldDefinition]) -> list[str]:
        values = record.get("values", {})
        return [clean_value(values.get(item.key, "")) for item in fields]

    def _palette(self) -> dict[str, str]:
        if self.settings.print_friendly_reports:
            return {
                "title_fill": "#E8EEF2",
                "title_text": "#17212B",
                "subtitle_fill": "#E8EEF2",
                "subtitle_text": "#3E4C59",
                "header_fill": "#D6E0E7",
                "header_text": "#17212B",
                "meta_fill": "#F2F5F7",
                "row_alt": "#F7F9FA",
                "border": "#AAB7C2",
                "accent": "#7B8792",
                "page_bg": "#FFFFFF",
                "text": "#1D2730",
                "muted": "#55616C",
            }
        return {
            "title_fill": self.settings.title_bg_color,
            "title_text": self.settings.title_fg_color,
            "subtitle_fill": self.settings.subtitle_bg_color,
            "subtitle_text": self.settings.subtitle_fg_color,
            "header_fill": self.settings.primary_color,
            "header_text": "#FFFFFF",
            "meta_fill": "#F1F3F5",
            "row_alt": "#F8F9FA",
            "border": "#CED4DA",
            "accent": self.settings.accent_color,
            "page_bg": self.settings.background_color,
            "text": self.settings.text_color,
            "muted": self.settings.muted_color,
        }

    # --------------------- HTML (تفاعلي مع رسوم بيانية) ---------------------
    def export_html(self, records: list[dict[str, Any]], recipient: str, output: Path,
                    selected_fields: Optional[List[FieldDefinition]] = None) -> None:
        fields = [f for f in (selected_fields or self.settings.report_fields) if f.key != "scheduled_delivery_date"]
        s = self.settings
        now = dt.datetime.now()
        metrics = report_record_metrics(records)

        def html_cell(record: dict[str, Any], field: FieldDefinition) -> str:
            value = clean_value(record.get("values", {}).get(field.key, ""))
            status_classes = {
                "جاهز للتسليم": "status-ready",
                "تم التسليم": "status-delivered",
                "تعديل - يعتمد الحالي ويلغى السابق": "status-revision",
                "مؤجل": "status-postponed",
                "ملغي": "status-cancelled",
            }
            css_class = status_classes.get(value, "") if field.key == "delivery_status" else ""
            class_attr = f' class="{css_class}"' if css_class else ""
            return f"<td{class_attr}>{html_escape(value)}</td>"

        # تحليل البيانات لإنشاء مخططات CSS
        analytics = {}
        for f in fields:
            if f.key in ["showroom", "marble_color", "salesperson"]:
                counts = {}
                for r in records:
                    val = clean_value(r.get("values", {}).get(f.key, "غير محدد")) or "غير محدد"
                    counts[val] = counts.get(val, 0) + 1
                if counts:
                    analytics[f.label] = counts

        analytics_html = ""
        for title, counts in analytics.items():
            max_val = max(counts.values()) if counts else 1
            bars = "".join([
                f'<div class="bar"><div class="bar-fill" style="width: {(v/max_val)*100:.2f}%">{html_escape(k)} ({v})</div></div>'
                for k, v in sorted(counts.items(), key=lambda x: x[1], reverse=True)
            ])
            analytics_html += f'<div class="chart-card"><h3>{html_escape(title)}</h3>{bars}</div>'

        # تقسيم HTML للطباعة بتقدير أكبر من السابق؛ PDF/PNG يستخدمان قياساً فعلياً ديناميكياً للصفحة.
        print_row_target = report_rows_per_page(fields)
        print_chunks = (
            [records[index:index + print_row_target] for index in range(0, len(records), print_row_target)]
            if records else [[]]
        )
        print_weights = report_column_weights(fields, records)
        print_weight_total = sum(print_weights) or 1.0
        print_colgroup = "<colgroup>" + "".join(
            f'<col style="width:{weight / print_weight_total * 100:.3f}%">'
            for weight in print_weights
        ) + "</colgroup>"
        print_sections = "".join(
            '<section class="print-section">'
            f'<div class="print-page-head"><b>{html_escape(s.report_title)}</b>'
            f'<span>{html_escape(s.company_name)} | {html_escape(s.report_subtitle)}</span></div>'
            f'<div class="print-page-meta">صفحة {page_index + 1} من {len(print_chunks)} | '
            f'السجلات: {metrics["records"]} | العقود: {metrics["contracts"]} | '
            f'العقود المتكررة: {metrics["multi_kitchen_contracts"]} | سجلات الصفحة: {len(page_records)}</div>'
            '<table class="print-table">' + print_colgroup + '<thead><tr>'
            + "".join(f"<th>{html_escape(field.label)}</th>" for field in fields)
            + "</tr></thead><tbody>"
            + "".join(
                "<tr>" + "".join(
                    html_cell(record, field) for field in fields
                ) + "</tr>" for record in page_records
            )
            + "</tbody></table>"
            f'<div class="print-page-foot"><span>المستلم: {html_escape(recipient) or "...................."}</span>'
            '<span>التوقيع: ....................</span>'
            f'<span>التاريخ: {now:%Y-%m-%d}</span></div></section>'
            for page_index, page_records in enumerate(print_chunks)
        )

        html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html_escape(s.report_title)} - {html_escape(s.report_title_en)}</title>
<style>
:root {{ --primary: {s.primary_color}; --accent: {s.accent_color}; --bg: {s.background_color}; --surface: #ffffff; --text: #1a1a2e; --muted: #6c757d; --border: #ccd5dd; --title-bg: {s.title_bg_color}; --title-fg: {s.title_fg_color}; }}
body {{ font-family: '{s.font_family}', 'Segoe UI', Tahoma, Arial, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 20px; }}
.container {{ max-width: 1400px; margin: auto; background: var(--surface); border-radius: 16px; box-shadow: 0 10px 30px rgba(0,0,0,0.05); overflow: hidden; }}
.header {{ background: linear-gradient(135deg, var(--title-bg), #eef5ef); color: var(--title-fg); padding: 40px; text-align: center; border-bottom: 5px solid var(--accent); }}
.header h1 {{ margin: 0; font-size: 2.5rem; }}
.header .en-title {{ margin-top: 7px; font: 700 1.35rem 'Segoe UI', Arial, sans-serif; letter-spacing: 0.4px; direction: ltr; }}
.header p {{ margin: 10px 0 0; opacity: 0.9; font-size: 1.2rem; }}
.content {{ padding: 30px; }}
.search-box {{ width: 100%; padding: 12px 20px; font-size: 1rem; border: 2px solid var(--border); border-radius: 8px; margin-bottom: 20px; outline: none; transition: 0.3s; box-sizing: border-box; }}
.search-box:focus {{ border-color: var(--accent); box-shadow: 0 0 0 3px rgba(201, 168, 76, 0.2); }}
table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 0.95rem; table-layout: fixed; }}
th {{ background: #d6e0e7; color: #17212b; padding: 15px; text-align: center; position: sticky; top: 0; border: 1px solid #aab7c2; overflow-wrap: anywhere; word-break: break-word; }}
td {{ padding: 12px; text-align: center; border: 1px solid #c7d0d8; overflow-wrap: anywhere; word-break: break-word; }}
tr:nth-child(even) {{ background: #f8f9fa; }}
tr:hover {{ background: #e9ecef; transition: 0.2s; }}
.chart-container {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }}
.chart-card {{ background: #f8f9fa; padding: 20px; border-radius: 12px; border-right: 4px solid var(--accent); }}
.chart-card h3 {{ margin: 0 0 15px; color: var(--primary); }}
.bar {{ background: #e0e0e0; border-radius: 4px; margin: 8px 0; position: relative; overflow: hidden; }}
.bar-fill {{ background: var(--accent); height: 24px; border-radius: 4px; display: flex; align-items: center; justify-content: flex-end; padding-right: 10px; color: white; font-size: 0.8rem; font-weight: bold; transition: width 0.5s ease; }}
.footer {{ text-align: center; padding: 20px; color: var(--muted); border-top: 1px solid var(--border); margin-top: 30px; }}
.print-only {{ display: none; }}
.status-ready {{ background: #dceeff !important; font-weight: 700; }}
.status-delivered {{ background: #ddf1e5 !important; font-weight: 700; }}
.status-revision {{ background: #ffe8a3 !important; font-weight: 700; }}
.status-postponed {{ background: #fff1cc !important; font-weight: 700; }}
.status-cancelled {{ background: #fde3e3 !important; font-weight: 700; }}
@page {{ size: A4 landscape; margin: 3mm 2.5mm; }}
@media print {{
  .search-box, .chart-container, .screen-report, .screen-only, .container > .header, .container > .footer {{ display: none !important; }}
  .print-only {{ display: block; }}
  body {{ background: white; padding: 0; }}
  .container {{ width: 100%; max-width: none; box-shadow: none; border-radius: 0; }}
  .content {{ padding: 0; }}
  .print-section {{ margin: 0; break-after: page; page-break-after: always; }}
  .print-section:last-child {{ break-after: auto; page-break-after: auto; }}
  .print-page-head {{ display: flex; justify-content: space-between; align-items: center; background: #eef3f5; border-bottom: 3pt solid var(--accent); padding: 7pt 9pt; margin-bottom: 5pt; }}
  .print-page-head b {{ font-size: 15pt; color: #17212b; }}
  .print-page-head span, .print-page-meta {{ font-size: 8pt; color: #4d5963; }}
  .print-page-meta {{ text-align: center; background: #f5f6f7; border: 0.5pt solid #c7d0d8; padding: 3pt; margin-bottom: 4pt; }}
  .print-table {{ width: 100%; font-size: 8.7pt; table-layout: fixed; border-collapse: collapse; }}
  .print-table th {{ background: #ddd; color: #111; }}
  .print-table th, .print-table td {{ padding: 5pt 2pt; min-height: 34pt; line-height: 1.24; overflow-wrap: anywhere; word-break: break-word; vertical-align: middle; text-align: center; }}
  .print-table thead {{ display: table-header-group; }}
  .print-table tr {{ break-inside: avoid; page-break-inside: avoid; }}
  .print-page-foot {{ display: flex; justify-content: space-between; border-top: 1pt solid var(--accent); margin-top: 6pt; padding-top: 4pt; font-size: 8pt; }}
}}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>{html_escape(s.report_title)}</h1>
        <div class="en-title">{html_escape(s.report_title_en)}</div>
        <p>{html_escape(s.company_name)} | {html_escape(s.report_subtitle)}</p>
    </div>
    <div class="content">
        <div class="screen-only" style="display: flex; justify-content: space-between; background: #f1f3f5; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <span><b>السجلات:</b> {metrics['records']}</span>
            <span><b>العقود:</b> {metrics['contracts']}</span>
            <span><b>العقود المتكررة:</b> {metrics['multi_kitchen_contracts']}</span>
            <span><b>التاريخ:</b> {now:%Y-%m-%d}</span>
        </div>
        {'<div class="chart-container">' + analytics_html + '</div>' if analytics_html else ''}
        <input type="text" class="search-box" id="searchInput" onkeyup="filterTable()" placeholder="🔍 بحث حي في الجدول...">
        <table id="dataTable" class="screen-report">
            <thead><tr>{"".join(f"<th>{html_escape(f.label)}</th>" for f in fields)}</tr></thead>
            <tbody>
                {"".join("<tr>" + "".join(html_cell(r, f) for f in fields) + "</tr>" for r in records)}
            </tbody>
        </table>
        <div class="print-only">{print_sections}</div>
        <div class="screen-only" style="margin-top: 40px; display: flex; justify-content: space-around; border-top: 2px solid var(--border); padding-top: 20px;">
            <div style="text-align: center;"><b>اسم المستلم:</b><br>{html_escape(recipient) or '....................'}</div>
            <div style="text-align: center;"><b>التوقيع:</b><br>....................</div>
            <div style="text-align: center;"><b>التاريخ:</b><br>{now:%Y-%m-%d}</div>
        </div>
    </div>
    <div class="footer">{html_escape(s.footer_note)}</div>
</div>
<script>
function filterTable() {{
    let input = document.getElementById('searchInput').value.toLowerCase();
    let rows = document.getElementById('dataTable').getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    for (let row of rows) {{
        let text = row.textContent.toLowerCase();
        row.style.display = text.includes(input) ? "" : "none";
    }}
}}
</script>
</body>
</html>"""
        output.parent.mkdir(parents=True, exist_ok=True)
        with open(output, 'w', encoding='utf-8') as f:
            f.write(html)

    # --------------------- Excel ---------------------
    def export_excel(self, records: list[dict[str, Any]], recipient: str, output: Path,
                     selected_fields: Optional[List[FieldDefinition]] = None) -> None:
        if Workbook is None:
            raise RuntimeError("openpyxl غير مثبتة.")
        fields = [f for f in (selected_fields if selected_fields else self.settings.report_fields) if f.key != "scheduled_delivery_date"]
        metrics = report_record_metrics(records)
        wb = Workbook()
        ws = wb.active
        ws.title = "تقرير التسليم"
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A7"
        ws.print_title_rows = "6:6"

        count = len(fields)
        use_portrait = count <= 4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.sheet_properties.pageSetUpPr.autoPageBreaks = False
        ws.page_setup.orientation = "portrait" if use_portrait else "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        # صفحة واحدة عرضًا، وعدد غير محدود من الصفحات إلى الأسفل.
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = None
        ws.page_setup.blackAndWhite = self.settings.print_friendly_reports
        ws.print_options.horizontalCentered = True

        ws.oddFooter.center.text = f"صفحة &P من &N | {self.settings.company_name}"
        ws.oddFooter.center.size = 10

        palette = self._palette()
        title_fill = palette["title_fill"].lstrip("#")
        title_text = palette["title_text"].lstrip("#")
        subtitle_fill = palette["subtitle_fill"].lstrip("#")
        subtitle_text = palette["subtitle_text"].lstrip("#")
        header_fill = palette["header_fill"].lstrip("#")
        header_text = palette["header_text"].lstrip("#")
        row_even_fill = palette["row_alt"].lstrip("#")
        row_odd_fill = "FFFFFF"
        border_color = palette["border"].lstrip("#")
        accent_color = palette["accent"].lstrip("#")

        thin = Side(style="thin", color=border_color)
        medium = Side(style="medium", color=border_color)
        thick = Side(style="thick", color="343A40")
        border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
        border_header = Border(left=medium, right=medium, top=medium, bottom=medium)

        # العنوان الرئيسي
        title_row = 1
        ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=count)
        cell = ws.cell(title_row, 1, f"{self.settings.report_title}\n{self.settings.report_title_en}")
        cell.font = Font(name=self.settings.font_family, size=22, bold=True, color=title_text)
        cell.fill = PatternFill("solid", fgColor=subtitle_fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[title_row].height = 46

        # العنوان الفرعي
        sub_row = 2
        ws.merge_cells(start_row=sub_row, start_column=1, end_row=sub_row, end_column=count)
        cell = ws.cell(sub_row, 1, f"{self.settings.company_name} | {self.settings.report_subtitle}")
        cell.font = Font(name=self.settings.font_family, size=14, bold=True, color=subtitle_text)
        cell.fill = PatternFill("solid", fgColor=title_fill)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[sub_row].height = 26
        ws.row_dimensions[3].height = 6
        ws.row_dimensions[5].height = 6

        now = dt.datetime.now()
        meta_row = 4
        meta_info = [
            f"السجلات: {metrics['records']}",
            f"العقود: {metrics['contracts']}",
            f"العقود المتكررة: {metrics['multi_kitchen_contracts']}",
            f"الإصدار: {now:%Y-%m-%d}  {now:%H:%M}",
        ][:max(1, min(4, count))]
        segment = max(1, count // 3)
        meta_segments = [
            (math.floor(index * count / len(meta_info)) + 1, math.floor((index + 1) * count / len(meta_info)))
            for index in range(len(meta_info))
        ]
        for i, (start, end) in enumerate(meta_segments):
            ws.merge_cells(start_row=meta_row, start_column=start, end_row=meta_row, end_column=end)
            cell = ws.cell(meta_row, start, meta_info[i])
            cell.font = Font(name=self.settings.font_family, size=10, bold=True, color="495057")
            cell.fill = PatternFill("solid", fgColor="E9ECEF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[meta_row].height = 24

        # توزيع العرض على صفحة واحدة مع إعطاء النصوص الطويلة مساحة أكبر.
        target_total_width = 82.0 if use_portrait else 118.0
        natural_widths = report_column_weights(fields, records)
        width_factor = target_total_width / sum(natural_widths) if sum(natural_widths) else 1.0
        minimum_column_width = 4.5 if count <= 20 else 3.2
        excel_widths = [max(minimum_column_width, width * width_factor) for width in natural_widths]
        body_font_size = max(9.0, min(11.2, 12.2 - max(0, count - 6) * 0.42))
        header_font_size = min(12.0, body_font_size + 1.2)

        # رأس الجدول
        header_row = 6
        for col, item in enumerate(fields, 1):
            cell = ws.cell(header_row, col, item.label)
            cell.font = Font(name=self.settings.font_family, size=header_font_size, bold=True, color=header_text)
            cell.fill = PatternFill("solid", fgColor=header_fill)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_header
            ws.column_dimensions[get_column_letter(col)].width = excel_widths[col - 1]
        header_lines = max(
            1,
            max(
                math.ceil((len(item.label) * 1.0) / max(4.0, excel_widths[index] - 1.5))
                for index, item in enumerate(fields)
            ),
        )
        ws.row_dimensions[header_row].height = min(42, max(30, header_lines * 12 + 7))

        # البيانات
        for row_idx, record in enumerate(records, header_row + 1):
            fill = row_even_fill if row_idx % 2 == 0 else row_odd_fill
            row_values = self._values(record, fields)
            for col, value in enumerate(row_values, 1):
                cell = ws.cell(row_idx, col, value)
                cell.font = Font(name=self.settings.font_family, size=body_font_size, color="212529")
                cell.fill = PatternFill("solid", fgColor=fill)
                if fields[col - 1].key == "delivery_status" and clean_value(value) in REPORT_STATUS_FILLS:
                    cell.fill = PatternFill("solid", fgColor=REPORT_STATUS_FILLS[clean_value(value)])
                    cell.font = Font(
                        name=self.settings.font_family, size=body_font_size, bold=True, color="212529",
                    )
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True, shrink_to_fit=False,
                )
                cell.border = border_thin
            ws.row_dimensions[row_idx].height = 48

        # فواصل رأسية بواقع 9 أو 10 سجلات، والصفحة الأخيرة تحمل المتبقي فقط.
        if Break is not None and records:
            page_row_target = report_rows_per_page(fields)
            for cumulative_rows in range(page_row_target, len(records), page_row_target):
                ws.row_breaks.append(Break(id=header_row + cumulative_rows))

        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(count)}{header_row + len(records)}"

        # حدود خارجية
        last_data_row = max(header_row + 1, header_row + len(records))
        for col in range(1, count + 1):
            ws.cell(header_row, col).border = Border(
                left=ws.cell(header_row, col).border.left,
                right=ws.cell(header_row, col).border.right,
                top=thick,
                bottom=ws.cell(header_row, col).border.bottom
            )
            ws.cell(last_data_row, col).border = Border(
                left=ws.cell(last_data_row, col).border.left,
                right=ws.cell(last_data_row, col).border.right,
                top=ws.cell(last_data_row, col).border.top,
                bottom=thick
            )
        for row in range(header_row, last_data_row + 1):
            ws.cell(row, 1).border = Border(
                left=thick,
                right=ws.cell(row, 1).border.right,
                top=ws.cell(row, 1).border.top,
                bottom=ws.cell(row, 1).border.bottom
            )
            ws.cell(row, count).border = Border(
                left=ws.cell(row, count).border.left,
                right=thick,
                top=ws.cell(row, count).border.top,
                bottom=ws.cell(row, count).border.bottom
            )

        # تذييل
        footer_row = last_data_row + 3
        footer_data = [
            f"اسم المستلم: {clean_value(recipient) or '........................'}",
            "التوقيع: ........................",
            f"التاريخ: {now:%Y-%m-%d}"
        ]
        for i, (start, end) in enumerate([(1, segment), (segment+1, min(count, segment*2)), (min(count, segment*2)+1, count)]):
            if start > end:
                continue
            ws.merge_cells(start_row=footer_row, start_column=start, end_row=footer_row, end_column=end)
            cell = ws.cell(footer_row, start, footer_data[i])
            cell.font = Font(name=self.settings.font_family, size=10, bold=True, color="495057")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=Side(style="medium", color=accent_color))
        ws.row_dimensions[footer_row].height = 26

        note_row = footer_row + 2
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=count)
        cell = ws.cell(
            note_row, 1,
            f"{self.settings.footer_note} | يُسمح بتكرار رقم العقد. "
            "حالة «تعديل - يعتمد الحالي ويلغى السابق» تحفظ السابق كسجل ملغي ولا تحذفه.",
        )
        cell.font = Font(name=self.settings.font_family, size=9, italic=True, color="6C757D")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[note_row].height = 32

        # شعار
        if self.settings.logo_path and Path(self.settings.logo_path).is_file() and ExcelImage is not None and count >= 3:
            try:
                logo = ExcelImage(self.settings.logo_path)
                logo.width, logo.height = 100, 60
                ws.add_image(logo, "A1")
            except Exception:
                pass

        ws.print_area = f"A1:{get_column_letter(count)}{note_row}"
        ws.page_margins.left = 0.05
        ws.page_margins.right = 0.05
        ws.page_margins.top = 0.10
        ws.page_margins.bottom = 0.10
        ws.page_margins.header = 0.03
        ws.page_margins.footer = 0.03
        self._add_excel_summary(wb, ws, records, fields, palette)
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)

    def _add_excel_summary(self, workbook: Any, data_sheet: Any, records: list[dict[str, Any]],
                           fields: list[FieldDefinition], palette: dict[str, str]) -> None:
        """صفحة ملخص تنفيذية مرتبطة بصيغ ببيانات ورقة التقرير."""
        summary = workbook.create_sheet("لوحة الملخص", 0)
        summary.sheet_view.rightToLeft = True
        summary.sheet_view.showGridLines = False
        summary.freeze_panes = "A7"
        summary.page_setup.orientation = "landscape"
        summary.page_setup.paperSize = summary.PAPERSIZE_A4
        summary.page_setup.fitToWidth = 1
        summary.page_setup.fitToHeight = 0
        summary.sheet_properties.pageSetUpPr.fitToPage = True
        summary.page_setup.blackAndWhite = self.settings.print_friendly_reports
        summary.page_margins.left = 0.05
        summary.page_margins.right = 0.05
        summary.page_margins.top = 0.10
        summary.page_margins.bottom = 0.10
        summary.merge_cells("A1:H2")
        title = summary["A1"]
        title.value = "لوحة ملخص تسليم الرخام\nMarble Delivery Dashboard"
        title.font = Font(name=self.settings.font_family, size=24, bold=True, color=palette["title_text"].lstrip("#"))
        title.fill = PatternFill("solid", fgColor=palette["title_fill"].lstrip("#"))
        title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        summary.row_dimensions[1].height = 34
        summary.row_dimensions[2].height = 24

        field_columns = {field.key: index + 1 for index, field in enumerate(fields)}
        first_data_row = 7
        last_data_row = max(first_data_row, first_data_row + len(records) - 1)
        metrics = report_record_metrics(records)
        kpis = [
            ("A4:B5", "السجلات", str(metrics["records"])),
            ("D4:E5", "العقود", str(metrics["contracts"])),
            ("G4:H5", "العقود المتكررة", str(metrics["multi_kitchen_contracts"])),
        ]
        for area, label, value in kpis:
            summary.merge_cells(area)
            cell = summary[area.split(":")[0]]
            cell.value = f"{label}\n{value}"
            cell.font = Font(name=self.settings.font_family, size=15, bold=True, color=palette["text"].lstrip("#"))
            cell.fill = PatternFill("solid", fgColor=palette["meta_fill"].lstrip("#"))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(
                left=Side(style="thin", color=palette["border"].lstrip("#")),
                right=Side(style="thin", color=palette["border"].lstrip("#")),
                top=Side(style="thin", color=palette["border"].lstrip("#")),
                bottom=Side(style="medium", color=palette["accent"].lstrip("#")),
            )

        blocks = [
            ("showroom", "الفرع / المعرض", 1, 7),
            ("marble_color", "لون الرخام", 4, 7),
            ("salesperson", "موظف المعرض", 7, 7),
            ("sink_type", "نوع المجلى", 1, 19),
            ("marble_edge", "شفة الرخام", 4, 19),
            ("delivery_status", "حالة التسليم", 7, 19),
        ]
        chart_specs: list[tuple[int, int, int, int, str]] = []
        for key, label, start_col, header_row in blocks:
            if key not in field_columns:
                continue
            counts: dict[str, int] = {}
            for record in records:
                value = clean_value(record.get("values", {}).get(key, "")) or "غير محدد"
                counts[value] = counts.get(value, 0) + 1
            top_items = sorted(counts.items(), key=lambda pair: (-pair[1], pair[0]))[:8]
            summary.cell(header_row, start_col, label)
            summary.cell(header_row, start_col + 1, "عدد السجلات")
            for cell in summary.iter_cols(min_col=start_col, max_col=start_col + 1, min_row=header_row, max_row=header_row):
                for header_cell in cell:
                    header_cell.font = Font(name=self.settings.font_family, size=11, bold=True, color=palette["header_text"].lstrip("#"))
                    header_cell.fill = PatternFill("solid", fgColor=palette["header_fill"].lstrip("#"))
                    header_cell.alignment = Alignment(horizontal="center")
            data_col_letter = get_column_letter(field_columns[key])
            for offset, (category, _count) in enumerate(top_items, header_row + 1):
                summary.cell(offset, start_col, category)
                summary.cell(offset, start_col + 1, f'=COUNTIF(\'{data_sheet.title}\'!${data_col_letter}${first_data_row}:${data_col_letter}${last_data_row},{get_column_letter(start_col)}{offset})')
                for col in (start_col, start_col + 1):
                    cell = summary.cell(offset, col)
                    cell.font = Font(name=self.settings.font_family, size=10, color=palette["text"].lstrip("#"))
                    cell.fill = PatternFill("solid", fgColor=(palette["row_alt"] if offset % 2 == 0 else "#FFFFFF").lstrip("#"))
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = Border(bottom=Side(style="thin", color=palette["border"].lstrip("#")))
            if top_items:
                chart_specs.append((start_col, start_col + 1, header_row, header_row + len(top_items), f"توزيع السجلات حسب {label}"))

        summary_last_row = 30
        if BarChart is not None and Reference is not None:
            positions = [
                ("A31", 46), ("E31", 46), ("A48", 63),
                ("E48", 63), ("A65", 80), ("E65", 80),
            ]
            for (position, chart_end_row), (category_col, value_col, header_row, max_row, title_text) in zip(positions, chart_specs):
                chart = BarChart()
                chart.type = "bar"
                chart.style = 10
                chart.title = title_text
                chart.y_axis.title = "الفئة"
                chart.x_axis.title = "عدد السجلات"
                chart.x_axis.scaling.min = 0
                chart.x_axis.scaling.max = max(3, len(records))
                chart.x_axis.majorUnit = 1
                chart.x_axis.numFmt = "0"
                chart.x_axis.delete = True
                chart.x_axis.tickLblPos = "none"
                chart.x_axis.majorGridlines = None
                chart.height = 6.6
                chart.width = 10.8
                chart.legend = None
                if DataLabelList is not None:
                    chart.dLbls = DataLabelList()
                    chart.dLbls.showVal = True
                    chart.dLbls.showLegendKey = False
                chart.add_data(Reference(summary, min_col=value_col, min_row=header_row, max_row=max_row), titles_from_data=True)
                chart.set_categories(Reference(summary, min_col=category_col, min_row=header_row + 1, max_row=max_row))
                summary.add_chart(chart, position)
                summary_last_row = max(summary_last_row, chart_end_row)

        for col in range(1, 9):
            summary.column_dimensions[get_column_letter(col)].width = 18
        summary.print_area = f"A1:H{summary_last_row}"
        summary.oddFooter.center.text = f"صفحة &P من &N | {self.settings.company_name}"

    def _resolve_report_page(self, fields: list[FieldDefinition]):
        """يعيد مقاس الصفحة واتجاهها الحقيقيين وفق إعدادات التقرير، بما في ذلك Letter والمقاس المخصص."""
        page_code = str(self.settings.report_page_size).upper()
        if page_code == "A3":
            base_size = A3
        elif page_code == "LETTER":
            base_size = LETTER
        elif page_code == "CUSTOM":
            base_size = (
                max(80.0, float(self.settings.custom_page_width_mm)) * mm,
                max(80.0, float(self.settings.custom_page_height_mm)) * mm,
            )
        else:
            base_size = A4
        orientation = str(self.settings.report_orientation).lower()
        if orientation == "auto":
            orientation = "portrait" if len(fields) <= 4 else "landscape"
        page_size = portrait(base_size) if orientation == "portrait" else landscape(base_size)
        return page_size, page_size[0], page_size[1], orientation

    @staticmethod
    def _page_mm(page_size: tuple[float, float]) -> tuple[float, float]:
        return page_size[0] / mm, page_size[1] / mm

    # --------------------- PDF ---------------------
    def _export_pdf_one_page(self, records: list[dict[str, Any]], recipient: str, output: Path,
                             selected_fields: Optional[List[FieldDefinition]] = None) -> None:
        """ينشئ تقرير A4 من صفحة واحدة ويختار أكبر خط يمكن أن يتسع فعلياً."""
        self._require_arabic_shaping()
        regular_name, bold_name = self._register_pdf_fonts()
        fields = list(reversed([f for f in (selected_fields if selected_fields else self.settings.report_fields) if f.key != "scheduled_delivery_date"]))
        page_size = landscape(A4)
        page_width, page_height = page_size
        palette = self._palette()
        title_fill = colors.HexColor(palette["title_fill"])
        title_text = colors.HexColor(palette["title_text"])
        subtitle_fill = colors.HexColor(palette["subtitle_fill"])
        subtitle_text = colors.HexColor(palette["subtitle_text"])
        header_fill = colors.HexColor(palette["header_fill"])
        header_text = colors.HexColor(palette["header_text"])
        row_alt = colors.HexColor(palette["row_alt"])
        border_color = colors.HexColor(palette["border"])
        accent = colors.HexColor(palette["accent"])
        muted = colors.HexColor(palette["muted"])
        text_color = colors.HexColor(palette["text"])

        side_margin = 2.5 * mm
        table_bottom = 17 * mm
        table_top = page_height - 36 * mm
        available_width = page_width - 2 * side_margin
        available_height = table_top - table_bottom
        raw_widths = report_column_weights(fields, records)
        width_factor = available_width / sum(raw_widths) if sum(raw_widths) else 1.0
        col_widths = [width * width_factor for width in raw_widths]

        output.parent.mkdir(parents=True, exist_ok=True)
        pdf = PDFCanvas(str(output), pagesize=page_size)
        pdf.setTitle(f"{self.settings.report_title} - {self.settings.report_title_en}")
        pdf.setAuthor(self.settings.company_name)

        def build_table(font_size: float) -> tuple[Table, float, float]:
            body_leading = max(font_size + 1.2, font_size * 1.22)
            header_leading = max(font_size + 1.4, font_size * 1.25)
            normal_style = ParagraphStyle(
                f"SinglePageCell{font_size}", fontName=regular_name, fontSize=font_size,
                leading=body_leading, alignment=TA_CENTER, textColor=text_color,
            )
            header_style = ParagraphStyle(
                f"SinglePageHeader{font_size}", fontName=bold_name, fontSize=font_size,
                leading=header_leading, alignment=TA_CENTER, textColor=header_text,
            )

            def paragraph(value: Any, style: ParagraphStyle) -> Paragraph:
                return Paragraph(escape(shape_arabic(value)), style)

            data = [[paragraph(item.label, header_style) for item in fields]]
            for record in records:
                data.append([paragraph(value, normal_style) for value in self._values(record, fields)])
            if not records:
                data.append([
                    paragraph("لا توجد سجلات ضمن هذا التقرير", normal_style),
                    *[paragraph("", normal_style) for _ in fields[1:]],
                ])
            report_table = Table(data, colWidths=col_widths, hAlign="CENTER", splitByRow=0)
            padding = max(1.0, font_size * 0.28)
            commands = [
                ("BACKGROUND", (0, 0), (-1, 0), header_fill),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.45, border_color),
                ("BOX", (0, 0), (-1, -1), 0.8, accent),
                ("LINEBELOW", (0, 0), (-1, 0), 0.8, accent),
                ("TOPPADDING", (0, 0), (-1, -1), padding),
                ("BOTTOMPADDING", (0, 0), (-1, -1), padding),
                ("LEFTPADDING", (0, 0), (-1, -1), max(1.0, padding * 0.75)),
                ("RIGHTPADDING", (0, 0), (-1, -1), max(1.0, padding * 0.75)),
            ]
            for row_index in range(1, len(data)):
                if row_index % 2 == 0:
                    commands.append(("BACKGROUND", (0, row_index), (-1, row_index), row_alt))
            report_table.setStyle(TableStyle(commands))
            wrapped_width, wrapped_height = report_table.wrapOn(pdf, available_width, available_height)
            return report_table, wrapped_width, wrapped_height

        selected_table = None
        table_width = table_height = 0.0
        selected_font_size = REPORT_MIN_FONT_SIZE
        candidate = REPORT_MAX_FONT_SIZE
        while candidate >= REPORT_MIN_FONT_SIZE - 0.01:
            candidate_table, candidate_width, candidate_height = build_table(candidate)
            selected_table, table_width, table_height = candidate_table, candidate_width, candidate_height
            selected_font_size = candidate
            if candidate_height <= available_height + 0.1:
                break
            candidate -= 0.5

        scale = min(1.0, available_width / max(1.0, table_width), available_height / max(1.0, table_height))
        drawn_width = table_width * scale
        drawn_height = table_height * scale
        table_x = side_margin + (available_width - drawn_width) / 2
        table_y = table_top - drawn_height

        # رأس مدمج يحافظ على أكبر مساحة ممكنة للبيانات.
        pdf.setFillColor(title_fill)
        pdf.rect(0, page_height - 18*mm, page_width, 18*mm, stroke=0, fill=1)
        pdf.setFillColor(subtitle_fill)
        pdf.rect(0, page_height - 28*mm, page_width, 10*mm, stroke=0, fill=1)
        pdf.setFillColor(accent)
        pdf.rect(0, page_height - 29*mm, page_width, 1.2*mm, stroke=0, fill=1)
        pdf.setFillColor(title_text)
        pdf.setFont(bold_name, 18)
        pdf.drawCentredString(page_width/2, page_height - 8*mm, shape_arabic(self.settings.report_title))
        pdf.setFont(bold_name, 8.5)
        pdf.drawCentredString(page_width/2, page_height - 15*mm, self.settings.report_title_en)
        pdf.setFillColor(subtitle_text)
        pdf.setFont(regular_name, 9.5)
        pdf.drawCentredString(
            page_width/2, page_height - 23.5*mm,
            shape_arabic(f"{self.settings.company_name} | {self.settings.report_subtitle}"),
        )

        now = dt.datetime.now()
        reference = report_reference(records, now)
        metrics = report_record_metrics(records)
        pdf.setFillColor(text_color)
        pdf.setFont(regular_name, 7.5)
        meta_text = (
            f"السجلات: {metrics['records']}     |     العقود: {metrics['contracts']}"
            f"     |     العقود المتكررة: {metrics['multi_kitchen_contracts']}     |     التاريخ: {now:%Y-%m-%d}"
            f"     |     وقت الإصدار: {now:%H:%M}     |     مرجع التقرير: {reference}"
        )
        pdf.drawCentredString(page_width/2, page_height - 33*mm, shape_arabic(meta_text))

        if self.settings.logo_path and Path(self.settings.logo_path).is_file():
            try:
                logo_width = max(15, min(45, int(self.settings.logo_width_mm))) * mm
                pdf.drawImage(
                    self.settings.logo_path, 7*mm, page_height - 26*mm,
                    width=logo_width, height=min(21*mm, logo_width * 0.68),
                    preserveAspectRatio=True, mask="auto",
                )
            except Exception:
                logging.exception("تعذر إضافة الشعار إلى PDF")

        try:
            qr_size = 13 * mm
            qr_widget = QrCodeWidget("|".join([
                "MarbleDeliveryPro", f"Company={clean_value(self.settings.company_name)}",
                f"Report={clean_value(self.settings.report_title)}", f"Reference={reference}",
                f"Created={now:%Y-%m-%dT%H:%M}", f"Records={len(records)}",
            ]))
            x1, y1, x2, y2 = qr_widget.getBounds()
            qr_drawing = Drawing(
                qr_size, qr_size,
                transform=[qr_size/(x2-x1), 0, 0, qr_size/(y2-y1), 0, 0],
            )
            qr_drawing.add(qr_widget)
            renderPDF.draw(qr_drawing, pdf, page_width - 20*mm, page_height - 15.5*mm)
        except Exception:
            logging.exception("تعذر إضافة رمز التحقق إلى PDF")

        pdf.saveState()
        pdf.translate(table_x, table_y)
        pdf.scale(scale, scale)
        selected_table.drawOn(pdf, 0, 0)
        pdf.restoreState()

        pdf.setFillColor(text_color)
        pdf.setFont(regular_name, 7.5)
        pdf.drawRightString(
            page_width - side_margin, 14*mm,
            shape_arabic(f"اسم المستلم: {clean_value(recipient) or '........................'}"),
        )
        pdf.drawCentredString(page_width/2, 14*mm, shape_arabic("التوقيع: ........................"))
        pdf.drawString(side_margin, 14*mm, shape_arabic(f"التاريخ: {now:%Y-%m-%d}"))
        pdf.setStrokeColor(accent)
        pdf.line(side_margin, 11*mm, page_width - side_margin, 11*mm)
        pdf.setFillColor(muted)
        pdf.setFont(regular_name, 7)
        pdf.drawRightString(page_width - side_margin, 5.5*mm, shape_arabic(self.settings.footer_note))
        pdf.drawString(side_margin, 5.5*mm, shape_arabic("صفحة 1 من 1"))
        pdf.drawCentredString(page_width/2, 5.5*mm, f"REF: {reference} | FONT: {selected_font_size:g}")
        pdf.showPage()
        pdf.save()

    def export_pdf(self, records: list[dict[str, Any]], recipient: str, output: Path,
                   selected_fields: Optional[List[FieldDefinition]] = None) -> None:
        """تصدير PDF متجهي عالي الوضوح مع منع تصغير الخط إلى أحجام غير مقروءة."""
        self._require_arabic_shaping()
        regular_name, bold_name = self._register_pdf_fonts()
        fields = list(reversed([f for f in (selected_fields if selected_fields else self.settings.report_fields) if f.key != "scheduled_delivery_date"]))
        page_size, page_width, page_height, orientation = self._resolve_report_page(fields)

        raw = self._palette()
        title_fill = colors.HexColor(raw["title_fill"])
        title_text = colors.HexColor(raw["title_text"])
        subtitle_fill = colors.HexColor(raw["subtitle_fill"])
        subtitle_text = colors.HexColor(raw["subtitle_text"])
        header_fill = colors.HexColor(raw["header_fill"])
        header_text = colors.HexColor(raw["header_text"])
        row_alt = colors.HexColor(raw["row_alt"])
        border_color = colors.HexColor(raw["border"])
        accent = colors.HexColor(raw["accent"])
        muted = colors.HexColor(raw["muted"])
        text_color = colors.HexColor(raw["text"])

        side_margin = max(3.0, float(self.settings.report_margin_mm)) * mm
        available_width = page_width - 2 * side_margin
        right_style = ParagraphStyle(
            "ArabicRightV8", fontName=regular_name, fontSize=9.2, leading=11.4,
            alignment=TA_RIGHT, textColor=text_color,
        )

        def para(value: Any, style: ParagraphStyle) -> Paragraph:
            return Paragraph(escape(shape_arabic(value)), style)

        output.parent.mkdir(parents=True, exist_ok=True)
        doc = SimpleDocTemplate(
            str(output), pagesize=page_size,
            rightMargin=side_margin, leftMargin=side_margin,
            topMargin=39*mm, bottomMargin=24*mm,
            title=f"{self.settings.report_title} - {self.settings.report_title_en}",
            author=self.settings.company_name,
        )

        now = dt.datetime.now()
        reference = report_reference(records, now)
        metrics = report_record_metrics(records)

        def build_meta(page_records: list[dict[str, Any]]) -> Table:
            page_w_mm, page_h_mm = self._page_mm(page_size)
            meta = Table(
                [[
                    para(f"السجلات: {metrics['records']}", right_style),
                    para(f"العقود: {metrics['contracts']}", right_style),
                    para(f"العقود المتكررة: {metrics['multi_kitchen_contracts']}", right_style),
                    para(f"هذه الصفحة: {len(page_records)} | {now:%Y-%m-%d}", right_style),
                ]],
                colWidths=[available_width/4] * 4,
            )
            meta.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F0F3F5")),
                ("BOX", (0, 0), (-1, -1), 0.65, border_color),
                ("LINEBELOW", (0, 0), (-1, -1), 1.0, accent),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5.0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5.0),
            ]))
            return meta

        column_weights = report_column_weights(fields, records)
        width_factor = available_width / sum(column_weights) if sum(column_weights) else 1.0
        col_widths = [weight * width_factor for weight in column_weights]

        def build_page_table(page_records: list[dict[str, Any]], font_size: float) -> Table:
            header_font_size = min(font_size + 0.9, 11.2)
            normal_style = ParagraphStyle(
                f"ArabicCellV8_{font_size:.2f}", fontName=regular_name, fontSize=font_size,
                leading=max(font_size + 1.3, font_size * 1.22), alignment=TA_CENTER,
                textColor=text_color, splitLongWords=True,
            )
            status_style = ParagraphStyle(
                f"ArabicStatusV8_{font_size:.2f}", parent=normal_style,
                fontName=bold_name, textColor=text_color,
            )
            header_style = ParagraphStyle(
                f"ArabicHeaderV8_{font_size:.2f}", fontName=bold_name, fontSize=header_font_size,
                leading=max(header_font_size + 1.2, header_font_size * 1.20),
                alignment=TA_CENTER, textColor=header_text, splitLongWords=True,
            )
            table_data = [[para(item.label, header_style) for item in fields]]
            for record in page_records:
                values = self._values(record, fields)
                table_data.append([
                    para(
                        value,
                        status_style
                        if item.key == "delivery_status" and clean_value(value) in REPORT_STATUS_FILLS
                        else normal_style,
                    )
                    for item, value in zip(fields, values)
                ])
            if not page_records:
                table_data.append([
                    para("لا توجد سجلات ضمن هذا التقرير", normal_style),
                    *[para("", normal_style) for _ in fields[1:]],
                ])
            table = Table(
                table_data, colWidths=col_widths, repeatRows=1,
                hAlign="CENTER", splitByRow=0,
            )
            cell_padding = max(4.4, font_size * 0.50)
            style_commands = [
                ("BACKGROUND", (0, 0), (-1, 0), header_fill),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.55, border_color),
                ("BOX", (0, 0), (-1, -1), 1.10, accent),
                ("LINEBELOW", (0, 0), (-1, 0), 1.15, accent),
                ("TOPPADDING", (0, 0), (-1, 0), cell_padding + 1.2),
                ("BOTTOMPADDING", (0, 0), (-1, 0), cell_padding + 1.2),
                ("LEFTPADDING", (0, 0), (-1, -1), max(1.0, cell_padding * 0.70)),
                ("RIGHTPADDING", (0, 0), (-1, -1), max(1.0, cell_padding * 0.70)),
                ("TOPPADDING", (0, 1), (-1, -1), cell_padding),
                ("BOTTOMPADDING", (0, 1), (-1, -1), cell_padding),
            ]
            for row_index in range(1, len(table_data)):
                if row_index % 2 == 0:
                    style_commands.append(("BACKGROUND", (0, row_index), (-1, row_index), row_alt))
            status_column = next(
                (index for index, item in enumerate(fields) if item.key == "delivery_status"), None,
            )
            if status_column is not None:
                for row_index, record in enumerate(page_records, 1):
                    status = clean_value(record.get("values", {}).get("delivery_status", ""))
                    fill = REPORT_STATUS_FILLS.get(status)
                    if fill:
                        style_commands.append((
                            "BACKGROUND", (status_column, row_index), (status_column, row_index),
                            colors.HexColor(f"#{fill}"),
                        ))
            table.setStyle(TableStyle(style_commands))
            return table

        preferred_font_size = min(
            float(self.settings.pdf_base_font_size),
            max(float(self.settings.pdf_min_font_size), 11.2 - max(0, len(fields) - 6) * 0.22),
        )
        minimum_font_size = float(self.settings.pdf_min_font_size)

        def fit_chunk(page_records: list[dict[str, Any]]):
            meta = build_meta(page_records)
            _, meta_height = meta.wrap(available_width, doc.height)
            table_height_limit = doc.height - meta_height - 3*mm - 14
            candidate = preferred_font_size
            while candidate >= minimum_font_size - 0.01:
                table = build_page_table(page_records, candidate)
                _, table_height = table.wrap(available_width, table_height_limit)
                if table_height <= table_height_limit + 0.2:
                    return candidate, table_height
                candidate -= 0.25
            return None

        page_specs: list[tuple[list[dict[str, Any]], float]] = []
        if not records:
            fitted = fit_chunk([])
            if not fitted:
                raise RuntimeError("تعذر إنشاء صفحة التقرير بالإعدادات الحالية.")
            page_specs.append(([], fitted[0]))
        else:
            cursor = 0
            while cursor < len(records):
                chunk, fitted = largest_fitting_page_chunk(records[cursor:], fit_chunk)
                if not fitted or not chunk:
                    raise RuntimeError(
                        "يوجد نص طويل جداً لا يمكن ملاءمته داخل الصفحة مع الحد الأدنى المقروء للخط. "
                        "اختر A3 أو قلل عدد الحقول في التقرير."
                    )
                page_specs.append((chunk, fitted[0]))
                cursor += len(chunk)

        total_pages = len(page_specs)

        def draw_page(canvas, doc):
            canvas.saveState()
            canvas.setFillColor(title_fill)
            canvas.rect(0, page_height - 22*mm, page_width, 22*mm, stroke=0, fill=1)
            canvas.setFillColor(subtitle_fill)
            canvas.rect(0, page_height - 35*mm, page_width, 13*mm, stroke=0, fill=1)
            canvas.setFillColor(accent)
            canvas.rect(0, page_height - 36*mm, page_width, 1.3*mm, stroke=0, fill=1)

            title_size = 24 if str(self.settings.report_page_size).upper() == "A4" else 27
            canvas.setFont(bold_name, title_size)
            canvas.setFillColor(title_text)
            canvas.drawCentredString(page_width/2, page_height - 10*mm, shape_arabic(self.settings.report_title))
            canvas.setFont(bold_name, 11.0)
            canvas.drawCentredString(page_width/2, page_height - 18*mm, self.settings.report_title_en)
            canvas.setFont(regular_name, 11.5)
            canvas.setFillColor(subtitle_text)
            canvas.drawCentredString(
                page_width/2, page_height - 29*mm,
                shape_arabic(f"{self.settings.company_name} | {self.settings.report_subtitle}")
            )

            if self.settings.logo_path and Path(self.settings.logo_path).is_file():
                try:
                    logo_width = max(15, min(60, int(self.settings.logo_width_mm))) * mm
                    canvas.drawImage(
                        self.settings.logo_path, side_margin, page_height - 31*mm,
                        width=logo_width, height=min(27*mm, logo_width * 0.72),
                        preserveAspectRatio=True, mask="auto"
                    )
                except Exception:
                    logging.exception("تعذر إضافة الشعار إلى PDF")

            try:
                qr_size = 17 * mm
                qr_widget = QrCodeWidget("|".join([
                    "MarbleDeliveryPro",
                    f"Company={clean_value(self.settings.company_name)}",
                    f"Report={clean_value(self.settings.report_title)}",
                    f"Reference={reference}",
                    f"Created={now:%Y-%m-%dT%H:%M}",
                    f"Records={len(records)}",
                ]))
                x1, y1, x2, y2 = qr_widget.getBounds()
                qr_drawing = Drawing(
                    qr_size, qr_size,
                    transform=[qr_size/(x2-x1), 0, 0, qr_size/(y2-y1), 0, 0]
                )
                qr_drawing.add(qr_widget)
                renderPDF.draw(qr_drawing, canvas, page_width - side_margin - qr_size, page_height - 22*mm)
            except Exception:
                logging.exception("تعذر إضافة رمز التحقق إلى PDF")

            canvas.setFont(regular_name, 9.2)
            canvas.setFillColor(text_color)
            canvas.drawRightString(
                page_width - side_margin, 20*mm,
                shape_arabic(f"اسم المستلم: {clean_value(recipient) or '........................'}"),
            )
            canvas.drawCentredString(page_width/2, 20*mm, shape_arabic("التوقيع: ........................"))
            canvas.drawString(side_margin, 20*mm, shape_arabic(f"التاريخ: {now:%Y-%m-%d}"))
            canvas.setStrokeColor(accent)
            canvas.line(side_margin, 16*mm, page_width - side_margin, 16*mm)

            canvas.setStrokeColor(border_color)
            canvas.line(side_margin, 11*mm, page_width - side_margin, 11*mm)
            canvas.setFont(regular_name, 8.2)
            canvas.setFillColor(muted)
            canvas.drawRightString(page_width - side_margin, 6*mm, shape_arabic(self.settings.footer_note))
            canvas.drawString(side_margin, 6*mm, shape_arabic(f"صفحة {doc.page} من {total_pages}"))
            page_w_mm, page_h_mm = self._page_mm(page_size)
            canvas.drawCentredString(
                page_width/2, 6*mm,
                f"REF: {reference} | {page_w_mm:.0f}x{page_h_mm:.0f} mm | {orientation.upper()}"
            )
            canvas.restoreState()

        story = []
        for page_index, (page_records, font_size) in enumerate(page_specs):
            meta = build_meta(page_records)
            table = build_page_table(page_records, font_size)
            logging.info(
                "PDF v10.0.4 page %s/%s: %s rows, font %.2f pt, %s %s",
                page_index + 1, total_pages, len(page_records), font_size,
                self.settings.report_page_size, orientation,
            )
            story.extend([meta, Spacer(1, 3*mm), table])
            if page_index < total_pages - 1:
                story.append(PageBreak())
        doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)

    # --------------------- PNG ---------------------
    def _pil_fonts(self) -> tuple[Any, Any, Any]:
        if ImageFont is None:
            raise RuntimeError("Pillow غير مثبتة.")
        if self.regular_font_path:
            return (
                ImageFont.truetype(self.bold_font_path or self.regular_font_path, 58),
                ImageFont.truetype(self.bold_font_path or self.regular_font_path, 44),
                ImageFont.truetype(self.regular_font_path, 44),
            )
        return ImageFont.load_default(), ImageFont.load_default(), ImageFont.load_default()

    @staticmethod
    def _wrap_text(draw, text, font, max_width):
        text = clean_value(text)
        if not text:
            return [""]
        words = text.split(" ")
        fitted_words = []
        for word in words:
            word_box = draw.textbbox((0, 0), shape_arabic(word), font=font)
            if word_box[2] - word_box[0] <= max_width:
                fitted_words.append(word)
                continue
            piece = ""
            for character in word:
                candidate = piece + character
                candidate_box = draw.textbbox((0, 0), shape_arabic(candidate), font=font)
                if piece and candidate_box[2] - candidate_box[0] > max_width:
                    fitted_words.append(piece)
                    piece = character
                else:
                    piece = candidate
            if piece:
                fitted_words.append(piece)
        lines = []
        line = ""
        for word in fitted_words:
            candidate = f"{line} {word}".strip()
            box = draw.textbbox((0, 0), shape_arabic(candidate), font=font)
            if line and box[2] - box[0] > max_width:
                lines.append(line)
                line = word
            else:
                line = candidate
        if line:
            lines.append(line)
        return lines

    def _export_png_one_page(self, records: list[dict[str, Any]], recipient: str, output: Path,
                             selected_fields: Optional[List[FieldDefinition]] = None) -> None:
        """ينشئ PNG عالي الدقة بمقاس ورق حقيقي؛ عند تعدد الصفحات تُجمع عمودياً دون تصغيرها."""
        self._require_arabic_shaping()
        if Image is None or ImageDraw is None or ImageFont is None:
            raise RuntimeError("Pillow غير مثبتة.")

        fields = list(reversed([f for f in (selected_fields if selected_fields else self.settings.report_fields) if f.key != "scheduled_delivery_date"]))
        palette = self._palette()
        page_size, _, _, orientation = self._resolve_report_page(fields)
        page_w_mm, page_h_mm = self._page_mm(page_size)
        dpi = max(150, min(600, int(self.settings.png_dpi)))

        def px(mm_value: float) -> int:
            return max(1, int(round(mm_value / 25.4 * dpi)))

        def font_px(points: float) -> int:
            return max(8, int(round(points / 72.0 * dpi)))

        width = px(page_w_mm)
        height = px(page_h_mm)
        side_margin = px(max(3.0, float(self.settings.report_margin_mm)))
        content_width = width - 2 * side_margin
        table_top = px(46)
        table_bottom = height - px(29)
        available_height = table_bottom - table_top

        raw_widths = report_column_weights(fields, records)
        width_factor = content_width / sum(raw_widths) if sum(raw_widths) else 1.0
        col_widths = [max(px(10), int(item_width * width_factor)) for item_width in raw_widths]
        if col_widths:
            col_widths[-1] += content_width - sum(col_widths)

        scratch = Image.new("RGB", (max(10, width), 100), "white")
        scratch_draw = ImageDraw.Draw(scratch)

        def font_at(points: float, bold: bool = False):
            path = self.bold_font_path if bold else self.regular_font_path
            path = path or self.regular_font_path or self.bold_font_path
            return ImageFont.truetype(path, font_px(points)) if path else ImageFont.load_default()

        border_width = max(1, px(0.22))
        inner_pad = px(1.4)

        def measure(page_records: list[dict[str, Any]], body_points: float):
            header_font = font_at(min(body_points + 0.9, 11.4), True)
            data_font = font_at(body_points, False)
            header_line_step = max(font_px(body_points + 1.4), int(font_px(body_points + 0.9) * 1.22))
            data_line_step = max(font_px(body_points + 1.1), int(font_px(body_points) * 1.22))
            header_lines = [
                self._wrap_text(
                    scratch_draw, item.label, header_font,
                    max(px(4), col_widths[index] - inner_pad * 2),
                )
                for index, item in enumerate(fields)
            ]
            header_height = max(
                px(8.0),
                max((len(lines) for lines in header_lines), default=1) * header_line_step + px(3.0),
            )
            rows = []
            for record in page_records:
                wrapped = [
                    self._wrap_text(
                        scratch_draw, value, data_font,
                        max(px(4), col_widths[index] - inner_pad * 2),
                    )
                    for index, value in enumerate(self._values(record, fields))
                ]
                row_height = max(
                    px(7.5),
                    max((len(lines) for lines in wrapped), default=1) * data_line_step + px(3.0),
                )
                rows.append((record, wrapped, row_height))
            if not page_records:
                rows.append(({}, [["لا توجد سجلات"]] + [[""] for _ in fields[1:]], px(9.0)))
            total_height = header_height + sum(row_height for _, _, row_height in rows)
            return {
                "body_points": body_points,
                "header_font": header_font,
                "data_font": data_font,
                "header_step": header_line_step,
                "data_step": data_line_step,
                "header_lines": header_lines,
                "header_height": header_height,
                "rows": rows,
                "total_height": total_height,
            }

        preferred_points = min(
            float(self.settings.pdf_base_font_size),
            max(float(self.settings.pdf_min_font_size), 11.2 - max(0, len(fields) - 6) * 0.22),
        )
        minimum_points = float(self.settings.pdf_min_font_size)

        def fit_chunk(page_records: list[dict[str, Any]]):
            candidate = preferred_points
            while candidate >= minimum_points - 0.01:
                spec = measure(page_records, candidate)
                if spec["total_height"] <= available_height:
                    return spec
                candidate -= 0.25
            return None

        page_specs: list[tuple[list[dict[str, Any]], dict[str, Any]]] = []
        if not records:
            spec = fit_chunk([])
            if not spec:
                raise RuntimeError("تعذر إنشاء صورة التقرير بالإعدادات الحالية.")
            page_specs.append(([], spec))
        else:
            cursor = 0
            while cursor < len(records):
                chunk, spec = largest_fitting_page_chunk(records[cursor:], fit_chunk)
                if not spec or not chunk:
                    raise RuntimeError(
                        "يوجد نص طويل جداً لا يمكن رسمه بصورة واضحة. اختر A3 أو قلل الحقول."
                    )
                page_specs.append((chunk, spec))
                cursor += len(chunk)

        now = dt.datetime.now()
        metrics = report_record_metrics(records)
        reference = report_reference(records, now)
        total_pages = len(page_specs)

        def render_page(page_records: list[dict[str, Any]], spec: dict[str, Any], page_number: int):
            image = Image.new("RGB", (width, height), palette["page_bg"])
            draw = ImageDraw.Draw(image)

            # رأس التقرير بمقاسات فعلية ثابتة على الورق مهما تغير DPI.
            draw.rectangle((0, 0, width, px(22)), fill=palette["title_fill"])
            draw.rectangle((0, px(22), width, px(35)), fill=palette["subtitle_fill"])
            draw.rectangle((0, px(35), width, px(36.3)), fill=palette["accent"])
            title_font = font_at(19.0 if str(self.settings.report_page_size).upper() == "A4" else 21.0, True)
            english_font = font_at(9.4, True)
            subtitle_font = font_at(10.0, False)
            meta_font = font_at(8.2, False)
            footer_font = font_at(7.6, False)

            draw.text((width//2, px(9.5)), shape_arabic(self.settings.report_title),
                      font=title_font, fill=palette["title_text"], anchor="mm")
            draw.text((width//2, px(17.5)), self.settings.report_title_en,
                      font=english_font, fill=palette["title_text"], anchor="mm")
            draw.text((width//2, px(28.5)),
                      shape_arabic(f"{self.settings.company_name} | {self.settings.report_subtitle}"),
                      font=subtitle_font, fill=palette["subtitle_text"], anchor="mm")

            if self.settings.logo_path and Path(self.settings.logo_path).is_file():
                try:
                    logo = Image.open(self.settings.logo_path).convert("RGBA")
                    resampling = getattr(Image, "Resampling", Image)
                    max_w = px(max(15, min(60, int(self.settings.logo_width_mm))))
                    max_h = px(25)
                    logo.thumbnail((max_w, max_h), resampling.LANCZOS)
                    image.paste(logo, (side_margin, px(5)), logo)
                except Exception:
                    logging.exception("تعذر إضافة الشعار إلى PNG")

            meta_top = px(38.5)
            meta_bottom = px(44.0)
            draw.rounded_rectangle(
                (side_margin, meta_top, width-side_margin, meta_bottom),
                radius=px(1.4), fill=palette["meta_fill"], outline=palette["border"], width=border_width,
            )
            meta = (
                f"السجلات: {metrics['records']} | العقود: {metrics['contracts']} | "
                f"العقود المتكررة: {metrics['multi_kitchen_contracts']} | "
                f"هذه الصفحة: {len(page_records)} | {now:%Y-%m-%d}"
            )
            draw.text((width//2, (meta_top+meta_bottom)//2), shape_arabic(meta),
                      font=meta_font, fill=palette["text"], anchor="mm")

            # الجدول
            y = table_top
            x = side_margin
            header_font = spec["header_font"]
            data_font = spec["data_font"]
            header_step = spec["header_step"]
            data_step = spec["data_step"]
            header_height = spec["header_height"]
            for column_index, lines in enumerate(spec["header_lines"]):
                right = x + col_widths[column_index]
                draw.rectangle((x, y, right, y+header_height), fill=palette["header_fill"],
                               outline=palette["border"], width=border_width)
                line_y = y + header_height/2 - ((len(lines)-1) * header_step/2)
                for line in lines:
                    draw.text(((x+right)/2, line_y), shape_arabic(line), font=header_font,
                              fill=palette["header_text"], anchor="mm")
                    line_y += header_step
                x = right
            y += header_height

            status_index = next((i for i, item in enumerate(fields) if item.key == "delivery_status"), None)
            for row_index, (record, wrapped, row_height) in enumerate(spec["rows"]):
                x = side_margin
                default_fill = "#FFFFFF" if row_index % 2 == 0 else palette["row_alt"]
                status = clean_value(record.get("values", {}).get("delivery_status", "")) if record else ""
                status_fill = REPORT_STATUS_FILLS.get(status)
                for column_index, lines in enumerate(wrapped):
                    right = x + col_widths[column_index]
                    fill = default_fill
                    if status_index == column_index and status_fill:
                        fill = f"#{status_fill}"
                    draw.rectangle((x, y, right, y+row_height), fill=fill,
                                   outline=palette["border"], width=border_width)
                    line_y = y + row_height/2 - ((len(lines)-1) * data_step/2)
                    for line in lines:
                        draw.text(((x+right)/2, line_y), shape_arabic(line), font=data_font,
                                  fill=palette["text"], anchor="mm")
                        line_y += data_step
                    x = right
                y += row_height

            # التوقيع والتذييل
            signature_font = font_at(8.3, False)
            signature_y = height - px(21.0)
            signature_h = px(8.0)
            signature_items = [
                f"اسم المستلم: {clean_value(recipient) or '........................'}",
                "التوقيع: ........................",
                f"التاريخ: {now:%Y-%m-%d}",
            ]
            signature_width = content_width / 3
            for index, item in enumerate(signature_items):
                left = side_margin + index * signature_width
                right = side_margin + (index + 1) * signature_width
                draw.rectangle((left, signature_y, right, signature_y+signature_h),
                               fill="#FFFFFF", outline=palette["border"], width=border_width)
                draw.text(((left+right)/2, signature_y+signature_h/2), shape_arabic(item),
                          font=signature_font, fill=palette["text"], anchor="mm")

            line_y = height - px(10.5)
            draw.line((side_margin, line_y, width-side_margin, line_y), fill=palette["accent"], width=max(1, px(0.35)))
            draw.text((width-side_margin, height-px(5.5)), shape_arabic(self.settings.footer_note),
                      font=footer_font, fill=palette["muted"], anchor="rm")
            draw.text((side_margin, height-px(5.5)), shape_arabic(f"صفحة {page_number} من {total_pages}"),
                      font=footer_font, fill=palette["muted"], anchor="lm")
            draw.text((width//2, height-px(5.5)),
                      f"REF: {reference} | {page_w_mm:.0f}x{page_h_mm:.0f} mm | {dpi} DPI | {orientation.upper()}",
                      font=footer_font, fill=palette["muted"], anchor="mm")
            return image

        output.parent.mkdir(parents=True, exist_ok=True)
        generated_paths: list[Path] = []
        for index, (chunk, spec) in enumerate(page_specs):
            page = render_page(chunk, spec, index + 1)
            page_output = output if index == 0 else output.with_name(
                f"{output.stem}_صفحة_{index + 1:02d}{output.suffix or '.png'}"
            )
            page.save(page_output, format="PNG", optimize=True, dpi=(dpi, dpi))
            generated_paths.append(page_output)
        return generated_paths

    def export_png(self, records: list[dict[str, Any]], recipient: str, output: Path,
                   selected_fields: Optional[List[FieldDefinition]] = None):
        return self._export_png_one_page(records, recipient, output, selected_fields)

    def export_jpg(self, records: list[dict[str, Any]], recipient: str, output: Path,
                   selected_fields: Optional[List[FieldDefinition]] = None):
        if Image is None:
            raise RuntimeError("Pillow غير مثبتة.")
        with tempfile.TemporaryDirectory(prefix="marble_jpg_") as tmp_name:
            temp_output = Path(tmp_name) / "page.png"
            png_paths = self.export_png(records, recipient, temp_output, selected_fields)
            generated: list[Path] = []
            for index, png_path in enumerate(png_paths):
                target = output if index == 0 else output.with_name(f"{output.stem}_صفحة_{index + 1:02d}{output.suffix or '.jpg'}")
                target.parent.mkdir(parents=True, exist_ok=True)
                with Image.open(png_path) as image:
                    rgb = image.convert("RGB")
                    rgb.save(target, format="JPEG", quality=max(70, min(100, int(self.settings.jpg_quality))),
                             optimize=True, dpi=(self.settings.png_dpi, self.settings.png_dpi))
                generated.append(target)
            return generated

    def export_csv(self, records: list[dict[str, Any]], recipient: str, output: Path,
                   selected_fields: Optional[List[FieldDefinition]] = None) -> None:
        fields = [f for f in (selected_fields if selected_fields else self.settings.report_fields) if f.key != "scheduled_delivery_date"]
        DatabaseManager.export_to_csv(records, fields, output)

    # --------------------- دوال مساعدة للخطوط ---------------------
    def _register_pdf_fonts(self) -> tuple[str, str]:
        if colors is None:
            detail = f" ({REPORTLAB_IMPORT_ERROR})" if REPORTLAB_IMPORT_ERROR else ""
            raise RuntimeError(f"reportlab غير مثبتة أو غير مكتملة{detail}.")
        if not self.regular_font_path:
            raise RuntimeError("لم يتم العثور على خط عربي.")
        regular_name = "MarbleArabic"
        bold_name = "MarbleArabicBold"
        if regular_name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(regular_name, self.regular_font_path))
        if bold_name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(bold_name, self.bold_font_path or self.regular_font_path))
        return regular_name, bold_name

    @staticmethod
    def _require_arabic_shaping():
        if arabic_reshaper is None or get_display is None:
            raise RuntimeError("يلزم تثبيت arabic-reshaper وpython-bidi.")

# ============================================================================
# منتقي تاريخ عربي دون اعتماد خارجي
# ============================================================================

class LoginDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.result: Optional[dict[str, Any]] = None
        self.title("تسجيل الدخول")
        self.geometry("420x255")
        self.resizable(False, False)
        self.transient(parent); self.grab_set()
        self.configure(bg="#F4F7F9")
        tk.Label(self, text="تسجيل الدخول إلى نظام تسليم الرخام", bg="#F4F7F9", fg="#17324D", font=("Segoe UI", 15, "bold")).pack(pady=(22, 15))
        frame = tk.Frame(self, bg="#F4F7F9"); frame.pack(fill="x", padx=35)
        self.user_var = tk.StringVar(value="admin"); self.pass_var = tk.StringVar()
        for row, (label, var, show) in enumerate((("اسم المستخدم", self.user_var, ""), ("كلمة المرور", self.pass_var, "*"))):
            tk.Label(frame, text=label, bg="#F4F7F9", anchor="e").grid(row=row, column=1, sticky="e", padx=8, pady=7)
            entry = tk.Entry(frame, textvariable=var, justify="right", show=show, font=("Segoe UI", 11)); entry.grid(row=row, column=0, sticky="ew", pady=7, ipady=5)
            if row == 1: entry.bind("<Return>", lambda _e: self._login())
        frame.grid_columnconfigure(0, weight=1)
        buttons = tk.Frame(self, bg="#F4F7F9"); buttons.pack(fill="x", padx=35, pady=18)
        tk.Button(buttons, text="دخول", command=self._login, bg="#2B8C5E", fg="white", relief="flat", width=14, pady=7).pack(side="right", padx=5)
        tk.Button(buttons, text="إلغاء", command=self.destroy, bg="#6C757D", fg="white", relief="flat", width=12, pady=7).pack(side="right", padx=5)
    def _login(self):
        user = DatabaseManager.authenticate(self.user_var.get(), self.pass_var.get())
        if not user:
            messagebox.showerror("تسجيل الدخول", "اسم المستخدم أو كلمة المرور غير صحيحة.", parent=self); return
        self.result = user; self.destroy()

class UserManagementDialog(tk.Toplevel):
    def __init__(self, app: "MarbleDeliveryApp"):
        super().__init__(app.root); self.app = app
        self.title("المستخدمون والصلاحيات"); self.geometry("840x560"); self.transient(app.root); self.grab_set()
        self.configure(bg=app.settings.background_color)
        self.tree = ttk.Treeview(self, columns=("username","name","role","active"), show="headings", style="Modern.Treeview")
        for key, label, width in (("username","اسم المستخدم",160),("name","الاسم الظاهر",220),("role","الدور",130),("active","الحالة",100)):
            self.tree.heading(key,text=label); self.tree.column(key,width=width,anchor="center")
        self.tree.pack(fill="both", expand=True, padx=16, pady=(16,8)); self.tree.bind("<<TreeviewSelect>>", self._select)
        form=tk.Frame(self,bg=app.settings.surface_color); form.pack(fill="x",padx=16,pady=8)
        self.username=tk.StringVar(); self.display=tk.StringVar(); self.role=tk.StringVar(value="موظف"); self.password=tk.StringVar(); self.active=tk.BooleanVar(value=True); self.user_id=""
        for i,(label,var) in enumerate((("اسم المستخدم",self.username),("الاسم الظاهر",self.display),("كلمة مرور جديدة",self.password))):
            tk.Label(form,text=label,bg=app.settings.surface_color).grid(row=0,column=i*2+1,sticky="e",padx=5,pady=8)
            tk.Entry(form,textvariable=var,show="*" if "مرور" in label else "",justify="right",width=18).grid(row=0,column=i*2,sticky="ew",padx=5,pady=8,ipady=4)
        tk.Label(form,text="الدور",bg=app.settings.surface_color).grid(row=1,column=5,sticky="e",padx=5,pady=8)
        ttk.Combobox(form,textvariable=self.role,values=tuple(ROLE_PERMISSIONS),state="readonly",width=16).grid(row=1,column=4,padx=5,pady=8)
        tk.Checkbutton(form,text="نشط",variable=self.active,bg=app.settings.surface_color).grid(row=1,column=2,padx=5,pady=8)
        buttons=tk.Frame(self,bg=app.settings.background_color); buttons.pack(fill="x",padx=16,pady=(4,14))
        app.make_button(buttons,"حفظ المستخدم",self._save,app.settings.success_color,width=14,pady=7).pack(side="right",padx=4)
        app.make_button(buttons,"مستخدم جديد",self._new,app.settings.primary_color,width=14,pady=7).pack(side="right",padx=4)
        app.make_button(buttons,"حذف",self._delete,app.settings.danger_color,width=10,pady=7).pack(side="right",padx=4)
        tk.Label(buttons,text="ملاحظة: ترك كلمة المرور فارغة يبقي كلمة المرور الحالية.",bg=app.settings.background_color,fg=app.settings.muted_color).pack(side="left")
        self._refresh()
    def _refresh(self):
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for u in DatabaseManager.list_users():
            self.tree.insert("", "end", iid=u["id"], values=(u["username"],u["display_name"],u["role"],"نشط" if u["active"] else "موقوف"))
    def _select(self,_e=None):
        sel=self.tree.selection()
        if not sel:return
        u=next((x for x in DatabaseManager.list_users() if x["id"]==sel[0]),None)
        if not u:return
        self.user_id=u["id"]; self.username.set(u["username"]); self.display.set(u["display_name"]); self.role.set(u["role"]); self.password.set(""); self.active.set(u["active"])
    def _new(self): self.user_id=""; self.username.set(""); self.display.set(""); self.role.set("موظف"); self.password.set(""); self.active.set(True)
    def _save(self):
        try:
            DatabaseManager.save_user(self.username.get(),self.display.get(),self.role.get(),self.password.get(),user_id=self.user_id,active=self.active.get()); self._new(); self._refresh(); self.app.status_var.set("تم تحديث المستخدمين والصلاحيات")
        except Exception as exc: messagebox.showerror("المستخدمون",str(exc),parent=self)
    def _delete(self):
        if not self.user_id:return
        if messagebox.askyesno("حذف مستخدم","حذف المستخدم المحدد؟",parent=self):
            try: DatabaseManager.delete_user(self.user_id); self._new(); self._refresh()
            except Exception as exc: messagebox.showerror("المستخدمون",str(exc),parent=self)

class SwitchUserDialog(tk.Toplevel):
    def __init__(self, app: "MarbleDeliveryApp"):
        super().__init__(app.root); self.app=app; self.title("تبديل المستخدم"); self.geometry("430x260"); self.transient(app.root); self.grab_set(); self.configure(bg=app.settings.background_color)
        users=DatabaseManager.list_users(active_only=True); self.lookup={u["display_name"]:u["username"] for u in users}
        self.display=tk.StringVar(value=next(iter(self.lookup),"")); self.password=tk.StringVar()
        tk.Label(self,text="تبديل المستخدم الحالي",bg=app.settings.background_color,fg=app.settings.primary_color,font=(app.settings.font_family,15,"bold")).pack(pady=(25,15))
        ttk.Combobox(self,textvariable=self.display,values=tuple(self.lookup),state="readonly",width=30).pack(pady=7,ipady=4)
        tk.Entry(self,textvariable=self.password,show="*",justify="center",width=32).pack(pady=7,ipady=5)
        app.make_button(self,"دخول",self._login,app.settings.success_color,width=14,pady=7).pack(pady=14)
    def _login(self):
        username=self.lookup.get(self.display.get(),""); user=DatabaseManager.authenticate(username,self.password.get())
        if not user: messagebox.showerror("تبديل المستخدم","كلمة المرور غير صحيحة.",parent=self); return
        self.app.set_current_user(user); self.destroy()

class TrashDialog(tk.Toplevel):
    def __init__(self, app: "MarbleDeliveryApp"):
        super().__init__(app.root); self.app=app; self.title("سلة المحذوفات"); self.geometry("980x580"); self.transient(app.root); self.grab_set(); self.configure(bg=app.settings.background_color)
        self.tree=ttk.Treeview(self,columns=("date","contract","customer","status"),show="headings",selectmode="extended",style="Modern.Treeview")
        for key,label,width in (("date","تاريخ الحذف",180),("contract","العقد",180),("customer","العميل",260),("status","الحالة",180)):
            self.tree.heading(key,text=label); self.tree.column(key,width=width,anchor="center")
        self.tree.pack(fill="both",expand=True,padx=16,pady=16)
        buttons=tk.Frame(self,bg=app.settings.background_color); buttons.pack(fill="x",padx=16,pady=(0,16))
        app.make_button(buttons,"استعادة المحدد",self._restore,app.settings.success_color,width=16,pady=7).pack(side="right",padx=4)
        app.make_button(buttons,"حذف نهائي",self._purge,app.settings.danger_color,width=14,pady=7).pack(side="right",padx=4)
        app.make_button(buttons,"تحديث",self._refresh,app.settings.primary_color,width=10,pady=7).pack(side="right",padx=4)
        self._refresh()
    def _refresh(self):
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for item in DatabaseManager.list_deleted_records():
            values=item.get("record",{}).get("values",{}); self.tree.insert("","end",iid=item["id"],values=(item["deleted_at"],values.get("contract_number",""),values.get("customer_name",""),values.get("delivery_status","")))
    def _restore(self):
        ids=list(self.tree.selection());
        if not ids:return
        count=DatabaseManager.restore_deleted_records(ids); self.app.records=DatabaseManager.load_records(); self.app._apply_filters(); self._refresh(); self.app.status_var.set(f"تمت استعادة {count} سجل")
    def _purge(self):
        ids=list(self.tree.selection());
        if not ids:return
        if messagebox.askyesno("حذف نهائي","سيتم حذف السجلات والمرفقات نهائياً. متابعة؟",parent=self): DatabaseManager.purge_deleted_records(ids); self._refresh()

class DirectoryDialog(tk.Toplevel):
    TYPES={"العملاء":"customer","المعارض":"showroom","موظفو المعارض":"salesperson"}
    def __init__(self, app: "MarbleDeliveryApp"):
        super().__init__(app.root); self.app=app; self.title("دليل العملاء والمعارض"); self.geometry("930x600"); self.transient(app.root); self.grab_set(); self.configure(bg=app.settings.background_color)
        top=tk.Frame(self,bg=app.settings.background_color); top.pack(fill="x",padx=16,pady=12)
        self.type_label=tk.StringVar(value="العملاء"); ttk.Combobox(top,textvariable=self.type_label,values=tuple(self.TYPES),state="readonly",width=18).pack(side="right",padx=5); self.type_label.trace_add("write",lambda *_:self._refresh())
        app.make_button(top,"استخراج من السجلات",self._learn,app.settings.primary_color,width=18,pady=6).pack(side="right",padx=5)
        self.tree=ttk.Treeview(self,columns=("name","mobile","details"),show="headings",style="Modern.Treeview")
        for k,l,w in (("name","الاسم",260),("mobile","الجوال",180),("details","ملاحظات",360)):
            self.tree.heading(k,text=l); self.tree.column(k,width=w,anchor="center")
        self.tree.pack(fill="both",expand=True,padx=16,pady=8); self.tree.bind("<<TreeviewSelect>>",self._select)
        form=tk.Frame(self,bg=app.settings.surface_color); form.pack(fill="x",padx=16,pady=8)
        self.name=tk.StringVar(); self.mobile=tk.StringVar(); self.details=tk.StringVar()
        for i,(label,var) in enumerate((("الاسم",self.name),("الجوال",self.mobile),("ملاحظات",self.details))):
            tk.Label(form,text=label,bg=app.settings.surface_color).grid(row=0,column=i*2+1,sticky="e",padx=5,pady=8); tk.Entry(form,textvariable=var,justify="right").grid(row=0,column=i*2,sticky="ew",padx=5,pady=8,ipady=5); form.grid_columnconfigure(i*2,weight=1)
        buttons=tk.Frame(self,bg=app.settings.background_color); buttons.pack(fill="x",padx=16,pady=(2,14))
        app.make_button(buttons,"حفظ",self._save,app.settings.success_color,width=12,pady=7).pack(side="right",padx=4); app.make_button(buttons,"حذف",self._delete,app.settings.danger_color,width=10,pady=7).pack(side="right",padx=4)
        self._refresh()
    def _type(self): return self.TYPES.get(self.type_label.get(),"customer")
    def _refresh(self):
        if not hasattr(self,"tree"):return
        for iid in self.tree.get_children():self.tree.delete(iid)
        for item in DatabaseManager.list_directory(self._type()): self.tree.insert("","end",iid=str(item["id"]),values=(item["name"],item["mobile"],item["details"]))
    def _select(self,_e=None):
        sel=self.tree.selection();
        if not sel:return
        vals=self.tree.item(sel[0],"values"); self.name.set(vals[0] if vals else ""); self.mobile.set(vals[1] if len(vals)>1 else ""); self.details.set(vals[2] if len(vals)>2 else "")
    def _save(self):
        try: DatabaseManager.save_directory(self._type(),self.name.get(),self.mobile.get(),self.details.get()); self._refresh(); self.app._refresh_directory_values(); self.name.set(""); self.mobile.set(""); self.details.set("")
        except Exception as exc: messagebox.showerror("الدليل",str(exc),parent=self)
    def _delete(self):
        if self.name.get() and messagebox.askyesno("حذف","حذف العنصر من الدليل؟",parent=self): DatabaseManager.delete_directory(self._type(),self.name.get()); self._refresh(); self.app._refresh_directory_values()
    def _learn(self): DatabaseManager.learn_directories_from_records(); self._refresh(); self.app._refresh_directory_values(); self.app.status_var.set("تم تحديث الأدلة من السجلات الحالية")

class CustomerHistoryDialog(tk.Toplevel):
    def __init__(self, app: "MarbleDeliveryApp"):
        super().__init__(app.root)
        self.app = app
        self.title("سجل تسليمات العميل")
        self.geometry("1050x620")
        self.transient(app.root); self.grab_set()
        self.configure(bg=app.settings.background_color)
        names = sorted({clean_value(r.get("values", {}).get("customer_name", "")) for r in DatabaseManager.load_records() if clean_value(r.get("values", {}).get("customer_name", ""))})
        names = sorted(set(names) | {item["name"] for item in DatabaseManager.list_directory("customer")})
        top = tk.Frame(self, bg=app.settings.background_color); top.pack(fill="x", padx=16, pady=14)
        self.customer_var = tk.StringVar(value=names[0] if names else "")
        tk.Label(top, text="العميل", bg=app.settings.background_color, fg=app.settings.primary_color, font=(app.settings.font_family, 11, "bold")).pack(side="right", padx=6)
        combo = ttk.Combobox(top, textvariable=self.customer_var, values=names, state="normal", justify="right", width=34)
        combo.pack(side="right", padx=5, ipady=4); combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh()); combo.bind("<Return>", lambda _e: self._refresh())
        self.summary = tk.StringVar()
        tk.Label(top, textvariable=self.summary, bg="#EAF2F8", fg=app.settings.primary_color, padx=10, pady=6, font=(app.settings.font_family, 9, "bold")).pack(side="left", padx=4)
        self.tree = ttk.Treeview(self, columns=("contract","kitchen","status","showroom","scheduled","actual","updated"), show="headings", style="Modern.Treeview")
        specs=(("contract","العقد",150),("kitchen","المطبخ",90),("status","الحالة",190),("showroom","المعرض",150),("scheduled","موعد التسليم",120),("actual","التركيب الفعلي",120),("updated","آخر تعديل",165))
        for key,label,width in specs:
            self.tree.heading(key,text=label); self.tree.column(key,width=width,anchor="center")
        self.tree.pack(fill="both",expand=True,padx=16,pady=8)
        buttons=tk.Frame(self,bg=app.settings.background_color); buttons.pack(fill="x",padx=16,pady=(2,14))
        app.make_button(buttons,"عرض العميل في مساحة العمل",self._show_in_app,app.settings.primary_color,width=22,pady=7).pack(side="right",padx=4)
        app.make_button(buttons,"إغلاق",self.destroy,app.settings.muted_color,width=10,pady=7).pack(side="right",padx=4)
        self._refresh()
    def _matching(self):
        target=normalize_search_text(self.customer_var.get())
        if not target:return []
        return [r for r in DatabaseManager.load_records() if normalize_search_text(r.get("values",{}).get("customer_name","")) == target]
    def _refresh(self):
        if not hasattr(self,"tree"):return
        for iid in self.tree.get_children():self.tree.delete(iid)
        records=self._matching()
        for rec in records:
            v=rec.get("values",{})
            self.tree.insert("","end",iid=str(rec.get("id","")),values=(v.get("contract_number",""),v.get("kitchen_number",""),v.get("delivery_status",""),v.get("showroom",""),v.get("scheduled_delivery_date",""),v.get("actual_delivery_date",""),rec.get("updated_at","")))
        delivered=sum(1 for r in records if clean_value(r.get("values",{}).get("delivery_status",""))=="تم التسليم")
        overdue=sum(1 for r in records if record_is_overdue(r))
        self.summary.set(f"السجلات: {len(records)} | تم التسليم: {delivered} | متأخرة: {overdue}")
    def _show_in_app(self):
        name=clean_value(self.customer_var.get())
        if not name:return
        self.app.show_all_records(); self.app.search_var.set(name); self.app._apply_filters(); self.app.status_var.set(f"عرض سجل العميل: {name}"); self.destroy()

class ReportTemplateDialog(tk.Toplevel):
    def __init__(self, app: "MarbleDeliveryApp"):
        super().__init__(app.root); self.app=app; self.title("قوالب التقارير"); self.geometry("680x480"); self.transient(app.root); self.grab_set(); self.configure(bg=app.settings.background_color)
        self.tree=ttk.Treeview(self,columns=("name","updated"),show="headings",style="Modern.Treeview"); self.tree.heading("name",text="اسم القالب"); self.tree.heading("updated",text="آخر تعديل"); self.tree.column("name",width=320,anchor="center"); self.tree.column("updated",width=220,anchor="center"); self.tree.pack(fill="both",expand=True,padx=16,pady=16)
        self.name=tk.StringVar(); form=tk.Frame(self,bg=app.settings.background_color); form.pack(fill="x",padx=16,pady=(0,8)); tk.Entry(form,textvariable=self.name,justify="right").pack(side="right",fill="x",expand=True,ipady=5); tk.Label(form,text="اسم القالب",bg=app.settings.background_color).pack(side="right",padx=8)
        buttons=tk.Frame(self,bg=app.settings.background_color); buttons.pack(fill="x",padx=16,pady=(0,16)); app.make_button(buttons,"حفظ الإعدادات الحالية",self._save,app.settings.success_color,width=20,pady=7).pack(side="right",padx=4); app.make_button(buttons,"تطبيق القالب",self._apply,app.settings.primary_color,width=16,pady=7).pack(side="right",padx=4); app.make_button(buttons,"حذف",self._delete,app.settings.danger_color,width=10,pady=7).pack(side="right",padx=4); self._refresh()
    def _refresh(self):
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for item in DatabaseManager.list_report_templates(): self.tree.insert("","end",iid=item["id"],values=(item["name"],item["updated_at"]))
    def _save(self):
        try: DatabaseManager.save_report_template(self.name.get(),self.app.settings); self.name.set(""); self._refresh()
        except Exception as exc: messagebox.showerror("قوالب التقارير",str(exc),parent=self)
    def _selected(self):
        sel=self.tree.selection(); return next((x for x in DatabaseManager.list_report_templates() if sel and x["id"]==sel[0]),None)
    def _apply(self):
        item=self._selected();
        if not item:return
        payload=item.get("payload",{})
        for key,value in payload.items():
            if hasattr(self.app.settings,key): setattr(self.app.settings,key,copy.deepcopy(value))
        DatabaseManager.save_settings(self.app.settings); self.app._configure_styles(); self.app._build_ui(); self.app._apply_filters(); self.app.status_var.set(f"تم تطبيق قالب التقرير: {item['name']}"); self.destroy()
    def _delete(self):
        item=self._selected();
        if item and messagebox.askyesno("حذف قالب",f"حذف القالب «{item['name']}»؟",parent=self): DatabaseManager.delete_report_template(item["id"]); self._refresh()

class BulkEditDialog(tk.Toplevel):
    def __init__(self, app: "MarbleDeliveryApp", ids: list[str]):
        super().__init__(app.root); self.app=app; self.ids=ids; self.title("تعديل جماعي"); self.geometry("540x300"); self.transient(app.root); self.grab_set(); self.configure(bg=app.settings.background_color)
        self.field_label=tk.StringVar(value=app.settings.visible_fields[0].label if app.settings.visible_fields else ""); self.value=tk.StringVar(); self.clear=tk.BooleanVar(value=False)
        tk.Label(self,text=f"تعديل {len(ids)} سجل محدد",bg=app.settings.background_color,fg=app.settings.primary_color,font=(app.settings.font_family,14,"bold")).pack(pady=(25,15))
        ttk.Combobox(self,textvariable=self.field_label,values=tuple(f.label for f in app.settings.visible_fields),state="readonly",width=34).pack(pady=7,ipady=4)
        tk.Entry(self,textvariable=self.value,justify="right",width=38).pack(pady=7,ipady=5)
        tk.Checkbutton(self,text="إفراغ الحقل في السجلات المحددة",variable=self.clear,bg=app.settings.background_color).pack(pady=5)
        app.make_button(self,"تطبيق التعديل",self._apply,app.settings.success_color,width=16,pady=7).pack(pady=14)
    def _apply(self):
        field=next((f for f in self.app.settings.visible_fields if f.label==self.field_label.get()),None)
        if not field:return
        value="" if self.clear.get() else clean_value(self.value.get())
        try:
            records={r["id"]:r for r in DatabaseManager.load_records()}
            for rid in self.ids:
                rec=records.get(rid)
                if not rec: continue
                rec["values"][field.key]=value; rec["updated_at"]=dt.datetime.now().isoformat(timespec="seconds"); DatabaseManager.save_record(rec)
            self.app.records=DatabaseManager.load_records(); self.app._apply_filters(); DatabaseManager.log_action("تعديل جماعي",details=f"{field.label} | {len(self.ids)} سجل"); self.app.status_var.set(f"تم تعديل {len(self.ids)} سجل"); self.destroy()
        except Exception as exc: messagebox.showerror("التعديل الجماعي",str(exc),parent=self)

class DatePickerDialog(tk.Toplevel):
    MONTHS = [
        "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
        "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر",
    ]
    WEEKDAYS = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]

    def __init__(self, parent: tk.Misc, current_value: str = "", title: str = "اختيار التاريخ",
                 settings: Optional[AppSettings] = None):
        super().__init__(parent)
        self.result: Optional[str] = None
        self.settings = settings or AppSettings()
        try:
            selected = dt.datetime.strptime(clean_value(current_value), "%Y-%m-%d").date()
        except ValueError:
            selected = dt.date.today()
        self.selected_date = selected
        self.year = selected.year
        self.month = selected.month
        self.title(title)
        self.geometry("470x430")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.grab_set()
        self.configure(bg=self.settings.background_color)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())

        heading = tk.Label(
            self, text=title, bg=self.settings.primary_color, fg="white", pady=12,
            font=(self.settings.font_family, 14, "bold"),
        )
        heading.pack(fill="x")
        nav = tk.Frame(self, bg=self.settings.surface_color)
        nav.pack(fill="x", padx=12, pady=(12, 4))
        tk.Button(
            nav, text="▶", command=lambda: self._change_month(-1), relief="flat",
            bg=self.settings.primary_color, fg="white", width=5, cursor="hand2",
        ).pack(side="right", padx=4)
        self.month_label = tk.Label(
            nav, bg=self.settings.surface_color, fg=self.settings.primary_color,
            font=(self.settings.font_family, 14, "bold"),
        )
        self.month_label.pack(side="right", fill="x", expand=True)
        tk.Button(
            nav, text="◀", command=lambda: self._change_month(1), relief="flat",
            bg=self.settings.primary_color, fg="white", width=5, cursor="hand2",
        ).pack(side="left", padx=4)

        self.days_frame = tk.Frame(self, bg=self.settings.surface_color)
        self.days_frame.pack(fill="both", expand=True, padx=12, pady=6)
        for column in range(7):
            self.days_frame.grid_columnconfigure(column, weight=1, uniform="calendar")

        actions = tk.Frame(self, bg=self.settings.background_color)
        actions.pack(fill="x", padx=12, pady=(4, 12))
        tk.Button(
            actions, text="اليوم", command=self._choose_today, relief="flat", cursor="hand2",
            bg=self.settings.success_color, fg="white", padx=18, pady=7,
            font=(self.settings.font_family, 10, "bold"),
        ).pack(side="right", padx=4)
        tk.Button(
            actions, text="مسح التاريخ", command=self._clear, relief="flat", cursor="hand2",
            bg=self.settings.muted_color, fg="white", padx=18, pady=7,
            font=(self.settings.font_family, 10, "bold"),
        ).pack(side="left", padx=4)
        self._render_month()

    def _change_month(self, delta: int):
        month_index = self.year * 12 + self.month - 1 + delta
        self.year, month_zero = divmod(month_index, 12)
        self.month = month_zero + 1
        self._render_month()

    def _render_month(self):
        for child in self.days_frame.winfo_children():
            child.destroy()
        self.month_label.configure(text=f"{self.MONTHS[self.month - 1]} {self.year}")
        for index, label in enumerate(self.WEEKDAYS):
            tk.Label(
                self.days_frame, text=label, bg=self.settings.surface_color,
                fg=self.settings.text_color, font=(self.settings.font_family, 9, "bold"),
            ).grid(row=0, column=6-index, sticky="nsew", padx=1, pady=3)
        today = dt.date.today()
        for week_index, week in enumerate(calendar.monthcalendar(self.year, self.month), 1):
            self.days_frame.grid_rowconfigure(week_index, weight=1, uniform="calendar_row")
            for day_index, day_number in enumerate(week):
                column = 6 - day_index
                if not day_number:
                    tk.Label(self.days_frame, bg=self.settings.surface_color).grid(
                        row=week_index, column=column, sticky="nsew", padx=2, pady=2
                    )
                    continue
                date_value = dt.date(self.year, self.month, day_number)
                is_selected = date_value == self.selected_date
                is_today = date_value == today
                bg = self.settings.accent_color if is_selected else ("#DCEFE4" if is_today else "#F7F9FB")
                fg = "#1A1A1A"
                tk.Button(
                    self.days_frame, text=str(day_number), command=lambda value=date_value: self._choose(value),
                    relief="flat", cursor="hand2", bg=bg, fg=fg, activebackground=self.settings.accent_color,
                    font=(self.settings.font_family, 11, "bold" if is_selected or is_today else "normal"),
                ).grid(row=week_index, column=column, sticky="nsew", padx=2, pady=2)

    def _choose(self, value: dt.date):
        self.result = value.isoformat()
        self.destroy()

    def _choose_today(self):
        self._choose(dt.date.today())

    def _clear(self):
        self.result = ""
        self.destroy()


# ============================================================================
# نافذة خيارات التصدير المتقدمة
# ============================================================================
class ExportOptionsDialog(tk.Toplevel):
    def __init__(self, parent: "MarbleDeliveryApp", export_callback: Callable):
        super().__init__(parent.root)
        self.parent_app = parent
        self.export_callback = export_callback
        self.title("خيارات التصدير المتقدمة")
        screen_w = max(900, parent.root.winfo_screenwidth())
        screen_h = max(700, parent.root.winfo_screenheight())
        width = min(980, max(860, int(screen_w * 0.70)))
        height = min(780, max(690, int(screen_h * 0.78)))
        x = max(0, parent.root.winfo_x() + (parent.root.winfo_width() - width) // 2)
        y = max(0, parent.root.winfo_y() + (parent.root.winfo_height() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(min(840, screen_w - 40), min(650, screen_h - 60))
        self.transient(parent.root)
        self.grab_set()
        self.configure(bg=parent.settings.background_color)

        self.settings = parent.settings
        self.selected_fields = tk.StringVar(value="الكل")
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()
        self.include_time_var = tk.BooleanVar(value=False)
        self.print_direct_var = tk.BooleanVar(value=False)
        self.open_after_export_var = tk.BooleanVar(value=True)
        self.file_name_prefix_var = tk.StringVar(value="تقرير تسليم الرخام")
        self.export_format_var = tk.StringVar(value="pdf")

        prefs = self.settings.export_preferences
        self.selected_fields.set(prefs.get("selected_fields", "الكل"))
        self.include_time_var.set(prefs.get("include_time", False))
        self.print_direct_var.set(prefs.get("print_direct", False))
        self.open_after_export_var.set(prefs.get("open_after_export", True))
        self.file_name_prefix_var.set(prefs.get("file_name_prefix", "تقرير تسليم الرخام"))
        self.export_format_var.set(prefs.get("export_format", "pdf"))
        self.saved_field_keys = set(prefs.get("selected_field_keys", []))

        self._build()
        self.wait_visibility()
        self.focus_force()

    def _build(self):
        main_frame = tk.Frame(self, bg=self.settings.background_color)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(main_frame, text="الحقول المراد تضمينها:", bg=self.settings.background_color,
                 font=(self.settings.font_family, 12, "bold")).pack(anchor="w", pady=(0, 5))
        field_frame = tk.Frame(main_frame, bg=self.settings.surface_color, relief="solid", bd=1)
        field_frame.pack(fill="x", pady=(0, 15))
        all_radio = tk.Radiobutton(field_frame, text="حقول التقرير الافتراضية الخفيفة", variable=self.selected_fields,
                                   value="الكل", bg=self.settings.surface_color,
                                   font=(self.settings.font_family, 11))
        all_radio.pack(anchor="w", padx=10, pady=5)
        custom_radio = tk.Radiobutton(field_frame, text="اختيار حقول محددة", variable=self.selected_fields,
                                      value="مخصص", bg=self.settings.surface_color,
                                      font=(self.settings.font_family, 11))
        custom_radio.pack(anchor="w", padx=10, pady=5)

        self.field_checkboxes = {}
        fields_frame = tk.Frame(field_frame, bg=self.settings.surface_color)
        fields_frame.pack(fill="x", padx=20, pady=5)
        exportable_fields = [f for f in self.settings.visible_fields if f.key != "scheduled_delivery_date"]
        for idx, f in enumerate(exportable_fields):
            var = tk.BooleanVar(value=(
                f.key in self.saved_field_keys
                if self.saved_field_keys else f.key != "kitchen_number"
            ))
            cb = tk.Checkbutton(fields_frame, text=f.label, variable=var,
                                bg=self.settings.surface_color,
                                font=(self.settings.font_family, 10))
            cb.grid(row=idx//4, column=idx%4, sticky="w", padx=5, pady=2)
            self.field_checkboxes[f.key] = var

        date_frame = tk.LabelFrame(main_frame, text="نطاق التواريخ داخل نتائج الفلترة الحالية", bg=self.settings.surface_color,
                                   font=(self.settings.font_family, 11, "bold"))
        date_frame.pack(fill="x", pady=10)
        tk.Label(date_frame, text="من تاريخ:", bg=self.settings.surface_color).pack(side="right", padx=5)
        from_entry = tk.Entry(date_frame, textvariable=self.date_from_var, width=14, justify="right",
                              font=(self.settings.font_family, 11))
        from_entry.pack(side="right", padx=5)
        tk.Label(date_frame, text="إلى تاريخ:", bg=self.settings.surface_color).pack(side="right", padx=5)
        to_entry = tk.Entry(date_frame, textvariable=self.date_to_var, width=14, justify="right",
                            font=(self.settings.font_family, 11))
        to_entry.pack(side="right", padx=5)

        def set_today():
            today = dt.date.today().isoformat()
            if not self.date_from_var.get():
                self.date_from_var.set(today)
            if not self.date_to_var.get():
                self.date_to_var.set(today)
        tk.Button(date_frame, text="اليوم", command=set_today, bg=self.settings.accent_color,
                  fg="white", relief="flat", cursor="hand2").pack(side="right", padx=5)

        options_frame = tk.LabelFrame(main_frame, text="خيارات إضافية", bg=self.settings.surface_color,
                                      font=(self.settings.font_family, 11, "bold"))
        options_frame.pack(fill="x", pady=10)
        tk.Checkbutton(options_frame, text="إضافة الوقت إلى اسم الملف", variable=self.include_time_var,
                       bg=self.settings.surface_color, font=(self.settings.font_family, 10)).pack(anchor="w", padx=10, pady=3)
        tk.Checkbutton(options_frame, text="طباعة مباشرة (بدون حفظ ملف)", variable=self.print_direct_var,
                       bg=self.settings.surface_color, font=(self.settings.font_family, 10)).pack(anchor="w", padx=10, pady=3)
        tk.Checkbutton(options_frame, text="فتح الملف تلقائياً بعد اكتمال التصدير", variable=self.open_after_export_var,
                       bg=self.settings.surface_color, font=(self.settings.font_family, 10)).pack(anchor="w", padx=10, pady=3)
        orientation_labels = {"auto": "تلقائي", "portrait": "طولي", "landscape": "عرضي"}
        quality_text = (
            f"إعداد الطباعة الحالي: {self.settings.report_page_size} - "
            f"{orientation_labels.get(self.settings.report_orientation, self.settings.report_orientation)} | "
            f"PNG: {self.settings.png_dpi} DPI | أقل خط PDF: {self.settings.pdf_min_font_size:g} pt"
        )
        tk.Label(options_frame, text=quality_text, bg="#F4F7F9", fg=self.settings.muted_color,
                 font=(self.settings.font_family, 9, "bold"), anchor="e", justify="right", padx=8, pady=7).pack(
            fill="x", padx=10, pady=(6, 3)
        )

        format_frame = tk.Frame(options_frame, bg=self.settings.surface_color)
        format_frame.pack(anchor="w", padx=10, pady=5)
        tk.Label(format_frame, text="صيغة التصدير:", bg=self.settings.surface_color,
                 font=(self.settings.font_family, 10, "bold")).pack(side="right", padx=5)
        formats = ["pdf", "xlsx", "png", "jpg", "csv", "html"]
        for fmt in formats:
            rb = tk.Radiobutton(format_frame, text=fmt.upper(), variable=self.export_format_var,
                                value=fmt, bg=self.settings.surface_color,
                                font=(self.settings.font_family, 10))
            rb.pack(side="right", padx=5)

        name_frame = tk.Frame(main_frame, bg=self.settings.background_color)
        name_frame.pack(fill="x", pady=10)
        tk.Label(name_frame, text="اسم الملف:", bg=self.settings.background_color,
                 font=(self.settings.font_family, 11, "bold")).pack(side="right", padx=5)
        tk.Entry(name_frame, textvariable=self.file_name_prefix_var, width=40, justify="right",
                 font=(self.settings.font_family, 11)).pack(side="right", padx=5)

        btn_frame = tk.Frame(main_frame, bg=self.settings.background_color)
        btn_frame.pack(fill="x", pady=20)
        self.parent_app.make_button(btn_frame, "معاينة", self._preview, self.settings.primary_color, width=15).pack(side="right", padx=5)
        self.parent_app.make_button(btn_frame, "تصدير", self._export, self.settings.success_color, width=15).pack(side="right", padx=5)
        self.parent_app.make_button(btn_frame, "إلغاء", self.destroy, self.settings.danger_color, width=15).pack(side="right", padx=5)
        tk.Button(btn_frame, text="حفظ التفضيلات", command=self._save_preferences,
                  bg=self.settings.muted_color, fg="white", relief="flat",
                  cursor="hand2", font=(self.settings.font_family, 10, "bold")).pack(side="left", padx=5)

    def _get_selected_fields(self):
        if self.selected_fields.get() == "الكل":
            return self.settings.report_fields
        return [
            f for f in self.settings.visible_fields
            if f.key != "scheduled_delivery_date"
            and f.key in self.field_checkboxes
            and self.field_checkboxes[f.key].get()
        ]

    def _get_records(self):
        from_date = self.date_from_var.get().strip()
        to_date = self.date_to_var.get().strip()
        base_records = list(self.parent_app.filtered_records)
        if from_date or to_date:
            try:
                start = dt.datetime.strptime(from_date, "%Y-%m-%d").date() if from_date else None
                end = dt.datetime.strptime(to_date, "%Y-%m-%d").date() if to_date else None
                if start and end and start > end:
                    raise ValueError
                return [
                    record for record in base_records
                    if (record_date := record_date_for_basis(record, "التاريخ الأنسب")) is not None
                    and (start is None or record_date >= start)
                    and (end is None or record_date <= end)
                ]
            except ValueError:
                messagebox.showwarning("تاريخ غير صحيح", "استخدم صيغة YYYY-MM-DD وتأكد أن تاريخ البداية لا يتجاوز النهاية.", parent=self)
                return None
        return base_records

    def _preview(self):
        records = self._get_records()
        if records is None:
            return
        if not records:
            messagebox.showinfo("لا توجد سجلات", "لا توجد سجلات في النطاق المحدد.", parent=self)
            return
        fields = self._get_selected_fields()
        if not fields:
            messagebox.showwarning("تنبيه", "اختر حقلًا واحدًا على الأقل للمعاينة.", parent=self)
            return
        preview_path = PREVIEW_DIR / "معاينة_تقرير_تسليم_الرخام.pdf"
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
            progress.close()

    def _export(self):
        records = self._get_records()
        if records is None:
            return
        if not records:
            messagebox.showinfo("لا توجد سجلات", "لا توجد سجلات للتصدير.", parent=self)
            return

        fields = self._get_selected_fields()
        if not fields:
            messagebox.showwarning("تنبيه", "لم تختر أي حقل للتصدير.", parent=self)
            return

        fmt = self.export_format_var.get()
        now = dt.datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("_%H%M") if self.include_time_var.get() else ""
        file_name = f"{self.file_name_prefix_var.get()} - {date_str}{time_str}"

        if self.print_direct_var.get():
            if fmt != "pdf":
                messagebox.showwarning("تنبيه", "الطباعة المباشرة مدعومة فقط بصيغة PDF.", parent=self)
                return
            temp_dir = tempfile.gettempdir()
            temp_file = Path(temp_dir) / f"{file_name}.pdf"
            try:
                exporter = ReportExporter(self.settings)
                exporter.export_pdf(records, self.parent_app.recipient_var.get(), temp_file, fields)
                DatabaseManager.archive_report(
                    temp_file, "pdf", self.settings, self.parent_app.recipient_var.get(), records, fields
                )
                print_pdf_file(temp_file)
                messagebox.showinfo("طباعة", "تم إرسال التقرير إلى الطابعة.", parent=self)
                self.destroy()
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل الطباعة: {e}", parent=self)
            return

        ext_map = {"pdf": ".pdf", "xlsx": ".xlsx", "png": ".png", "jpg": ".jpg", "csv": ".csv", "html": ".html"}
        type_labels = {"pdf": "PDF", "xlsx": "Excel", "png": "PNG", "jpg": "JPG", "csv": "CSV", "html": "HTML"}
        filetypes = [(type_labels[fmt], f"*{ext_map[fmt]}")]
        file_path = filedialog.asksaveasfilename(
            parent=self, title="حفظ التقرير",
            initialfile=f"{file_name}{ext_map[fmt]}",
            defaultextension=ext_map[fmt],
            filetypes=filetypes
        )
        if not file_path:
            return

        try:
            self.parent_app.status_var.set("جارٍ إنشاء التقرير...")
            self.parent_app.root.update_idletasks()
            result = self.export_callback(
                fmt, records, self.parent_app.recipient_var.get(), fields, Path(file_path),
                self.open_after_export_var.get(),
            )
            if result is False:
                self.destroy()
                return
            self.parent_app.status_var.set(f"تم التصدير: {file_path}")
            messagebox.showinfo("اكتمل", f"تم حفظ التقرير في:\n{file_path}", parent=self)
            self.destroy()
        except Exception as e:
            logging.exception("فشل التصدير")
            messagebox.showerror("خطأ", str(e), parent=self)
            self.parent_app.status_var.set("فشل التصدير")

    def _save_preferences(self):
        prefs = {
            "selected_fields": self.selected_fields.get(),
            "include_time": self.include_time_var.get(),
            "print_direct": self.print_direct_var.get(),
            "open_after_export": self.open_after_export_var.get(),
            "file_name_prefix": self.file_name_prefix_var.get(),
            "export_format": self.export_format_var.get(),
            "selected_field_keys": [key for key, var in self.field_checkboxes.items() if var.get()],
        }
        self.settings.export_preferences = prefs
        try:
            DatabaseManager.save_settings(self.settings)
            messagebox.showinfo("تم", "تم حفظ تفضيلات التصدير.", parent=self)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}", parent=self)

# ============================================================================
# نافذة الإعدادات المتقدمة
# ============================================================================
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: MarbleDeliveryApp, settings: AppSettings, on_save: Callable[[AppSettings], None]):
        super().__init__(parent.root)
        self.parent_app = parent
        self.settings = copy.deepcopy(settings)
        self.on_save = on_save
        self.title("الإعدادات المتقدمة")
        width = min(960, max(760, parent.root.winfo_screenwidth() - 80))
        height = min(780, max(620, parent.root.winfo_screenheight() - 100))
        x = parent.root.winfo_x() + (parent.root.winfo_width() - width) // 2
        y = parent.root.winfo_y() + (parent.root.winfo_height() - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(min(850, max(700, parent.root.winfo_screenwidth() - 140)),
                     min(650, max(560, parent.root.winfo_screenheight() - 180)))
        self.transient(parent.root)
        self.grab_set()
        self.configure(bg=settings.background_color)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._variables = {}
        self._build()
        self.wait_visibility()
        self.focus_force()

    def _build(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=18, pady=(18, 8))
        identity = tk.Frame(notebook, bg=self.settings.surface_color)
        appearance_tab = tk.Frame(notebook, bg=self.settings.surface_color)
        reports_tab = tk.Frame(notebook, bg=self.settings.surface_color)
        fields_tab = tk.Frame(notebook, bg=self.settings.surface_color)
        operations = tk.Frame(notebook, bg=self.settings.surface_color)
        notebook.add(identity, text="هوية الشركة")
        notebook.add(appearance_tab, text="المظهر والألوان")
        notebook.add(reports_tab, text="التقارير والطباعة")
        notebook.add(fields_tab, text="الحقول")
        notebook.add(operations, text="التشغيل والنسخ")
        self._build_identity(identity)
        self._build_reports(reports_tab)
        appearance_canvas = tk.Canvas(appearance_tab, bg=self.settings.surface_color, highlightthickness=0)
        appearance_scroll = ttk.Scrollbar(appearance_tab, orient="vertical", command=appearance_canvas.yview)
        appearance = tk.Frame(appearance_canvas, bg=self.settings.surface_color)
        appearance_window = appearance_canvas.create_window((0, 0), window=appearance, anchor="nw")
        appearance.bind("<Configure>", lambda _event: appearance_canvas.configure(scrollregion=appearance_canvas.bbox("all")))
        appearance_canvas.bind("<Configure>", lambda event: appearance_canvas.itemconfigure(appearance_window, width=event.width))
        appearance_canvas.configure(yscrollcommand=appearance_scroll.set)
        appearance_canvas.pack(side="left", fill="both", expand=True)
        appearance_scroll.pack(side="right", fill="y")
        def scroll_appearance(event):
            appearance_canvas.yview_scroll(-int(event.delta / 120), "units")
        appearance_canvas.bind("<Enter>", lambda _event: appearance_canvas.bind_all("<MouseWheel>", scroll_appearance))
        appearance_canvas.bind("<Leave>", lambda _event: appearance_canvas.unbind_all("<MouseWheel>"))
        self._build_appearance(appearance)
        self._build_fields(fields_tab)
        self._build_operations(operations)
        footer = tk.Frame(self, bg=self.settings.background_color)
        footer.pack(fill="x", padx=18, pady=(0, 16))
        self.parent_app.make_button(footer, "حفظ", self._save, self.settings.success_color).pack(side="right", padx=5)
        self.parent_app.make_button(footer, "إلغاء", self.destroy, self.settings.muted_color).pack(side="right", padx=5)
        self.parent_app.make_button(footer, "استعادة الافتراضيات", self._restore_defaults, self.settings.danger_color).pack(side="left", padx=5)

    def _entry_row(self, parent, row, label, key, value):
        var = tk.StringVar(value=value)
        self._variables[key] = var
        tk.Label(parent, text=label, bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, self.settings.base_font_size, "bold"), anchor="e").grid(
            row=row, column=1, sticky="e", padx=12, pady=8
        )
        entry = tk.Entry(parent, textvariable=var, justify="right", relief="solid", bd=1,
                         font=(self.settings.font_family, self.settings.base_font_size))
        entry.grid(row=row, column=0, sticky="ew", padx=12, pady=8, ipady=6)

    def _build_identity(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        self._entry_row(parent, 0, "اسم الشركة", "company_name", self.settings.company_name)
        self._entry_row(parent, 1, "عنوان التقرير", "report_title", self.settings.report_title)
        self._entry_row(parent, 2, "عنوان التقرير بالإنجليزية", "report_title_en", self.settings.report_title_en)
        self._entry_row(parent, 3, "العنوان الفرعي", "report_subtitle", self.settings.report_subtitle)
        self._entry_row(parent, 4, "عبارة التذييل", "footer_note", self.settings.footer_note)
        self._entry_row(parent, 5, "مسار الشعار", "logo_path", self.settings.logo_path)
        self.parent_app.make_button(parent, "اختيار شعار", self._choose_logo, self.settings.primary_color, width=16).grid(
            row=6, column=0, sticky="e", padx=12, pady=6
        )

    def _choose_logo(self):
        path = filedialog.askopenfilename(parent=self, title="اختيار شعار",
                                          filetypes=[("صور", "*.png *.jpg *.jpeg *.bmp")])
        if path:
            self._variables["logo_path"].set(path)

    def _build_reports(self, parent):
        """إعدادات مستقلة لمقاسات الورق ووضوح PDF/PNG/JPG والطباعة."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=0)
        s = self.settings
        self._variables["report_page_size"] = tk.StringVar(value=s.report_page_size)
        orientation_labels = {"auto": "تلقائي", "portrait": "طولي", "landscape": "عرضي"}
        self._variables["report_orientation"] = tk.StringVar(value=orientation_labels.get(s.report_orientation, "تلقائي"))
        for key, value, var_type in [
            ("report_margin_mm", s.report_margin_mm, tk.DoubleVar), ("pdf_base_font_size", s.pdf_base_font_size, tk.DoubleVar),
            ("pdf_min_font_size", s.pdf_min_font_size, tk.DoubleVar), ("png_dpi", s.png_dpi, tk.IntVar),
            ("jpg_quality", s.jpg_quality, tk.IntVar), ("custom_page_width_mm", s.custom_page_width_mm, tk.DoubleVar),
            ("custom_page_height_mm", s.custom_page_height_mm, tk.DoubleVar),
        ]:
            self._variables[key] = var_type(value=value)
        rows = [
            ("مقاس الورق", "report_page_size", ("A4", "A3", "LETTER", "CUSTOM")),
            ("اتجاه الصفحة", "report_orientation", ("تلقائي", "طولي", "عرضي")),
        ]
        for row, (label, key, values) in enumerate(rows):
            tk.Label(parent, text=label, bg=s.surface_color, fg=s.text_color, font=(s.font_family, s.base_font_size, "bold"), anchor="e").grid(row=row, column=1, sticky="e", padx=14, pady=7)
            ttk.Combobox(parent, textvariable=self._variables[key], values=values, state="readonly", justify="center", font=(s.font_family, s.base_font_size)).grid(row=row, column=0, sticky="ew", padx=14, pady=7)
        numeric_rows = [
            ("عرض الصفحة المخصصة", "custom_page_width_mm", 80, 1000, 1, "مم"),
            ("ارتفاع الصفحة المخصصة", "custom_page_height_mm", 80, 1000, 1, "مم"),
            ("هامش الصفحة", "report_margin_mm", 3, 20, 0.5, "مم"),
            ("حجم خط PDF الأساسي", "pdf_base_font_size", 8, 14, 0.25, "pt"),
            ("أقل خط مسموح في PDF", "pdf_min_font_size", 6.5, 12, 0.25, "pt"),
            ("دقة PNG/JPG", "png_dpi", 150, 600, 50, "DPI"),
            ("جودة JPG", "jpg_quality", 70, 100, 1, "%"),
        ]
        for offset, (label, key, low, high, step, suffix) in enumerate(numeric_rows, start=2):
            tk.Label(parent, text=label, bg=s.surface_color, fg=s.text_color, font=(s.font_family, s.base_font_size, "bold"), anchor="e").grid(row=offset, column=1, sticky="e", padx=14, pady=6)
            holder = tk.Frame(parent, bg=s.surface_color); holder.grid(row=offset, column=0, sticky="ew", padx=14, pady=6)
            tk.Spinbox(holder, from_=low, to=high, increment=step, textvariable=self._variables[key], justify="center", font=(s.font_family, s.base_font_size)).pack(side="right", fill="x", expand=True, ipady=4)
            tk.Label(holder, text=suffix, bg=s.surface_color, fg=s.muted_color, width=6).pack(side="right", padx=8)
        note = (
            "PDF متجهي مع خطوط عربية مضمّنة. عند كثرة البيانات يضيف المحرك صفحات بدلاً من ضغط النص تحت الحد الأدنى. "
            "PNG/JPG يستخدمان المقاس الحقيقي للورق ودقة DPI المختارة؛ 300 DPI مناسب للطباعة و600 DPI للأرشفة عالية الجودة. "
            "CUSTOM يسمح بإدخال أي مقاس بالملليمتر."
        )
        tk.Label(parent, text=note, wraplength=720, justify="right", anchor="e", bg="#F4F7F9", fg=s.muted_color, font=(s.font_family, 10), padx=14, pady=12).grid(row=9, column=0, columnspan=2, sticky="ew", padx=14, pady=12)

    def _build_operations(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        self._entry_row(parent, 0, "بادئة الترقيم التلقائي", "auto_contract_prefix", self.settings.auto_contract_prefix)
        self._entry_row(parent, 1, "مجلد مزامنة النسخ الاحتياطية", "sync_folder", self.settings.sync_folder)
        self.parent_app.make_button(
            parent, "اختيار مجلد", self._choose_sync_folder, self.settings.primary_color, width=16
        ).grid(row=2, column=0, sticky="e", padx=12, pady=4)

        self._variables["logo_width_mm"] = tk.IntVar(value=self.settings.logo_width_mm)
        for row, (label, key, low, high, suffix) in enumerate([
            ("عرض الشعار في التقارير", "logo_width_mm", 15, 60, "مم"),
        ], 3):
            tk.Label(parent, text=label, bg=self.settings.surface_color, fg=self.settings.text_color,
                     font=(self.settings.font_family, self.settings.base_font_size, "bold"), anchor="e").grid(
                row=row, column=1, sticky="e", padx=12, pady=10
            )
            holder = tk.Frame(parent, bg=self.settings.surface_color)
            holder.grid(row=row, column=0, sticky="ew", padx=12, pady=10)
            tk.Spinbox(holder, from_=low, to=high, textvariable=self._variables[key], justify="center",
                       font=(self.settings.font_family, self.settings.base_font_size)).pack(side="right", fill="x", expand=True, ipady=4)
            tk.Label(holder, text=suffix, bg=self.settings.surface_color, fg=self.settings.muted_color).pack(side="right", padx=8)

        self._variables["overdue_warning"] = tk.BooleanVar(value=self.settings.overdue_warning)
        self._variables["backup_keep_count"] = tk.IntVar(value=self.settings.backup_keep_count)
        self._variables["autosave_seconds"] = tk.IntVar(value=self.settings.autosave_seconds)
        self._variables["require_login"] = tk.BooleanVar(value=self.settings.require_login)
        self._variables["compact_form_default"] = tk.BooleanVar(value=self.settings.compact_form_default)
        self._variables["default_export_folder"] = tk.StringVar(value=self.settings.default_export_folder)
        tk.Label(parent, text="مجلد التصدير الافتراضي", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, self.settings.base_font_size, "bold"), anchor="e").grid(row=5, column=1, sticky="e", padx=12, pady=7)
        tk.Entry(parent, textvariable=self._variables["default_export_folder"], justify="right").grid(row=5, column=0, sticky="ew", padx=12, pady=7, ipady=4)
        perf = tk.Frame(parent, bg=self.settings.surface_color); perf.grid(row=6, column=0, columnspan=2, sticky="ew", padx=12, pady=6)
        tk.Label(perf, text="عدد النسخ التلقائية", bg=self.settings.surface_color).pack(side="right", padx=6)
        tk.Spinbox(perf, from_=5, to=100, textvariable=self._variables["backup_keep_count"], width=7, justify="center").pack(side="right")
        tk.Label(perf, text="حفظ المسودة كل", bg=self.settings.surface_color).pack(side="right", padx=(18, 6))
        tk.Spinbox(perf, from_=2, to=60, textvariable=self._variables["autosave_seconds"], width=7, justify="center").pack(side="right")
        tk.Label(perf, text="ثانية", bg=self.settings.surface_color).pack(side="right", padx=4)
        tk.Checkbutton(
            parent, text="تنبيه عند وجود تسليمات متأخرة عند بدء البرنامج",
            variable=self._variables["overdue_warning"], bg=self.settings.surface_color,
            fg=self.settings.text_color, activebackground=self.settings.surface_color,
            font=(self.settings.font_family, 11, "bold"), anchor="e",
        ).grid(row=7, column=0, columnspan=2, sticky="e", padx=12, pady=6)
        tk.Checkbutton(parent, text="طلب تسجيل الدخول عند تشغيل البرنامج", variable=self._variables["require_login"], bg=self.settings.surface_color, fg=self.settings.text_color, activebackground=self.settings.surface_color, font=(self.settings.font_family, 10, "bold"), anchor="e").grid(row=8, column=0, columnspan=2, sticky="e", padx=12, pady=5)
        tk.Checkbutton(parent, text="بدء البرنامج ولوحة الإدخال مطوية لتكبير التقرير", variable=self._variables["compact_form_default"], bg=self.settings.surface_color, fg=self.settings.text_color, activebackground=self.settings.surface_color, font=(self.settings.font_family, 10, "bold"), anchor="e").grid(row=9, column=0, columnspan=2, sticky="e", padx=12, pady=5)

        note = (
            "يمكن تحديد مجلد شبكي أو مجلد سحابي محلي (مثل OneDrive) لحفظ نسخة تلقائية إضافية. "
            "تبقى قاعدة البيانات الأساسية محلية لتجنب تعارض التعديلات."
        )
        tk.Label(parent, text=note, wraplength=650, justify="right", anchor="e",
                 bg="#F4F7F9", fg=self.settings.muted_color,
                 font=(self.settings.font_family, 10), padx=12, pady=12).grid(
            row=10, column=0, columnspan=2, sticky="ew", padx=12, pady=10
        )

    def _choose_sync_folder(self):
        path = filedialog.askdirectory(parent=self, title="اختيار مجلد مزامنة النسخ الاحتياطية")
        if path:
            self._variables["sync_folder"].set(path)

    def _build_appearance(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        families = sorted(set(tkfont.families(self)))
        self._variables["font_family"] = tk.StringVar(value=self.settings.font_family)
        tk.Label(parent, text="الخط", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, self.settings.base_font_size, "bold")).grid(row=0, column=1, sticky="e", padx=12, pady=8)
        ttk.Combobox(parent, textvariable=self._variables["font_family"], values=families, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=12, pady=8
        )
        for row, (label, key, value) in enumerate([
            ("حجم النص", "base_font_size", self.settings.base_font_size),
            ("حجم العنوان", "heading_font_size", self.settings.heading_font_size),
        ], 1):
            self._variables[key] = tk.IntVar(value=value)
            tk.Label(parent, text=label, bg=self.settings.surface_color, fg=self.settings.text_color,
                     font=(self.settings.font_family, self.settings.base_font_size, "bold")).grid(row=row, column=1, sticky="e", padx=12, pady=8)
            tk.Spinbox(parent, from_=10 if key=="base_font_size" else 16,
                       to=22 if key=="base_font_size" else 34, textvariable=self._variables[key],
                       justify="center", font=(self.settings.font_family, self.settings.base_font_size)).grid(row=row, column=0, sticky="ew", padx=12, pady=8)

        color_items = [
            ("لون خلفية العنوان", "title_bg_color"),
            ("لون نص العنوان", "title_fg_color"),
            ("لون خلفية العنوان الفرعي", "subtitle_bg_color"),
            ("لون نص العنوان الفرعي", "subtitle_fg_color"),
        ]
        start_row = 3
        for idx, (label, key) in enumerate(color_items):
            row = start_row + idx
            value = getattr(self.settings, key)
            var = tk.StringVar(value=value)
            self._variables[key] = var
            tk.Label(parent, text=label, bg=self.settings.surface_color, fg=self.settings.text_color,
                     font=(self.settings.font_family, 10, "bold")).grid(row=row, column=1, sticky="e", padx=12, pady=5)
            container = tk.Frame(parent, bg=self.settings.surface_color)
            container.grid(row=row, column=0, sticky="ew", padx=12, pady=5)
            entry = tk.Entry(container, textvariable=var, justify="center", width=16)
            entry.pack(side="right", fill="x", expand=True, ipady=4)
            tk.Button(container, text="اختيار", command=lambda k=key: self._choose_color(k),
                      bg=value, fg="white", relief="flat", cursor="hand2").pack(side="left", padx=(0, 6), ipadx=10, ipady=3)

        colors_map = [
            ("اللون الرئيسي", "primary_color"),
            ("اللون المميز", "accent_color"),
            ("لون الخلفية", "background_color"),
            ("لون البطاقات", "surface_color"),
            ("لون النص", "text_color"),
            ("لون النص الثانوي", "muted_color"),
            ("لون النجاح", "success_color"),
            ("لون الحذف", "danger_color"),
        ]
        for idx, (label, key) in enumerate(colors_map):
            row = start_row + len(color_items) + idx
            value = getattr(self.settings, key)
            var = tk.StringVar(value=value)
            self._variables[key] = var
            tk.Label(parent, text=label, bg=self.settings.surface_color, fg=self.settings.text_color,
                     font=(self.settings.font_family, 10, "bold")).grid(row=row, column=1, sticky="e", padx=12, pady=5)
            container = tk.Frame(parent, bg=self.settings.surface_color)
            container.grid(row=row, column=0, sticky="ew", padx=12, pady=5)
            entry = tk.Entry(container, textvariable=var, justify="center", width=16)
            entry.pack(side="right", fill="x", expand=True, ipady=4)
            tk.Button(container, text="اختيار", command=lambda k=key: self._choose_color(k),
                      bg=value, fg="white", relief="flat", cursor="hand2").pack(side="left", padx=(0, 6), ipadx=10, ipady=3)

        self._variables["print_friendly_reports"] = tk.BooleanVar(value=self.settings.print_friendly_reports)
        tk.Checkbutton(parent, text="نمط تقارير فاتح للطباعة", variable=self._variables["print_friendly_reports"],
                       bg=self.settings.surface_color, fg=self.settings.text_color,
                       activebackground=self.settings.surface_color, anchor="e",
                       font=(self.settings.font_family, 11, "bold")).grid(row=20, column=0, columnspan=2, sticky="e", padx=12, pady=10)

    def _choose_color(self, key):
        selected = colorchooser.askcolor(self._variables[key].get(), parent=self, title="اختيار اللون")[1]
        if selected:
            self._variables[key].set(selected.upper())

    def _build_fields(self, parent):
        toolbar = tk.Frame(parent, bg=self.settings.surface_color)
        toolbar.pack(fill="x", padx=12, pady=12)
        for text, command, color in [
            ("إضافة", self._add_field, self.settings.success_color),
            ("تعديل", self._edit_field, self.settings.primary_color),
            ("حذف", self._delete_field, self.settings.danger_color),
            ("↑", lambda: self._move_field(-1), self.settings.muted_color),
            ("↓", lambda: self._move_field(1), self.settings.muted_color),
        ]:
            self.parent_app.make_button(toolbar, text, command, color, width=10).pack(side="right", padx=4)
        columns = ("label", "visible", "required", "width")
        self.fields_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        self.fields_tree.heading("label", text="اسم الحقل")
        self.fields_tree.heading("visible", text="ظاهر")
        self.fields_tree.heading("required", text="مطلوب")
        self.fields_tree.heading("width", text="العرض")
        self.fields_tree.column("label", width=300, anchor="e")
        self.fields_tree.column("visible", width=80, anchor="center")
        self.fields_tree.column("required", width=80, anchor="center")
        self.fields_tree.column("width", width=100, anchor="center")
        self.fields_tree.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.fields_tree.bind("<Double-1>", lambda e: self._edit_field())
        self._refresh_fields()

    def _refresh_fields(self, selected=None):
        self.fields_tree.delete(*self.fields_tree.get_children())
        for idx, f in enumerate(self.settings.fields):
            self.fields_tree.insert("", "end", iid=str(idx), values=(f.label, "نعم" if f.visible else "لا",
                                                                     "نعم" if f.required else "لا", f.width))
        if selected is not None and 0 <= selected < len(self.settings.fields):
            self.fields_tree.selection_set(str(selected))
            self.fields_tree.see(str(selected))

    def _selected_field_index(self):
        sel = self.fields_tree.selection()
        return int(sel[0]) if sel else None

    def _field_editor(self, item, title):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.geometry("520x480")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=self.settings.surface_color)
        label_var = tk.StringVar(value=item.label)
        width_var = tk.IntVar(value=item.width)
        visible_var = tk.BooleanVar(value=item.visible)
        required_var = tk.BooleanVar(value=item.required)
        kind_var = tk.StringVar(value=item.kind)
        options_var = tk.StringVar(value="، ".join(item.options))
        result = {"saved": False}
        rows = [("اسم الحقل", label_var), ("العرض (80-320)", width_var)]
        for row, (lbl, var) in enumerate(rows):
            tk.Label(dialog, text=lbl, bg=self.settings.surface_color, fg=self.settings.text_color,
                     font=(self.settings.font_family, 11, "bold")).grid(row=row, column=1, sticky="e", padx=16, pady=12)
            tk.Entry(dialog, textvariable=var, justify="right").grid(row=row, column=0, sticky="ew", padx=16, pady=12, ipady=5)
        tk.Label(dialog, text="نوع التحقق", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, 11, "bold")).grid(row=2, column=1, sticky="e", padx=16, pady=12)
        ttk.Combobox(dialog, textvariable=kind_var, values=("text", "phone", "number", "date", "choice"), state="readonly").grid(row=2, column=0, sticky="ew", padx=16, pady=12)
        tk.Label(dialog, text="خيارات القائمة", bg=self.settings.surface_color, fg=self.settings.text_color,
                 font=(self.settings.font_family, 11, "bold")).grid(row=3, column=1, sticky="e", padx=16, pady=12)
        tk.Entry(dialog, textvariable=options_var, justify="right").grid(row=3, column=0, sticky="ew", padx=16, pady=12, ipady=5)
        tk.Label(dialog, text="افصل الخيارات بفاصلة عربية أو إنجليزية", bg=self.settings.surface_color,
                 fg=self.settings.muted_color).grid(row=4, column=0, sticky="e", padx=16, pady=(0, 6))
        tk.Checkbutton(dialog, text="إظهار الحقل", variable=visible_var, bg=self.settings.surface_color,
                       fg=self.settings.text_color, anchor="e").grid(row=5, column=0, columnspan=2, sticky="e", padx=16, pady=6)
        tk.Checkbutton(dialog, text="مطلوب", variable=required_var, bg=self.settings.surface_color,
                       fg=self.settings.text_color, anchor="e").grid(row=6, column=0, columnspan=2, sticky="e", padx=16, pady=6)
        dialog.grid_columnconfigure(0, weight=1)

        def save():
            label = clean_value(label_var.get())
            if not label:
                messagebox.showwarning("تنبيه", "أدخل اسم الحقل.", parent=dialog)
                return
            try:
                width = int(width_var.get())
            except:
                width = 0
            if not 80 <= width <= 320:
                messagebox.showwarning("تنبيه", "العرض بين 80 و320.", parent=dialog)
                return
            item.label = label
            item.width = width
            item.visible = visible_var.get()
            item.required = required_var.get()
            item.kind = kind_var.get()
            item.options = [clean_value(value) for value in re.split(r"[,،\n]", options_var.get()) if clean_value(value)]
            if item.kind == "choice" and not item.options:
                messagebox.showwarning("تنبيه", "أدخل خياراً واحداً على الأقل لحقل القائمة.", parent=dialog)
                return
            result["saved"] = True
            dialog.destroy()

        btns = tk.Frame(dialog, bg=self.settings.surface_color)
        btns.grid(row=7, column=0, columnspan=2, pady=18)
        self.parent_app.make_button(btns, "حفظ", save, self.settings.success_color).pack(side="right", padx=5)
        self.parent_app.make_button(btns, "إلغاء", dialog.destroy, self.settings.muted_color).pack(side="right", padx=5)
        dialog.wait_window()
        return result["saved"]

    def _add_field(self):
        f = FieldDefinition(f"custom_{uuid.uuid4().hex[:10]}", "حقل جديد", system=False)
        if self._field_editor(f, "إضافة حقل"):
            self.settings.fields.append(f)
            self._refresh_fields(len(self.settings.fields)-1)

    def _edit_field(self):
        idx = self._selected_field_index()
        if idx is None:
            messagebox.showwarning("تنبيه", "حدد حقلًا.", parent=self)
            return
        if self._field_editor(self.settings.fields[idx], "تعديل حقل"):
            self._refresh_fields(idx)

    def _delete_field(self):
        idx = self._selected_field_index()
        if idx is None:
            return
        f = self.settings.fields[idx]
        if f.system:
            messagebox.showinfo("حقل أساسي", "لا يمكن حذف الحقول الأساسية.", parent=self)
            return
        if messagebox.askyesno("تأكيد", f"حذف الحقل «{f.label}»؟", parent=self):
            del self.settings.fields[idx]
            self._refresh_fields(max(0, idx-1))

    def _move_field(self, delta):
        idx = self._selected_field_index()
        if idx is None:
            return
        new = idx + delta
        if 0 <= new < len(self.settings.fields):
            self.settings.fields[idx], self.settings.fields[new] = self.settings.fields[new], self.settings.fields[idx]
            self._refresh_fields(new)

    def _restore_defaults(self):
        if messagebox.askyesno("استعادة", "استعادة الإعدادات الافتراضية؟", parent=self):
            self.settings = AppSettings()
            self.destroy()
            SettingsDialog(self.parent_app, self.settings, self.on_save)

    def _save(self):
        for key in ["company_name", "report_title", "report_title_en", "report_subtitle", "footer_note", "logo_path", "font_family",
                    "title_bg_color", "title_fg_color", "subtitle_bg_color", "subtitle_fg_color"]:
            setattr(self.settings, key, clean_value(self._variables[key].get()))
        self.settings.base_font_size = int(self._variables["base_font_size"].get())
        self.settings.heading_font_size = int(self._variables["heading_font_size"].get())
        self.settings.print_friendly_reports = bool(self._variables["print_friendly_reports"].get())
        self.settings.auto_contract_prefix = re.sub(r"[^A-Za-z0-9_-]", "", clean_value(self._variables["auto_contract_prefix"].get()).upper()) or "MD"
        self.settings.sync_folder = clean_value(self._variables["sync_folder"].get())
        self.settings.logo_width_mm = max(15, min(60, int(self._variables["logo_width_mm"].get())))
        self.settings.report_page_size = str(self._variables["report_page_size"].get()).upper()
        orientation_values = {"تلقائي": "auto", "طولي": "portrait", "عرضي": "landscape"}
        self.settings.report_orientation = orientation_values.get(
            str(self._variables["report_orientation"].get()), "auto"
        )
        self.settings.report_margin_mm = max(3.0, min(20.0, float(self._variables["report_margin_mm"].get())))
        self.settings.pdf_base_font_size = max(8.0, min(14.0, float(self._variables["pdf_base_font_size"].get())))
        self.settings.pdf_min_font_size = max(6.5, min(self.settings.pdf_base_font_size, float(self._variables["pdf_min_font_size"].get())))
        self.settings.png_dpi = max(150, min(600, int(self._variables["png_dpi"].get())))
        self.settings.jpg_quality = max(70, min(100, int(self._variables["jpg_quality"].get())))
        self.settings.custom_page_width_mm = max(80.0, min(1000.0, float(self._variables["custom_page_width_mm"].get())))
        self.settings.custom_page_height_mm = max(80.0, min(1000.0, float(self._variables["custom_page_height_mm"].get())))
        self.settings.default_export_folder = clean_value(self._variables["default_export_folder"].get())
        self.settings.backup_keep_count = max(5, min(100, int(self._variables["backup_keep_count"].get())))
        self.settings.autosave_seconds = max(2, min(60, int(self._variables["autosave_seconds"].get())))
        self.settings.require_login = bool(self._variables["require_login"].get())
        self.settings.compact_form_default = bool(self._variables["compact_form_default"].get())
        self.settings.overdue_warning = bool(self._variables["overdue_warning"].get())
        color_keys = ["primary_color", "accent_color", "background_color", "surface_color",
                      "text_color", "muted_color", "success_color", "danger_color"]
        for key in color_keys:
            value = str(self._variables[key].get()).upper()
            if not is_hex_color(value):
                messagebox.showwarning("لون غير صحيح", f"القيمة «{value}» غير صالحة.", parent=self)
                return
            setattr(self.settings, key, value)
        for key in ["title_bg_color", "title_fg_color", "subtitle_bg_color", "subtitle_fg_color"]:
            value = getattr(self.settings, key)
            if not is_hex_color(value):
                messagebox.showwarning("لون غير صحيح", f"القيمة «{value}» غير صالحة للون العنوان.", parent=self)
                return
        if not self.settings.company_name or not self.settings.report_title or not self.settings.report_title_en:
            messagebox.showwarning("بيانات ناقصة", "اسم الشركة وعنوانا التقرير العربي والإنجليزي مطلوبة.", parent=self)
            return
        if self.settings.logo_path and not Path(self.settings.logo_path).is_file():
            messagebox.showwarning("الشعار", "مسار الشعار غير موجود.", parent=self)
            return
        if not any(f.visible for f in self.settings.fields):
            messagebox.showwarning("الحقول", "يجب إظهار حقل واحد على الأقل.", parent=self)
            return
        self.on_save(self.settings)
        self.destroy()

# ============================================================================
# التطبيق الرئيسي
# ============================================================================
class MarbleDeliveryApp:
    def __init__(self, root: tk.Tk, current_user: Optional[dict[str, Any]] = None):
        self.root = root
        self.current_user = current_user or dict(CURRENT_USER)
        self.set_current_user(self.current_user, refresh_ui=False)
        self.settings = DatabaseManager.load_settings()
        self.records = DatabaseManager.load_records()
        self.filtered_records = self.records[:]
        self.search_var = tk.StringVar()
        self.recipient_var = tk.StringVar()
        self.status_var = tk.StringVar(value="جاهز")
        self.clock_var = tk.StringVar()
        self.stats_vars = {
            "total": tk.StringVar(), "today": tk.StringVar(), "filtered": tk.StringVar(),
            "ready": tk.StringVar(), "delivered": tk.StringVar(), "overdue": tk.StringVar(),
        }
        self.logo_image = None
        self.entries: dict[str, tk.Entry] = {}
        self.selected_record_id: str | None = None
        # None = عرض الأرشيف الكامل، set = السجلات التابعة لمساحة التقرير الحالية فقط.
        # إنشاء تقرير جديد لا يحذف أي شيء من قاعدة البيانات؛ بل يفتح مساحة عمل مستقلة وفارغة.
        self.current_report_record_ids: set[str] | None = None
        self.report_scope_var = tk.StringVar(value="عرض جميع السجلات")
        self.selection_count_var = tk.StringVar(value="المحدد: 0")
        self.table_sort_key: str | None = None
        self.table_sort_desc = False
        self._entry_order: list[str] = []
        self._form_undo_stack: list[dict[str, Any]] = []
        self._form_undo_busy = False
        self._ignore_tree_select_once = False
        self.filters: dict[str, str] = {}
        self._search_trace_id: str | None = None
        self._filter_after_id = None
        self._column_save_after_id = None
        self._session_after_id = None
        self._record_autosave_after_id = None
        self._record_autosave_busy = False
        self._last_record_snapshot = ""
        self.sidebar_visible = False
        self.sidebar_width = 240
        self.sidebar_manual_override = False
        self._resize_after_id = None
        self.form_visible = not self.settings.compact_form_default
        self.filters_visible = True
        self.stats_visible = True
        self.report_focus_mode = False
        self.window_fullscreen = False

        self.root.title(f"{APP_NAME} - الإصدار {APP_VERSION}")
        screen_w = max(900, self.root.winfo_screenwidth())
        screen_h = max(650, self.root.winfo_screenheight())
        preferred_w = min(1500, max(980, int(screen_w * 0.94)))
        preferred_h = min(900, max(680, int(screen_h * 0.90)))
        min_w = min(1100, max(900, int(screen_w * 0.78)))
        min_h = min(700, max(620, int(screen_h * 0.72)))
        x = max(0, (screen_w - preferred_w) // 2)
        y = max(0, (screen_h - preferred_h) // 2)
        self.root.geometry(f"{preferred_w}x{preferred_h}+{x}+{y}")
        self.root.minsize(min_w, min_h)
        self.sidebar_width = 220 if screen_w < 1450 else 240
        self.sidebar_visible = False
        try:
            self.root.state("zoomed")
        except:
            pass
        icon_file = resource_path("app_icon.ico")
        if icon_file.is_file():
            try:
                self.root.iconbitmap(default=str(icon_file))
            except Exception:
                logging.exception("تعذر تطبيق أيقونة النافذة")
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self._configure_styles()
        self._build_ui()
        self._bind_shortcuts()
        self.root.bind("<Configure>", self._handle_root_resize, add="+")
        previous_session = DatabaseManager.load_session_state()
        self._restore_report_workspace(previous_session.get("payload", {}))
        restored_session = False
        if not previous_session.get("clean_shutdown", True):
            restored_session = self._restore_session_state(previous_session.get("payload", {}))
        self._apply_filters()
        DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
        self._session_after_id = self.root.after(int(self.settings.autosave_seconds * 1000), self._autosave_session_state)
        if restored_session:
            self.status_var.set("تمت استعادة مسودة الجلسة السابقة تلقائيًا")
        self._update_clock()
        self.root.after(300, self._dependency_notice)
        self.root.after(500, self._database_health_notice)
        self.root.after(900, self._overdue_notice)

    def _configure_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except:
            pass
        s = self.settings
        style.configure(
            "Modern.Treeview", background=s.surface_color, fieldbackground=s.surface_color,
            foreground=s.text_color, rowheight=max(36, s.base_font_size * 2.6),
            font=(s.font_family, s.base_font_size), borderwidth=0,
        )
        style.configure(
            "Modern.Treeview.Heading", background=s.primary_color, foreground="white",
            font=(s.font_family, s.base_font_size+1, "bold"), relief="flat", padding=7,
        )
        style.map("Modern.Treeview", background=[("selected", s.accent_color)], foreground=[("selected", "white")])
        style.configure("TNotebook", background=s.background_color, borderwidth=0)
        style.configure("TNotebook.Tab", font=(s.font_family, s.base_font_size, "bold"), padding=(16, 8))
        style.configure("TCombobox", font=(s.font_family, s.base_font_size))
        style.configure("TLabelframe", background=s.surface_color, foreground=s.text_color)
        style.configure("TLabelframe.Label", font=(s.font_family, s.base_font_size, "bold"))

    def make_button(self, parent, text, command, color, width=20, fg="white", pady=10):
        return tk.Button(
            parent, text=text, command=command, bg=color, fg=fg,
            activebackground=color, activeforeground=fg, relief="flat", bd=0,
            cursor="hand2", width=width, pady=pady,
            font=(self.settings.font_family, self.settings.base_font_size, "bold"),
        )

    def _build_ui(self):
        for child in self.root.winfo_children():
            child.destroy()
        s = self.settings
        self.root.configure(bg=s.background_color)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, minsize=self.sidebar_width if self.sidebar_visible else 0)

        main = tk.Frame(self.root, bg=s.background_color)
        sidebar = tk.Frame(self.root, bg=s.primary_color, width=self.sidebar_width)
        main.grid(row=0, column=0, sticky="nsew")
        sidebar.grid(row=0, column=1, sticky="nsew")
        sidebar.grid_propagate(False)
        self.main_frame = main
        self.sidebar_frame = sidebar
        if not self.sidebar_visible:
            sidebar.grid_remove()
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        self._build_sidebar(sidebar)
        self._build_header(main)
        self._build_stats(main)
        self._build_filters(main)
        self._build_form(main)
        self._build_table(main)
        self._apply_panel_visibility()

        status = tk.Label(
            main, textvariable=self.status_var, bg="#E9ECEF", fg=s.muted_color,
            anchor="e", padx=7, pady=2, font=(s.font_family, 9),
        )
        status.grid(row=5, column=0, sticky="ew", padx=2, pady=(0, 2))

    def _build_sidebar(self, parent):
        """لوحة معلومات اختيارية على اليمين. جميع الأوامر الأساسية أصبحت في القوائم العلوية."""
        s = self.settings
        canvas = tk.Canvas(parent, bg=s.primary_color, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas, bg=s.primary_color)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        logo_holder = tk.Frame(content, bg=s.primary_color, height=132)
        logo_holder.pack(fill="x", pady=(16, 4))
        logo_holder.pack_propagate(False)
        self.logo_image = None
        if s.logo_path and Path(s.logo_path).is_file() and Image is not None and ImageTk is not None:
            try:
                logo = Image.open(s.logo_path).convert("RGBA")
                resampling = getattr(Image, "Resampling", Image)
                logo.thumbnail((150, 82), resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(logo)
            except Exception:
                pass
        if self.logo_image:
            tk.Label(logo_holder, image=self.logo_image, bg=s.primary_color).pack(pady=(0, 3))
        else:
            tk.Label(logo_holder, text="◆", bg=s.primary_color, fg=s.accent_color,
                     font=(s.font_family, 36, "bold")).pack(pady=(0, 1))
        tk.Label(logo_holder, text=s.company_name, bg=s.primary_color, fg="white",
                 wraplength=190, font=(s.font_family, 13, "bold")).pack()

        tk.Label(content, text="لوحة معلومات التقرير", bg=s.primary_color, fg=s.accent_color,
                 anchor="e", padx=16, font=(s.font_family, 11, "bold")).pack(fill="x", pady=(10, 6))
        tk.Label(
            content,
            text="تم نقل أوامر التصدير والطباعة والسجلات والأدوات إلى قوائم منسدلة أعلى الشاشة لتوسيع مساحة التقرير.",
            bg=s.primary_color, fg="#DDE6EE", justify="right", anchor="e", wraplength=188,
            padx=16, font=(s.font_family, 9),
        ).pack(fill="x", pady=(0, 10))

        info = tk.Frame(content, bg="#24445F", highlightthickness=0)
        info.pack(fill="x", padx=12, pady=4)
        for label, variable in [
            ("الإجمالي", self.stats_vars["total"]),
            ("الظاهرة", self.stats_vars["filtered"]),
            ("جاهزة", self.stats_vars["ready"]),
            ("تم التسليم", self.stats_vars["delivered"]),
        ]:
            row = tk.Frame(info, bg="#24445F")
            row.pack(fill="x", padx=10, pady=3)
            tk.Label(row, textvariable=variable, bg="#24445F", fg="white",
                     font=(s.font_family, 11, "bold")).pack(side="left")
            tk.Label(row, text=label, bg="#24445F", fg="#DDE6EE",
                     font=(s.font_family, 9, "bold")).pack(side="right")

        tk.Label(content, text="اسم المستلم", bg=s.primary_color, fg="#DDE6EE",
                 anchor="e", padx=16, font=(s.font_family, 9, "bold")).pack(fill="x", pady=(12, 4))
        tk.Entry(content, textvariable=self.recipient_var, justify="right", relief="flat",
                 font=(s.font_family, s.base_font_size)).pack(fill="x", padx=14, ipady=5)

        tk.Frame(content, bg=s.accent_color, height=1).pack(fill="x", padx=14, pady=14)
        tk.Label(content, text="Ctrl+M  إظهار/إخفاء اللوحة", bg=s.primary_color, fg="#B8C5D0",
                 font=(s.font_family, 8)).pack(anchor="e", padx=16, pady=2)
        tk.Label(content, text=f"الإصدار {APP_VERSION}", bg=s.primary_color, fg="#B8C5D0",
                 font=(s.font_family, 9)).pack(pady=12)

    def open_export_options(self):
        if not self._require_permission("export", "خيارات التصدير"):
            return
        ExportOptionsDialog(self, self._export_with_options)

    def _export_with_options(self, file_type, records, recipient, fields, output_path, open_after=False):
        if file_type not in EXPORT_METHOD_NAMES:
            raise ValueError("صيغة تصدير غير مدعومة.")
        after = open_with_default_app if open_after else None
        self._start_export_task(file_type, records, recipient, fields, Path(output_path), after_success=after)
        return False

    def _make_top_menu(self, parent, text, items):
        s = self.settings
        button = tk.Menubutton(
            parent, text=f"{text} ▾", bg=s.primary_color, fg="white",
            activebackground=s.accent_color, activeforeground="#17212B",
            relief="flat", bd=0, cursor="hand2", padx=11, pady=3,
            font=(s.font_family, 10, "bold"), direction="below",
        )
        menu = tk.Menu(
            button, tearoff=False, bg=s.surface_color, fg=s.text_color,
            activebackground=s.accent_color, activeforeground="#17212B",
            bd=1, relief="solid", font=(s.font_family, 10),
        )
        for item in items:
            if item is None:
                menu.add_separator()
                continue
            label, command, *extra = item
            accelerator = extra[0] if extra else ""
            menu.add_command(label=label, command=command, accelerator=accelerator)
        button.configure(menu=menu)
        return button

    def _build_header(self, parent):
        """رأس مضغوط بقوائم منسدلة أعلى اليمين لزيادة مساحة التقرير."""
        s = self.settings
        header = tk.Frame(parent, bg=s.surface_color, height=116, highlightbackground="#DCE3EA", highlightthickness=1)
        header.grid(row=0, column=0, sticky="ew", padx=2, pady=(2, 2))
        header.grid_propagate(False)

        nav = tk.Frame(header, bg=s.primary_color, height=30)
        nav.pack(fill="x")
        nav.pack_propagate(False)

        menus = [
            ("ملف", [
                ("تقرير جديد - شاشة فارغة", self.start_new_report, "Ctrl+Shift+N"),
                ("عرض السجلات السابقة", self.show_all_records, ""),
                None,
                ("استيراد Excel", self.import_from_excel, "Ctrl+O"),
                ("استيراد من نسخة قديمة", self.import_from_old_backup, ""),
                None,
                ("نسخة احتياطية", self.create_backup, ""),
                ("استعادة نسخة احتياطية", self.restore_backup, ""),
                None,
                ("تسجيل خروج", self.logout, ""),
            ]),
            ("السجل", [
                ("سجل جديد", self.start_new_record, "Ctrl+N"),
                ("حفظ الآن", self.save_or_update_record, "Ctrl+S"),
                ("تعديل المحدد", self.edit_selected_record, ""),
                ("إفراغ حقول الإدخال", self.empty_form_for_new_report, ""),
                None,
                ("إضافة مرفقات", self.open_attachments, ""),
                ("ترقيم تلقائي", self.generate_contract_number, ""),
                ("تكرار السجل", self.duplicate_record, "Ctrl+D"),
                ("حذف المحدد", self.delete_selected, "Delete"),
            ]),
            ("التصدير والطباعة", [
                ("خيارات التصدير المتقدمة", self.open_export_options, ""),
                None,
                ("تصدير PDF عالي الجودة", lambda: self._export("pdf"), "Ctrl+P"),
                ("تصدير PNG عالي الدقة", lambda: self._export("png"), "Ctrl+I"),
                ("تصدير JPG عالي الجودة", lambda: self._export("jpg"), ""),
                ("تصدير Excel", lambda: self._export("xlsx"), "Ctrl+E"),
                ("تصدير CSV", lambda: self._export("csv"), "Ctrl+Shift+C"),
                ("تصدير HTML", lambda: self._export("html"), "Ctrl+H"),
                None,
                ("معاينة PDF", self.preview_current_report, ""),
                ("طباعة التقرير مباشرة", self.print_current_report, ""),
            ]),
            ("البيانات", [
                ("دليل العملاء والمعارض", self.open_directories, ""),
                ("سجل تسليمات العميل", self.open_customer_history, ""),
                ("تعديل جماعي للمحدد", self.bulk_edit_selected, ""),
                ("سلة المحذوفات", self.open_trash, ""),
                None,
                ("المستخدمون والصلاحيات", self.open_user_management, ""),
                ("تبديل المستخدم", self.switch_user, ""),
            ]),
            ("التقارير", [
                ("قوالب التقارير", self.open_report_templates, ""),
                ("تحليلات المبيعات والتسليم", self.open_dashboard, ""),
                ("أرشيف التقارير والبحث", self.open_report_archive, ""),
                ("سجل النشاط", self.open_activity_log, ""),
                ("مركز التكامل والاستمرارية", self.open_integration_center, ""),
            ]),
            ("عرض", [
                ("وضع تركيز التقرير", self.toggle_report_focus_mode, "F9"),
                ("ملء الشاشة", self.toggle_window_fullscreen, "F11"),
                None,
                ("إظهار/إخفاء لوحة الإدخال", self.toggle_form_panel, "F2"),
                ("إظهار/إخفاء البحث والفلاتر", self.toggle_filters_panel, "F3"),
                ("إظهار/إخفاء الإحصائيات", self.toggle_stats_panel, "F4"),
                ("توسيع مساحة التقرير", self.expand_report_workspace, ""),
                ("إظهار/إخفاء اللوحة الجانبية", self.toggle_sidebar, "Ctrl+M"),
                ("اختيار أعمدة السجلات", self._choose_table_columns, ""),
                ("تحديث البيانات", self.refresh_from_database, "F5"),
            ]),
            ("الإعدادات", [
                ("إعدادات البرنامج", self.open_settings, ""),
                ("فحص وإصلاح قاعدة البيانات", self.repair_database, ""),
                ("خيارات التصدير والطباعة", self.open_export_options, ""),
            ]),
        ]
        # تعبئة من اليمين لليسار؛ أول قائمة تبقى في أقصى اليمين.
        for title, items in menus:
            self._make_top_menu(nav, title, items).pack(side="right", padx=(1, 0), pady=1)

        # شريط عمليات مباشرة لا يغيّر صلاحيات المستخدم أو منطق حفظ البيانات.
        quick = tk.Frame(header, bg="#EEF3F8", height=30)
        quick.pack(fill="x", padx=1)
        quick.pack_propagate(False)
        actions = [
            ("+ سجل جديد", self.start_new_record, s.success_color),
            ("حفظ السجل", self.save_or_update_record, s.primary_color),
            ("استيراد Excel", self.import_from_excel, "#315F84"),
            ("تصدير PDF", lambda: self._export("pdf"), "#9B4D42"),
            ("تصدير Excel", lambda: self._export("xlsx"), "#267348"),
            ("نسخة احتياطية", self.create_backup, "#62598C"),
            ("تحديث", self.refresh_from_database, "#596A75"),
        ]
        for label, command, color in actions:
            tk.Button(
                quick, text=label, command=command, bg=color, fg="white",
                activebackground=s.accent_color, activeforeground="#15212A",
                relief="flat", bd=0, cursor="hand2", padx=9, pady=2,
                font=(s.font_family, 9, "bold"),
            ).pack(side="right", padx=(3, 1), pady=2)

        body = tk.Frame(header, bg=s.surface_color)
        body.pack(fill="both", expand=True, padx=6, pady=2)

        title = tk.Frame(body, bg=s.surface_color)
        title.pack(side="right", fill="y")
        tk.Label(title, text=s.report_title, bg=s.surface_color, fg=s.primary_color,
                 anchor="e", font=(s.font_family, max(16, s.heading_font_size-3), "bold")).pack(anchor="e")
        tk.Label(title, text=s.report_title_en, bg=s.surface_color, fg=s.accent_color,
                 anchor="e", font=("Segoe UI", 8, "bold")).pack(anchor="e")
        tk.Label(title, text="تقارير وتسليم وطباعة احترافية عالية الوضوح", bg=s.surface_color,
                 fg=s.muted_color, anchor="e", font=(s.font_family, 8)).pack(anchor="e")

        left = tk.Frame(body, bg=s.surface_color)
        left.pack(side="left", fill="y")
        recipient_row = tk.Frame(left, bg=s.surface_color)
        recipient_row.pack(anchor="w", pady=(0, 1))
        tk.Label(recipient_row, text="المستلم:", bg=s.surface_color, fg=s.primary_color,
                 font=(s.font_family, 8, "bold")).pack(side="right", padx=(4, 0))
        tk.Entry(recipient_row, textvariable=self.recipient_var, justify="right", width=20,
                 relief="solid", bd=1, font=(s.font_family, 8)).pack(side="right", ipady=1)
        tk.Label(left, text=f"المستخدم: {clean_value(self.current_user.get('display_name', ''))}", bg=s.surface_color, fg=s.primary_color, justify="left", font=(s.font_family, 8, "bold")).pack(anchor="w")
        tk.Label(left, textvariable=self.clock_var, bg=s.surface_color, fg=s.muted_color,
                 justify="left", font=(s.font_family, 8, "bold")).pack(anchor="w")

    def set_current_user(self, user: dict[str, Any], refresh_ui: bool = True):
        global CURRENT_USER
        self.current_user = dict(user or CURRENT_USER)
        CURRENT_USER = dict(self.current_user)
        if refresh_ui and hasattr(self, "main_frame"):
            self._build_ui(); self._apply_filters(); self.status_var.set(f"تم تسجيل الدخول باسم: {self.current_user.get('display_name','')}")

    def _has_permission(self, permission: str) -> bool:
        perms = set(self.current_user.get("permissions", []))
        role = self.current_user.get("role", "موظف")
        perms |= set(ROLE_PERMISSIONS.get(role, set()))
        return "*" in perms or permission in perms

    def _require_permission(self, permission: str, action_label: str) -> bool:
        if self._has_permission(permission):
            return True
        messagebox.showwarning("صلاحيات", f"المستخدم الحالي لا يملك صلاحية: {action_label}", parent=self.root)
        return False

    def _apply_panel_visibility(self):
        for attr, visible in (("stats_frame", self.stats_visible), ("filters_frame", self.filters_visible), ("form_frame", self.form_visible)):
            frame = getattr(self, attr, None)
            if frame is None: continue
            if visible: frame.grid()
            else: frame.grid_remove()

    def toggle_form_panel(self):
        self.form_visible = not self.form_visible; self._apply_panel_visibility(); self.status_var.set("تم إظهار لوحة الإدخال" if self.form_visible else "تم طي لوحة الإدخال لتوسيع التقرير")

    def toggle_filters_panel(self):
        self.filters_visible = not self.filters_visible; self._apply_panel_visibility()

    def toggle_stats_panel(self):
        self.stats_visible = not self.stats_visible; self._apply_panel_visibility()

    def toggle_report_focus_mode(self):
        self.report_focus_mode = not self.report_focus_mode
        if self.report_focus_mode:
            self._focus_saved = (self.form_visible, self.filters_visible, self.stats_visible, self.sidebar_visible)
            self.form_visible = self.filters_visible = self.stats_visible = False
            if self.sidebar_visible: self.toggle_sidebar(manual=False)
            self.status_var.set("وضع تركيز التقرير: مساحة السجلات بأقصى حجم")
        else:
            saved = getattr(self, "_focus_saved", (True, True, True, False))
            self.form_visible, self.filters_visible, self.stats_visible, side = saved
            if bool(side) != self.sidebar_visible: self.toggle_sidebar(manual=False)
            self.status_var.set("تم إنهاء وضع تركيز التقرير")
        self._apply_panel_visibility()

    def toggle_window_fullscreen(self):
        self.window_fullscreen = not self.window_fullscreen
        try: self.root.attributes("-fullscreen", self.window_fullscreen)
        except tk.TclError: pass

    def open_trash(self):
        if self._require_permission("records.restore", "استعادة السجلات المحذوفة"): TrashDialog(self)

    def open_directories(self):
        DirectoryDialog(self)

    def open_customer_history(self):
        CustomerHistoryDialog(self)

    def open_report_templates(self):
        ReportTemplateDialog(self)

    def open_user_management(self):
        if self._require_permission("*", "إدارة المستخدمين والصلاحيات"): UserManagementDialog(self)

    def switch_user(self):
        SwitchUserDialog(self)

    def bulk_edit_selected(self):
        if not self._require_permission("records.edit", "التعديل الجماعي"): return
        selected=list(self.tree.selection()) if hasattr(self,"tree") else []
        if not selected: messagebox.showwarning("التعديل الجماعي","حدد سجلين أو أكثر من الجدول.",parent=self.root); return
        BulkEditDialog(self, selected)

    def repair_database(self):
        if not self._require_permission("backup", "إصلاح قاعدة البيانات"): return
        try:
            result=DatabaseManager.repair_database(); messagebox.showinfo("فحص قاعدة البيانات", f"اكتمل الفحص والتحسين.\n{result}", parent=self.root); self.status_var.set("تم فحص قاعدة البيانات وتحسين الفهارس")
        except Exception as exc: messagebox.showerror("قاعدة البيانات",str(exc),parent=self.root)

    def _directory_values(self, key: str) -> list[str]:
        entity={"customer_name":"customer","showroom":"showroom","salesperson":"salesperson"}.get(key)
        if not entity:return []
        return [item["name"] for item in DatabaseManager.list_directory(entity)]

    def _refresh_directory_values(self):
        for key in ("customer_name","showroom","salesperson"):
            entry=self.entries.get(key)
            if isinstance(entry,ttk.Combobox): entry.configure(values=self._directory_values(key))

    def _directory_autofill(self, key: str):
        if key != "customer_name": return
        entry=self.entries.get("customer_name"); mobile_entry=self.entries.get("mobile")
        if entry is None or mobile_entry is None:return
        selected=clean_value(entry.get())
        item=next((x for x in DatabaseManager.list_directory("customer") if x["name"]==selected),None)
        if item and clean_value(item.get("mobile","")) and not clean_value(mobile_entry.get()):
            mobile_entry.delete(0,tk.END); mobile_entry.insert(0,item["mobile"])

    def expand_report_workspace(self):
        if self.sidebar_visible:
            self.toggle_sidebar()
        self.form_visible = False
        self.filters_visible = False
        self.stats_visible = False
        self._apply_panel_visibility()
        self.status_var.set("تم توسيع مساحة التقرير وإخفاء اللوحات غير الضرورية")

    def _sync_workspace_button(self):
        # بقيت للتوافق مع منطق إظهار/إخفاء اللوحة الجانبية.
        return

    def toggle_sidebar(self, manual=True):
        if not hasattr(self, "sidebar_frame"):
            return
        if manual:
            self.sidebar_manual_override = True
        self.sidebar_visible = not self.sidebar_visible
        if self.sidebar_visible:
            self.root.grid_columnconfigure(1, minsize=self.sidebar_width)
            self.sidebar_frame.grid()
            self.status_var.set("تم إظهار قائمة الأدوات")
        else:
            self.sidebar_frame.grid_remove()
            self.root.grid_columnconfigure(1, minsize=0)
            self.status_var.set("تم توسيع مساحة السجلات ومعاينة التقرير")
        self._sync_workspace_button()

    def _handle_root_resize(self, event):
        if event.widget is not self.root or self.sidebar_manual_override:
            return
        if self._resize_after_id:
            try:
                self.root.after_cancel(self._resize_after_id)
            except tk.TclError:
                pass
        self._resize_after_id = self.root.after(120, self._apply_compact_layout)

    def _apply_compact_layout(self):
        self._resize_after_id = None
        if self.root.winfo_width() < 1250 and self.sidebar_visible:
            self.toggle_sidebar(manual=False)
            self.status_var.set("تم تفعيل التخطيط المدمج تلقائياً لملاءمة حجم الشاشة")

    def _build_stats(self, parent):
        s = self.settings
        container = tk.Frame(parent, bg=s.background_color)
        container.grid(row=1, column=0, sticky="ew", padx=2, pady=(0, 2))
        self.stats_frame = container
        for col in range(6):
            container.grid_columnconfigure(col, weight=1)
        cards = [
            (0, "المتأخرة", "overdue", s.danger_color),
            (1, "تم التسليم", "delivered", s.success_color),
            (2, "جاهزة", "ready", "#2E6F9E"),
            (3, "سجلات اليوم", "today", s.accent_color),
            (4, "الظاهرة", "filtered", s.muted_color),
            (5, "الإجمالي", "total", s.primary_color),
        ]
        for col, label, key, color in cards:
            card = tk.Frame(container, bg=s.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
            card.grid(row=0, column=col, sticky="ew", padx=2, ipady=0)
            tk.Label(card, textvariable=self.stats_vars[key], bg=s.surface_color, fg=color,
                     font=(s.font_family, 16, "bold")).pack(pady=(0, 0))
            tk.Label(card, text=label, bg=s.surface_color, fg=s.muted_color,
                     font=(s.font_family, 8)).pack(pady=(0, 1))

    def _build_filters(self, parent):
        s = self.settings
        frame = tk.Frame(parent, bg=s.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
        frame.grid(row=2, column=0, sticky="ew", padx=2, pady=(0, 2))
        self.filters_frame = frame

        search_row = tk.Frame(frame, bg=s.surface_color)
        search_row.pack(fill="x", padx=5, pady=(3, 1))
        tk.Label(search_row, text="بحث شامل:", bg=s.surface_color, fg=s.text_color,
                 font=(s.font_family, s.base_font_size, "bold")).pack(side="right", padx=(4, 8))
        search_entry = tk.Entry(
            search_row, textvariable=self.search_var, justify="right", relief="solid", bd=1,
            font=(s.font_family, s.base_font_size), width=42,
        )
        search_entry.pack(side="right", fill="x", expand=True, ipady=2, padx=3)
        search_entry.bind("<Return>", lambda _event: self._apply_filters())
        self.search_entry = search_entry
        tk.Button(
            search_row, text="مسح ×", command=lambda: self.search_var.set(""),
            bg="#E7EEF4", fg=s.primary_color, relief="flat", bd=0,
            cursor="hand2", font=(s.font_family, 9, "bold"), padx=6,
        ).pack(side="right", padx=2, ipady=2)
        if self._search_trace_id:
            try:
                self.search_var.trace_remove("write", self._search_trace_id)
            except tk.TclError:
                pass
        self._search_trace_id = self.search_var.trace_add("write", self._schedule_filters)
        self.make_button(
            search_row, "بحث", self._apply_filters, s.primary_color, width=9, pady=2,
        ).pack(side="right", padx=4)
        tk.Button(
            search_row, text="إعادة ضبط", command=self._reset_filters,
            bg=s.muted_color, fg="white", relief="flat", cursor="hand2",
            font=(s.font_family, 9, "bold"),
        ).pack(side="left", padx=3, ipadx=6, ipady=2)

        filter_frame = tk.Frame(frame, bg=s.surface_color)
        filter_frame.pack(fill="x", padx=5, pady=1)
        tk.Label(filter_frame, text="فلترة حسب:", bg=s.surface_color, fg=s.text_color,
                 font=(s.font_family, s.base_font_size, "bold")).pack(side="right", padx=5)
        self.filter_field_var = tk.StringVar(value="الكل")
        preferred_keys = ["customer_name", "mobile", "contract_number"]
        ordered_fields = sorted(
            s.visible_fields,
            key=lambda item: (
                preferred_keys.index(item.key) if item.key in preferred_keys else len(preferred_keys),
                s.visible_fields.index(item),
            ),
        )
        filter_fields = ["الكل"] + [f.label for f in ordered_fields]
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_field_var, values=filter_fields,
                                    state="readonly", width=18, font=(s.font_family, s.base_font_size))
        filter_combo.pack(side="right", padx=5)
        filter_combo.bind("<<ComboboxSelected>>", self._on_filter_field_changed)
        self.filter_value_var = tk.StringVar()
        self.filter_value_combo = ttk.Combobox(
            filter_frame, textvariable=self.filter_value_var, justify="right",
            state="normal", font=(s.font_family, s.base_font_size), width=18,
        )
        self.filter_value_combo.pack(side="right", padx=4, ipady=2)
        self.filter_value_var.trace_add("write", self._schedule_filters)
        self.status_filter_var = tk.StringVar(value="كل الحالات")
        status_combo = ttk.Combobox(
            filter_frame, textvariable=self.status_filter_var,
            values=("كل الحالات", *STATUS_OPTIONS), state="readonly", width=13,
            font=(s.font_family, s.base_font_size),
        )
        status_combo.pack(side="right", padx=5)
        status_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_filters())

        date_frame = tk.Frame(frame, bg=s.surface_color)
        date_frame.pack(fill="x", padx=5, pady=(1, 3))
        tk.Label(
            date_frame, text="تاريخ الفلترة:", bg=s.surface_color, fg=s.text_color,
            font=(s.font_family, s.base_font_size, "bold"),
        ).pack(side="right", padx=5)
        self.filter_date_basis_var = tk.StringVar(value="التاريخ الأنسب")
        basis_combo = ttk.Combobox(
            date_frame, textvariable=self.filter_date_basis_var,
            values=("التاريخ الأنسب", "التركيب الفعلي", "تاريخ إنشاء السجل", "تاريخ آخر تعديل"),
            state="readonly", width=18, font=(s.font_family, 10),
        )
        basis_combo.pack(side="right", padx=5)
        basis_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_filters())
        self.filter_from_var = tk.StringVar()
        self.filter_to_var = tk.StringVar()
        for label, variable in (("من", self.filter_from_var), ("إلى", self.filter_to_var)):
            tk.Label(date_frame, text=label, bg=s.surface_color, fg=s.text_color).pack(side="right", padx=(7, 2))
            entry = tk.Entry(
                date_frame, textvariable=variable, justify="center", relief="solid", bd=1,
                width=12, font=(s.font_family, 10),
            )
            entry.pack(side="right", ipady=2)
            entry.bind("<Return>", lambda _event: self._apply_filters())
            tk.Button(
                date_frame, text="تقويم", command=lambda var=variable, title=label: self._pick_filter_date(var, title),
                bg="#E8EEF2", fg=s.primary_color, relief="flat", cursor="hand2",
                font=(s.font_family, 8, "bold"),
            ).pack(side="right", padx=1)
        tk.Label(
            date_frame, text="يمكن ترك أحد التاريخين فارغًا", bg=s.surface_color, fg=s.muted_color,
            font=(s.font_family, 9),
        ).pack(side="left", padx=6)
        self.make_button(
            date_frame, "تصدير نتائج الفلترة", self.open_export_options,
            s.success_color, width=15, pady=2,
        ).pack(side="left", padx=4)

    def _pick_filter_date(self, variable: tk.StringVar, label: str):
        picker = DatePickerDialog(self.root, variable.get(), f"اختيار تاريخ {label}", self.settings)
        self.root.wait_window(picker)
        if picker.result is not None:
            variable.set(picker.result)
            self._apply_filters()

    def _filter_date_bounds(self) -> tuple[Optional[dt.date], Optional[dt.date]]:
        from_var = getattr(self, "filter_from_var", None)
        to_var = getattr(self, "filter_to_var", None)
        from_text = clean_value(from_var.get()) if from_var is not None else ""
        to_text = clean_value(to_var.get()) if to_var is not None else ""
        try:
            start = dt.datetime.strptime(from_text, "%Y-%m-%d").date() if from_text else None
            end = dt.datetime.strptime(to_text, "%Y-%m-%d").date() if to_text else None
        except ValueError as exc:
            raise ValueError("استخدم صيغة التاريخ YYYY-MM-DD أو اختر التاريخ من زر التقويم.") from exc
        if start and end and start > end:
            raise ValueError("تاريخ البداية يجب ألا يكون بعد تاريخ النهاية.")
        return start, end

    def _schedule_filters(self, *_args):
        if self._filter_after_id:
            try:
                self.root.after_cancel(self._filter_after_id)
            except tk.TclError:
                pass
        self._filter_after_id = self.root.after(180, self._apply_filters)

    def _workspace_records(self) -> list[dict[str, Any]]:
        """السجلات المسموح بعرضها في مساحة التقرير الحالية."""
        if self.current_report_record_ids is None:
            return self.records
        allowed = self.current_report_record_ids
        return [record for record in self.records if str(record.get("id", "")) in allowed]

    def _sync_report_workspace_ui(self):
        if self.current_report_record_ids is None:
            self.report_scope_var.set("عرض جميع السجلات")
        else:
            count = len(self._workspace_records())
            self.report_scope_var.set(f"التقرير الحالي: {count} سجل")

    def _has_user_form_data(self) -> bool:
        return any(
            clean_value(entry.get()) and not (
                key == "delivery_status" and clean_value(entry.get()) == "جديد"
            )
            for key, entry in getattr(self, "entries", {}).items()
        )

    def _save_before_workspace_change(self, title: str, message: str) -> bool:
        self._cancel_record_autosave()
        if self._has_user_form_data() and not self._autosave_current_record(refresh=False):
            return messagebox.askyesno(title, message, parent=self.root)
        return True

    def start_new_report(self):
        """يفتح تقريرًا مستقلًا وفارغًا دون حذف السجلات السابقة من قاعدة البيانات."""
        if not self._save_before_workspace_change(
            "تقرير جديد",
            "البيانات الحالية غير مكتملة ولا يمكن حفظها كسجل. سيتم إفراغها من الشاشة؛ هل تريد فتح تقرير جديد فارغ؟",
        ):
            return
        self.current_report_record_ids = set()
        self._form_undo_stack.clear()
        self._reset_filters()
        self.recipient_var.set("")
        self.clear_form()
        self._on_filter_field_changed()
        self._apply_filters()
        DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
        DatabaseManager.log_action("إنشاء تقرير جديد", details="مساحة عمل فارغة مستقلة")
        self.status_var.set("تم فتح تقرير جديد فارغ تمامًا - السجلات السابقة محفوظة في الأرشيف")

    def show_all_records(self):
        """يعود إلى الأرشيف الكامل بعد حفظ السجل الحالي إن كان مكتملًا."""
        if not self._save_before_workspace_change(
            "عرض السجلات السابقة",
            "المسودة الحالية غير مكتملة وسيتم إفراغها من الشاشة. هل تريد عرض السجلات السابقة؟",
        ):
            return
        self.current_report_record_ids = None
        self._form_undo_stack.clear()
        self._reset_filters()
        self.clear_form()
        self._on_filter_field_changed()
        self._apply_filters()
        DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
        self.status_var.set("تم عرض جميع السجلات السابقة المحفوظة")

    def _on_filter_field_changed(self, _event=None):
        selected_label = self.filter_field_var.get()
        selected_key = next(
            (item.key for item in self.settings.visible_fields if item.label == selected_label), "",
        )
        values = []
        if selected_key:
            values = sorted({
                clean_value(record.get("values", {}).get(selected_key, ""))
                for record in self._workspace_records()
                if clean_value(record.get("values", {}).get(selected_key, ""))
            })
        if hasattr(self, "filter_value_combo"):
            self.filter_value_combo.configure(values=values)
        self._apply_filters()

    def _apply_filters(self):
        self._filter_after_id = None
        try:
            start_date, end_date = self._filter_date_bounds()
        except ValueError as exc:
            messagebox.showwarning("فلترة التاريخ", str(exc), parent=self.root)
            return
        query = normalize_search_text(self.search_var.get())
        field_label = self.filter_field_var.get()
        field_value = normalize_search_text(self.filter_value_var.get())
        selected_key = ""
        if field_label != "الكل" and field_value:
            for f in self.settings.visible_fields:
                if f.label == field_label:
                    selected_key = f.key
                    break
        exact_field_values = {
            normalize_search_text(record.get("values", {}).get(selected_key, ""))
            for record in self._workspace_records()
            if selected_key and clean_value(record.get("values", {}).get(selected_key, ""))
        }
        use_exact_field_match = bool(selected_key and field_value in exact_field_values)
        selected_status = self.status_filter_var.get() if hasattr(self, "status_filter_var") else "كل الحالات"
        date_basis = (
            self.filter_date_basis_var.get()
            if hasattr(self, "filter_date_basis_var") else "التاريخ الأنسب"
        )
        filtered = []
        for record in self._workspace_records():
            values = record.get("values", {})
            if query and query not in normalize_search_text(" ".join(str(value) for value in values.values())):
                continue
            if selected_key:
                normalized_record_value = normalize_search_text(values.get(selected_key, ""))
                if use_exact_field_match and normalized_record_value != field_value:
                    continue
                if not use_exact_field_match and field_value not in normalized_record_value:
                    continue
            record_status = clean_value(values.get("delivery_status", "")) or "جديد"
            if selected_status != "كل الحالات" and record_status != selected_status:
                continue
            if start_date or end_date:
                record_date = record_date_for_basis(record, date_basis)
                if record_date is None:
                    continue
                if start_date and record_date < start_date:
                    continue
                if end_date and record_date > end_date:
                    continue
            filtered.append(record)
        self.filtered_records = filtered
        self._refresh_table()
        self._update_stats()
        self._sync_report_workspace_ui()

    def _reset_filters(self):
        self.search_var.set("")
        self.filter_field_var.set("الكل")
        self.filter_value_var.set("")
        if hasattr(self, "status_filter_var"):
            self.status_filter_var.set("كل الحالات")
        if hasattr(self, "filter_date_basis_var"):
            self.filter_date_basis_var.set("التاريخ الأنسب")
        if hasattr(self, "filter_from_var"):
            self.filter_from_var.set("")
        if hasattr(self, "filter_to_var"):
            self.filter_to_var.set("")
        self._apply_filters()

    def _build_form(self, parent):
        s = self.settings
        card = tk.Frame(parent, bg=s.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
        card.grid(row=3, column=0, sticky="ew", padx=2, pady=(0, 2))
        self.form_frame = card
        top = tk.Frame(card, bg=s.surface_color)
        top.pack(fill="x", padx=6, pady=(1, 0))
        tk.Label(top, text="بيانات التسليم", bg=s.surface_color, fg=s.primary_color,
                 font=(s.font_family, 11, "bold")).pack(side="right")
        self.mode_label = tk.Label(top, text="إضافة سجل جديد", bg=s.surface_color, fg=s.success_color,
                                   font=(s.font_family, 9, "bold"))
        self.mode_label.pack(side="left")
        tk.Label(
            top, text="Enter: حفظ وفتح سجل جديد  |  Tab: الحقل التالي  |  Esc: تراجع",
            bg=s.surface_color, fg=s.muted_color, font=(s.font_family, 8),
        ).pack(side="left", padx=12)

        fields_frame = tk.Frame(card, bg=s.surface_color)
        fields_frame.pack(fill="x", padx=4, pady=0)
        field_columns = 6 if self.root.winfo_screenwidth() >= 1200 else 4
        for col in range(field_columns):
            fields_frame.grid_columnconfigure(col, weight=1)
        self.entries = {}
        visible_fields = [item for item in s.visible_fields if item.key != "scheduled_delivery_date"]
        self._entry_order = [item.key for item in visible_fields]
        for idx, item in enumerate(visible_fields):
            row = idx // field_columns
            col = (field_columns - 1) - (idx % field_columns)
            holder = tk.Frame(fields_frame, bg=s.surface_color)
            holder.grid(row=row, column=col, sticky="ew", padx=3, pady=1)
            marker = " *" if item.required else ""
            tk.Label(holder, text=f"{item.label}{marker}", bg=s.surface_color,
                     fg=s.text_color, anchor="e", font=(s.font_family, 8, "bold")).pack(fill="x")
            if item.kind == "choice":
                options = item.options or (list(STATUS_OPTIONS) if item.key == "delivery_status" else [])
                entry = ttk.Combobox(holder, values=options, state="readonly", justify="right",
                                     font=(s.font_family, s.base_font_size))
                if item.key == "delivery_status" and options:
                    entry.set("جديد")
            elif item.key in {"customer_name", "showroom", "salesperson"}:
                entry = ttk.Combobox(holder, values=self._directory_values(item.key), state="normal", justify="right",
                                     font=(s.font_family, s.base_font_size))
                entry.bind("<<ComboboxSelected>>", lambda _event, key=item.key: self._directory_autofill(key), add="+")
            else:
                date_row = tk.Frame(holder, bg=s.surface_color) if item.kind == "date" else None
                entry = tk.Entry(date_row or holder, justify="right", relief="solid", bd=1,
                                 highlightthickness=1, highlightbackground="#D6DEE7",
                                 highlightcolor=s.accent_color,
                                 bg="#FFF9E8" if item.required else s.surface_color,
                                 font=(s.font_family, s.base_font_size))
            if item.kind == "date":
                date_row.pack(fill="x", pady=(1, 0))
                entry.pack(side="right", fill="x", expand=True, ipady=1)
                tk.Button(
                    date_row, text="تقويم", command=lambda target=entry, label=item.label: self._pick_date(target, label),
                    bg=s.primary_color, fg="white", activebackground=s.accent_color, activeforeground="#111111",
                    relief="flat", bd=0, cursor="hand2", padx=6, pady=2,
                    font=(s.font_family, 9, "bold"),
                ).pack(side="left", padx=(4, 0))
            else:
                entry.pack(fill="x", pady=(0, 0), ipady=1)
            entry.bind("<KeyPress>", self._capture_form_undo, add="+")
            entry.bind("<Button-1>", self._capture_form_undo, add="+")
            entry.bind("<Return>", lambda event, index=idx: self._handle_enter(event, index))
            entry.bind("<KP_Enter>", lambda event, index=idx: self._handle_enter(event, index))
            entry.bind("<Shift-Return>", lambda event, index=idx: self._handle_shift_enter(event, index))
            entry.bind("<KeyRelease>", self._schedule_record_autosave, add="+")
            entry.bind("<FocusOut>", self._schedule_record_autosave, add="+")
            if isinstance(entry, ttk.Combobox):
                entry.bind("<<ComboboxSelected>>", self._schedule_record_autosave, add="+")
            self.entries[item.key] = entry


    def _focus_entry_index(self, index: int):
        if not self._entry_order:
            return
        index = max(0, min(index, len(self._entry_order) - 1))
        entry = self.entries.get(self._entry_order[index])
        if entry is None:
            return
        entry.focus_set()
        try:
            entry.selection_range(0, tk.END)
            entry.icursor(tk.END)
        except tk.TclError:
            pass

    def _handle_enter(self, _event, index: int):
        """Enter يعتمد السجل الحالي ثم يفتح نموذجًا جديدًا جاهزًا للإدخال."""
        self.save_and_start_new_record()
        return "break"

    def _handle_shift_enter(self, _event, index: int):
        self._focus_entry_index(index - 1)
        return "break"

    def _form_snapshot(self) -> dict[str, Any]:
        values: dict[str, str] = {}
        for key, entry in getattr(self, "entries", {}).items():
            try:
                values[key] = clean_value(entry.get())
            except tk.TclError:
                pass
        return {"record_id": self.selected_record_id, "values": values}

    def _capture_form_undo(self, event=None, force: bool = False):
        if self._form_undo_busy or self._record_autosave_busy or not getattr(self, "entries", None):
            return
        if event is not None and not force:
            keysym = str(getattr(event, "keysym", ""))
            ignored = {
                "Return", "KP_Enter", "Escape", "F5", "Tab", "ISO_Left_Tab",
                "Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next",
                "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
            }
            if keysym in ignored or (int(getattr(event, "state", 0)) & 0x4):
                return
            if keysym not in {"BackSpace", "Delete"} and not str(getattr(event, "char", "")):
                return
        snapshot = self._form_snapshot()
        signature = json.dumps(snapshot, ensure_ascii=False, sort_keys=True)
        if self._form_undo_stack:
            previous = json.dumps(self._form_undo_stack[-1], ensure_ascii=False, sort_keys=True)
            if previous == signature:
                return
        self._form_undo_stack.append(snapshot)
        del self._form_undo_stack[:-100]

    def _undo_form(self, _event=None):
        """يتراجع خطوة واحدة داخل نموذج الإدخال دون حذف أي سجل من قاعدة البيانات."""
        self._cancel_record_autosave()
        if not self._form_undo_stack:
            self.status_var.set("لا توجد خطوة إدخال سابقة للتراجع عنها")
            return "break"
        snapshot = self._form_undo_stack.pop()
        values = snapshot.get("values", {}) if isinstance(snapshot, dict) else {}
        record_id = snapshot.get("record_id") if isinstance(snapshot, dict) else None
        self._form_undo_busy = True
        self._record_autosave_busy = True
        try:
            for key, entry in self.entries.items():
                value = clean_value(values.get(key, ""))
                if isinstance(entry, ttk.Combobox):
                    entry.set(value or ("جديد" if key == "delivery_status" else ""))
                else:
                    entry.delete(0, tk.END)
                    entry.insert(0, value)
            self.selected_record_id = record_id if record_id and self._record_by_id(record_id) else None
            if hasattr(self, "tree"):
                if self.selected_record_id and self.tree.exists(self.selected_record_id):
                    self._ignore_tree_select_once = True
                    self.tree.selection_set(self.selected_record_id)
                    self.tree.see(self.selected_record_id)
                else:
                    self.tree.selection_remove(self.tree.selection())
            if self.selected_record_id:
                self.mode_label.configure(text="وضع التعديل - تم التراجع", fg=self.settings.accent_color)
            else:
                self.mode_label.configure(text="إضافة سجل جديد - تم التراجع", fg=self.settings.success_color)
        finally:
            self._record_autosave_busy = False
            self._form_undo_busy = False
        DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
        self._schedule_record_autosave()
        self.status_var.set("تم التراجع عن آخر تعديل في حقول الإدخال")
        return "break"

    def _pick_date(self, entry: tk.Entry, label: str):
        self._capture_form_undo(force=True)
        picker = DatePickerDialog(self.root, entry.get(), f"اختيار {label}", self.settings)
        self.root.wait_window(picker)
        if picker.result is None:
            return
        entry.delete(0, tk.END)
        entry.insert(0, picker.result)
        entry.focus_set()
        self._schedule_record_autosave()

    def _build_table(self, parent):
        s = self.settings
        card = tk.Frame(parent, bg=s.surface_color, highlightbackground="#DCE3EA", highlightthickness=1)
        card.grid(row=4, column=0, sticky="nsew", padx=2, pady=(0, 2))
        toolbar = tk.Frame(card, bg=s.surface_color)
        toolbar.pack(fill="x", padx=5, pady=1)
        tk.Label(toolbar, text="السجلات", bg=s.surface_color, fg=s.primary_color,
                 font=(s.font_family, 12, "bold")).pack(side="right")
        tk.Label(
            toolbar, textvariable=self.report_scope_var, bg="#EAF2F8", fg=s.primary_color,
            padx=7, pady=2, font=(s.font_family, 8, "bold"),
        ).pack(side="right", padx=6)
        tk.Label(
            toolbar, textvariable=self.selection_count_var, bg=s.surface_color,
            fg=s.muted_color, font=(s.font_family, 9, "bold"),
        ).pack(side="left", padx=7)

        table_frame = tk.Frame(card, bg=s.surface_color)
        table_frame.pack(fill="both", expand=True, padx=3, pady=(0, 3))
        fields = [item for item in s.table_fields if item.key != "scheduled_delivery_date"]
        columns = tuple(f.key for f in fields)
        self.tree = ttk.Treeview(
            table_frame, columns=columns, displaycolumns=tuple(reversed(columns)),
            show="headings", selectmode="extended", style="Modern.Treeview",
        )
        for f in fields:
            direction = " ↓" if self.table_sort_desc else " ↑"
            label = f.label + (direction if f.key == self.table_sort_key else "")
            self.tree.heading(f.key, text=label, command=lambda key=f.key: self._sort_table_by(key))
            self.tree.column(f.key, width=f.width, minwidth=80, anchor="center", stretch=True)
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<<TreeviewSelect>>", self._update_selection_count, add="+")
        self.tree.bind("<Double-1>", lambda e: self._on_select(None))
        self.tree.bind("<Button-3>", self._show_tree_context_menu)
        self.tree.bind("<ButtonRelease-1>", self._schedule_column_width_save, add="+")
        self.tree.tag_configure("even", background=s.surface_color)
        self.tree.tag_configure("odd", background="#F5F8FA")
        self.tree.tag_configure("delivered", background="#E7F4EC")
        self.tree.tag_configure("ready", background="#EAF2F8")
        self.tree.tag_configure("revision", background="#FFF0B8")
        self.tree.tag_configure("postponed", background="#FFF5D9")
        self.tree.tag_configure("cancelled", background="#FDE3E3")
        self.tree.tag_configure("overdue", background="#FDECEC")
        self.tree_context_menu = tk.Menu(self.root, tearoff=False, font=(s.font_family, 10))
        self.tree_context_menu.add_command(label="تعديل المحدد", command=self.edit_selected_record)
        self.tree_context_menu.add_command(label="تكرار كسجل جديد", command=self.duplicate_record)
        self.tree_context_menu.add_separator()
        self.tree_context_menu.add_command(label="نسخ السجلات  Ctrl+C", command=self.copy_selected_records)
        self.tree_context_menu.add_command(label="لصق كسجلات جديدة  Ctrl+V", command=self.paste_records_from_clipboard)
        self.root.after_idle(lambda: self.tree.xview_moveto(1.0) if self.tree.winfo_exists() else None)

    def _update_selection_count(self, _event=None):
        if hasattr(self, "tree") and hasattr(self, "selection_count_var"):
            self.selection_count_var.set(f"المحدد: {len(self.tree.selection())}")

    def _show_tree_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            if row_id not in self.tree.selection():
                self.tree.selection_set(row_id)
            self.tree.focus(row_id)
        try:
            self.tree_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tree_context_menu.grab_release()

    def _choose_table_columns(self):
        fields = [item for item in self.settings.visible_fields if item.key != "scheduled_delivery_date"]
        dialog = tk.Toplevel(self.root)
        dialog.title("اختيار أعمدة جدول السجلات")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.settings.surface_color)
        width = 720
        rows = max(1, (len(fields) + 2) // 3)
        height = min(650, max(300, 190 + rows * 45))
        x = self.root.winfo_x() + max(0, (self.root.winfo_width() - width) // 2)
        y = self.root.winfo_y() + max(0, (self.root.winfo_height() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.minsize(600, 300)

        tk.Label(
            dialog, text="اختر الأعمدة التي تريد ظهورها في جدول السجلات",
            bg=self.settings.surface_color, fg=self.settings.primary_color,
            font=(self.settings.font_family, 14, "bold"), anchor="e",
        ).pack(fill="x", padx=20, pady=(18, 4))
        tk.Label(
            dialog,
            text="إخفاء العمود هنا لا يحذف بياناته ولا يغيّر التصدير أو الطباعة أو مقاساتهما.",
            bg=self.settings.surface_color, fg=self.settings.muted_color,
            font=(self.settings.font_family, 10), anchor="e", justify="right",
        ).pack(fill="x", padx=20, pady=(0, 12))

        choices = tk.Frame(dialog, bg=self.settings.surface_color)
        choices.pack(fill="both", expand=True, padx=20, pady=4)
        for column in range(3):
            choices.grid_columnconfigure(column, weight=1)

        hidden = set(self.settings.hidden_table_columns)
        variables: dict[str, tk.BooleanVar] = {}
        for index, item in enumerate(fields):
            variable = tk.BooleanVar(value=item.key not in hidden)
            variables[item.key] = variable
            row = index // 3
            column = 2 - (index % 3)
            tk.Checkbutton(
                choices, text=item.label, variable=variable,
                bg=self.settings.surface_color, fg=self.settings.text_color,
                activebackground=self.settings.surface_color,
                font=(self.settings.font_family, 10, "bold"), anchor="e",
            ).grid(row=row, column=column, sticky="ew", padx=10, pady=8)

        def save_columns():
            selected = {key for key, variable in variables.items() if variable.get()}
            if not selected:
                messagebox.showwarning(
                    "الأعمدة", "يجب إبقاء عمود واحد على الأقل ظاهرًا في الجدول.", parent=dialog,
                )
                return
            visible_keys = {item.key for item in fields}
            retained_hidden = [
                key for key in self.settings.hidden_table_columns if key not in visible_keys
            ]
            newly_hidden = [item.key for item in fields if item.key not in selected]
            self.settings.hidden_table_columns = retained_hidden + newly_hidden
            try:
                DatabaseManager.save_settings(self.settings)
            except Exception as exc:
                messagebox.showerror("خطأ", f"تعذر حفظ اختيار الأعمدة: {exc}", parent=dialog)
                return
            dialog.destroy()
            self._build_ui()
            self._apply_filters()
            self.status_var.set("تم حفظ أعمدة الجدول دون تغيير التصدير أو الطباعة")

        buttons = tk.Frame(dialog, bg=self.settings.surface_color)
        buttons.pack(fill="x", padx=20, pady=(8, 18))
        self.make_button(
            buttons, "حفظ", save_columns, self.settings.success_color, width=12,
        ).pack(side="right", padx=5)
        self.make_button(
            buttons, "إلغاء", dialog.destroy, self.settings.muted_color, width=12,
        ).pack(side="right", padx=5)

    def _sort_table_by(self, key: str):
        """فرز بصري تنازلي/تصاعدي، دون تغيير ترتيب البيانات أو ملفات التصدير."""
        if self.table_sort_key == key:
            self.table_sort_desc = not self.table_sort_desc
        else:
            self.table_sort_key = key
            self.table_sort_desc = False
        fields = {item.key: item for item in self.settings.table_fields}
        for field in self.settings.table_fields:
            if field.key in self.tree["columns"]:
                arrow = (" ↓" if self.table_sort_desc else " ↑") if field.key == key else ""
                self.tree.heading(field.key, text=field.label + arrow)
        self._refresh_table()
        self.status_var.set(f"تم فرز السجلات حسب {fields[key].label}" if key in fields else "تم فرز السجلات")

    def _refresh_table(self):
        if not hasattr(self, "tree"):
            return
        fields = self.settings.table_fields
        table_records = self.filtered_records
        if self.table_sort_key and self.table_sort_key in {item.key for item in fields}:
            table_records = sorted(
                self.filtered_records,
                key=lambda rec: report_table_sort_value(rec.get("values", {}).get(self.table_sort_key, "")),
                reverse=self.table_sort_desc,
            )
        desired_ids = [str(record["id"]) for record in table_records]
        existing_ids = list(self.tree.get_children())
        desired_set = set(desired_ids)
        for iid in existing_ids:
            if iid not in desired_set:
                self.tree.delete(iid)
        for idx, rec in enumerate(table_records):
            values = [rec.get("values", {}).get(f.key, "") for f in fields]
            status = clean_value(rec.get("values", {}).get("delivery_status", "")) or "جديد"
            tag = "overdue" if record_is_overdue(rec) else {
                "تم التسليم": "delivered", "جاهز للتسليم": "ready",
                "تعديل - يعتمد الحالي ويلغى السابق": "revision",
                "مؤجل": "postponed", "ملغي": "cancelled",
            }.get(status, "even" if idx % 2 == 0 else "odd")
            iid = str(rec["id"])
            if self.tree.exists(iid):
                current = tuple(str(value) for value in self.tree.item(iid, "values"))
                normalized = tuple(str(value) for value in values)
                if current != normalized or tuple(self.tree.item(iid, "tags")) != (tag,):
                    self.tree.item(iid, values=values, tags=(tag,))
                if self.tree.index(iid) != idx:
                    self.tree.move(iid, "", idx)
            else:
                self.tree.insert("", idx, iid=iid, values=values, tags=(tag,))

    def _schedule_column_width_save(self, _event=None):
        if self._column_save_after_id:
            try:
                self.root.after_cancel(self._column_save_after_id)
            except tk.TclError:
                pass
        self._column_save_after_id = self.root.after(700, self._save_column_widths)

    def _save_column_widths(self):
        self._column_save_after_id = None
        if not hasattr(self, "tree"):
            return
        changed = False
        for item in self.settings.fields:
            if item.key in self.tree["columns"]:
                width = max(80, min(320, int(self.tree.column(item.key, "width"))))
                if item.width != width:
                    item.width = width
                    changed = True
        if changed:
            try:
                DatabaseManager.save_settings(self.settings)
            except Exception:
                logging.exception("تعذر حفظ عرض الأعمدة")

    def _update_stats(self):
        today = dt.date.today().isoformat()
        workspace_records = self._workspace_records()
        total = len(workspace_records)
        today_count = sum(1 for r in workspace_records if str(r.get("created_at", "")).startswith(today))
        filtered = len(self.filtered_records)
        ready = sum(1 for r in workspace_records if (clean_value(r.get("values", {}).get("delivery_status", "")) or "جديد") == "جاهز للتسليم")
        delivered = sum(1 for r in workspace_records if clean_value(r.get("values", {}).get("delivery_status", "")) == "تم التسليم")
        overdue = sum(1 for r in workspace_records if record_is_overdue(r))
        self.stats_vars["total"].set(str(total))
        self.stats_vars["today"].set(str(today_count))
        self.stats_vars["filtered"].set(str(filtered))
        self.stats_vars["ready"].set(str(ready))
        self.stats_vars["delivered"].set(str(delivered))
        self.stats_vars["overdue"].set(str(overdue))
        scope = "التقرير الحالي" if self.current_report_record_ids is not None else "جميع السجلات"
        active_filters = []
        query = clean_value(self.search_var.get())
        if query:
            active_filters.append(f"بحث شامل: {query}")
        if hasattr(self, "filter_field_var") and clean_value(self.filter_value_var.get()):
            active_filters.append(f"{self.filter_field_var.get()}: {clean_value(self.filter_value_var.get())}")
        if hasattr(self, "status_filter_var") and self.status_filter_var.get() != "كل الحالات":
            active_filters.append(f"الحالة: {self.status_filter_var.get()}")
        if hasattr(self, "filter_from_var") and clean_value(self.filter_from_var.get()):
            active_filters.append(f"من {clean_value(self.filter_from_var.get())}")
        if hasattr(self, "filter_to_var") and clean_value(self.filter_to_var.get()):
            active_filters.append(f"إلى {clean_value(self.filter_to_var.get())}")
        suffix = f" | الفلاتر: {'، '.join(active_filters)}" if active_filters else ""
        self.status_var.set(f"{scope}: نتائج مطابقة {filtered} من {total} سجل{suffix}")

    def _entry_values(self):
        vals = {}
        for f in self.settings.fields:
            if f.key in self.entries:
                vals[f.key] = clean_value(self.entries[f.key].get())
            elif self.selected_record_id:
                rec = self._record_by_id(self.selected_record_id)
                vals[f.key] = clean_value(rec.get("values", {}).get(f.key, "")) if rec else ""
            else:
                vals[f.key] = ""
        return vals

    def _record_by_id(self, rid):
        return next((r for r in self.records if r.get("id") == rid), None)

    def _prepare_revision_replacement(self, current: dict[str, Any]) -> Optional[dict[str, Any]]:
        """يلغي آخر سجل سابق للعقد عند اختيار حالة التعديل، من دون حذف أي بيانات."""
        values = current.get("values", {})
        if clean_value(values.get("delivery_status", "")) != "تعديل - يعتمد الحالي ويلغى السابق":
            return None
        if clean_value(values.get("_revision_replaces_record_id", "")):
            return None
        contract = clean_value(values.get("contract_number", ""))
        if not contract:
            return None
        current_id = str(current.get("id", ""))
        candidates = [
            record for record in self.records
            if str(record.get("id", "")) != current_id
            and clean_value(record.get("values", {}).get("contract_number", "")) == contract
            and clean_value(record.get("values", {}).get("delivery_status", "")) != "ملغي"
        ]
        if not candidates:
            return None
        current_created = clean_value(current.get("created_at", ""))
        older = [
            record for record in candidates
            if not current_created or clean_value(record.get("created_at", "")) <= current_created
        ]
        previous = max(
            older or candidates,
            key=lambda record: (
                clean_value(record.get("updated_at", "")),
                clean_value(record.get("created_at", "")),
            ),
        )
        now = dt.datetime.now().isoformat(timespec="seconds")
        previous_values = previous.setdefault("values", {})
        previous_values["delivery_status"] = "ملغي"
        previous_values["_revision_replaced_by"] = current_id
        previous["updated_at"] = now
        values["_revision_replaces_record_id"] = str(previous.get("id", ""))
        values["_revision_applied_at"] = now
        return previous

    def _validate_and_confirm_duplicate(self, values, current_id=None):
        errors = []
        for f in self.settings.visible_fields:
            val = clean_value(values.get(f.key, ""))
            if f.required and not val:
                errors.append(f"حقل «{f.label}» مطلوب.")
            if f.kind == "phone" and val:
                digits = re.sub(r"\D", "", val)
                if not re.fullmatch(r"[+\d\s()\-]+", val) or not 7 <= len(digits) <= 15:
                    errors.append(f"صيغة «{f.label}» غير صحيحة.")
            if f.kind == "number" and val:
                try:
                    float(val.replace(",", ""))
                except ValueError:
                    errors.append(f"حقل «{f.label}» يجب أن يكون رقمًا.")
            if f.kind == "date" and val:
                try:
                    dt.datetime.strptime(val, "%Y-%m-%d")
                except ValueError:
                    errors.append(f"حقل «{f.label}» يجب أن يكون تاريخًا بصيغة YYYY-MM-DD.")
            if len(val) > 250:
                errors.append(f"حقل «{f.label}» يتجاوز 250 حرفًا.")
        if errors:
            messagebox.showwarning("تحقق", "\n".join(f"• {e}" for e in errors), parent=self.root)
            return False
        scheduled = clean_value(values.get("scheduled_delivery_date", ""))
        actual = clean_value(values.get("actual_delivery_date", ""))
        if scheduled and actual:
            try:
                if dt.datetime.strptime(actual, "%Y-%m-%d").date() < dt.datetime.strptime(scheduled, "%Y-%m-%d").date():
                    if not messagebox.askyesno(
                        "مراجعة التاريخ",
                        "تاريخ التركيب الفعلي أسبق من موعد التسليم المطلوب. هل تريد المتابعة؟",
                        parent=self.root,
                    ):
                        return False
            except ValueError:
                pass
        return True

    def _persist_record(self, record) -> bool:
        try:
            replaced_record = self._prepare_revision_replacement(record)
            DatabaseManager.save_record(record)
            if replaced_record is not None:
                DatabaseManager.save_record(replaced_record)
                DatabaseManager.log_action(
                    "اعتماد تعديل وإلغاء السابق",
                    record.get("id", ""),
                    f"العقد: {clean_value(record.get('values', {}).get('contract_number', ''))} | "
                    f"السجل السابق: {replaced_record.get('id', '')}",
                )
            existing = self._record_by_id(record["id"])
            if existing:
                existing.update(record)
            else:
                self.records.append(record)
            if self.current_report_record_ids is not None:
                self.current_report_record_ids.add(str(record["id"]))
            self._apply_filters()
            self.status_var.set(
                "تم اعتماد السجل الحالي وإلغاء السجل السابق دون حذفه"
                if replaced_record is not None else "تم الحفظ في قاعدة البيانات"
            )
            return True
        except Exception as e:
            logging.exception("فشل الحفظ")
            messagebox.showerror("خطأ", f"تعذر الحفظ: {e}", parent=self.root)
            return False

    def _cancel_record_autosave(self):
        if self._record_autosave_after_id is not None:
            try:
                self.root.after_cancel(self._record_autosave_after_id)
            except tk.TclError:
                pass
            self._record_autosave_after_id = None

    def _schedule_record_autosave(self, *_args):
        if self._record_autosave_busy:
            return
        self._cancel_record_autosave()
        try:
            self._record_autosave_after_id = self.root.after(1600, self._autosave_current_record)
        except tk.TclError:
            self._record_autosave_after_id = None

    def _values_ready_for_autosave(self, values: dict[str, Any]) -> bool:
        for item in self.settings.visible_fields:
            value = clean_value(values.get(item.key, ""))
            if item.required and not value:
                return False
            if item.kind == "date" and value:
                try:
                    dt.datetime.strptime(value, "%Y-%m-%d")
                except ValueError:
                    return False
            if item.kind == "phone" and value:
                digits = re.sub(r"\D", "", value)
                if not re.fullmatch(r"[+\d\s()\-]+", value) or not 7 <= len(digits) <= 15:
                    return False
        return True

    def _autosave_current_record(self, refresh: bool = True) -> bool:
        self._record_autosave_after_id = None
        if self._record_autosave_busy or not getattr(self, "entries", None):
            return False
        self._record_autosave_busy = True
        try:
            values = self._entry_values()
            if not clean_value(values.get("delivery_status", "")):
                values["delivery_status"] = "جديد"
            current = self._record_by_id(self.selected_record_id) if self.selected_record_id else None
            if current is None:
                ensure_kitchen_identity(values, self.records)
            if values.get("delivery_status") == "تم التسليم" and not values.get("actual_delivery_date"):
                values["actual_delivery_date"] = dt.date.today().isoformat()
                actual_entry = self.entries.get("actual_delivery_date")
                if actual_entry is not None:
                    actual_entry.delete(0, tk.END)
                    actual_entry.insert(0, values["actual_delivery_date"])
            if not self._values_ready_for_autosave(values):
                DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
                self.status_var.set("حُفظت المسودة تلقائيًا؛ يكتمل حفظ السجل بعد تعبئة الحقول المطلوبة")
                return False

            snapshot = json.dumps(values, ensure_ascii=False, sort_keys=True)
            if current and snapshot == json.dumps(current.get("values", {}), ensure_ascii=False, sort_keys=True):
                self._last_record_snapshot = f"{current['id']}:{snapshot}"
                return True

            now = dt.datetime.now().isoformat(timespec="seconds")
            if current:
                current["values"] = values
                current["updated_at"] = now
                record = current
                action_text = "تم تحديث السجل وحفظه تلقائيًا في قاعدة البيانات"
            else:
                record = {"id": uuid.uuid4().hex, "created_at": now, "updated_at": now, "values": values}
                action_text = "تم إنشاء السجل وحفظه تلقائيًا في قاعدة البيانات"
            replaced_record = self._prepare_revision_replacement(record)
            DatabaseManager.save_record(record)
            if replaced_record is not None:
                DatabaseManager.save_record(replaced_record)
                action_text = "تم اعتماد السجل الحالي وإلغاء السجل السابق تلقائيًا دون حذفه"
            if current is None:
                self.records.append(record)
                self.selected_record_id = record["id"]
            if self.current_report_record_ids is not None:
                self.current_report_record_ids.add(str(record["id"]))
            self._last_record_snapshot = f"{record['id']}:{snapshot}"
            if refresh:
                self._apply_filters()
                if hasattr(self, "tree") and self.tree.exists(record["id"]):
                    self.tree.selection_set(record["id"])
                    self.tree.see(record["id"])
            self.mode_label.configure(text="محفوظ تلقائيًا - وضع التعديل", fg=self.settings.success_color)
            self.status_var.set(action_text)
            return True
        except Exception:
            logging.exception("فشل الحفظ التلقائي للسجل")
            self.status_var.set("تعذر الحفظ التلقائي؛ استخدم زر «حفظ الآن»")
            return False
        finally:
            self._record_autosave_busy = False

    def save_or_update_record(self):
        permission = "records.edit" if self.selected_record_id else "records.create"
        if not self._require_permission(permission, "حفظ السجل"):
            return False
        self._cancel_record_autosave()
        if self.selected_record_id:
            return self.update_record()
        return self.save_record()

    def save_and_start_new_record(self):
        """يحفظ الإضافة أو التعديل ثم ينتقل دائمًا إلى سجل جديد عند النجاح."""
        permission = "records.edit" if self.selected_record_id else "records.create"
        if not self._require_permission(permission, "حفظ السجل"):
            return False
        was_editing = bool(self.selected_record_id)
        if not self.save_or_update_record():
            return False
        if was_editing:
            self.clear_form()
            self._apply_filters()
        self.status_var.set("تم اعتماد السجل وفتح سجل جديد جاهز للإدخال")
        return True

    def start_new_record(self):
        if not self._require_permission("records.create", "إنشاء سجل جديد"):
            return
        self._cancel_record_autosave()
        has_user_data = any(
            clean_value(entry.get()) and not (key == "delivery_status" and clean_value(entry.get()) == "جديد")
            for key, entry in self.entries.items()
        )
        if has_user_data and not self._autosave_current_record():
            if not messagebox.askyesno(
                "سجل جديد",
                "المسودة الحالية غير مكتملة ولا يمكن حفظها كسجل بعد. هل تريد إفراغها وبدء سجل جديد؟",
                parent=self.root,
            ):
                return
        self._capture_form_undo(force=True)
        self.clear_form()
        self.status_var.set("سجل جديد جاهز للإدخال - سيُحفظ تلقائيًا بعد اكتمال الحقول المطلوبة")

    def empty_form_for_new_report(self):
        """يحفظ السجل المكتمل ثم يفرغ النموذج صراحةً لإدخال تقرير جديد."""
        self._cancel_record_autosave()
        has_user_data = any(
            clean_value(entry.get()) and not (
                key == "delivery_status" and clean_value(entry.get()) == "جديد"
            )
            for key, entry in self.entries.items()
        )
        if has_user_data and not self._autosave_current_record():
            if not messagebox.askyesno(
                "إفراغ الحقول",
                "البيانات الحالية غير مكتملة ولا يمكن حفظها كسجل. هل تريد إفراغ الحقول نهائيًا والبدء بتقرير جديد؟",
                parent=self.root,
            ):
                return
        self._capture_form_undo(force=True)
        self.clear_form()
        DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
        self.status_var.set("تم إفراغ جميع الحقول - النموذج جاهز لإدخال تقرير جديد")

    def save_record(self):
        self._cancel_record_autosave()
        values = self._entry_values()
        if not values.get("delivery_status"):
            values["delivery_status"] = "جديد"
        if values.get("delivery_status") == "تم التسليم" and not values.get("actual_delivery_date"):
            values["actual_delivery_date"] = dt.date.today().isoformat()
        ensure_kitchen_identity(values, self.records)
        if not self._validate_and_confirm_duplicate(values):
            return False
        now = dt.datetime.now().isoformat(timespec="seconds")
        rec = {"id": uuid.uuid4().hex, "created_at": now, "updated_at": now, "values": values}
        if not self._persist_record(rec):
            return False
        self._form_undo_stack.clear()
        self.clear_form()
        self._apply_filters()
        return True

    def update_record(self):
        self._cancel_record_autosave()
        if not self.selected_record_id:
            messagebox.showwarning("تنبيه", "حدد سجلًا لتحديثه.", parent=self.root)
            return False
        rec = self._record_by_id(self.selected_record_id)
        if not rec:
            return False
        values = self._entry_values()
        if values.get("delivery_status") == "تم التسليم" and not values.get("actual_delivery_date"):
            values["actual_delivery_date"] = dt.date.today().isoformat()
        if not self._validate_and_confirm_duplicate(values, self.selected_record_id):
            return False
        previous = copy.deepcopy(rec)
        rec["values"] = values
        rec["updated_at"] = dt.datetime.now().isoformat(timespec="seconds")
        if not self._persist_record(rec):
            rec.clear()
            rec.update(previous)
            return False
        self._form_undo_stack.clear()
        self._apply_filters()
        return True

    def edit_selected_record(self):
        """ينقل السجل المختار بوضوح إلى وضع التعديل ثم يركز أول حقل."""
        selected = list(self.tree.selection()) if hasattr(self, "tree") else []
        if not selected:
            messagebox.showwarning("تعديل السجل", "حدد سجلًا من الجدول أولًا.", parent=self.root)
            return
        self.tree.selection_set(selected[0])
        self._on_select(None)
        self.mode_label.configure(text="تعديل السجل المحدد", fg=self.settings.accent_color)
        self._focus_entry_index(0)
        self.status_var.set("أنت الآن في وضع التعديل؛ غيّر البيانات ثم اضغط «حفظ الآن»")

    def duplicate_record(self):
        if not self._require_permission("records.create", "تكرار السجل"):
            return
        if not self.selected_record_id:
            messagebox.showwarning("تكرار السجل", "حدد سجلًا لتكراره كسجل جديد.", parent=self.root)
            return
        rec = self._record_by_id(self.selected_record_id)
        if not rec:
            return
        now = dt.datetime.now().isoformat(timespec="seconds")
        dup = {"id": uuid.uuid4().hex, "created_at": now, "updated_at": now,
               "values": copy.deepcopy(rec.get("values", {}))}
        ensure_kitchen_identity(dup["values"], self.records)
        dup["values"]["delivery_status"] = "جديد"
        dup["values"]["actual_delivery_date"] = ""
        if not self._persist_record(dup):
            return
        self._form_undo_stack.clear()
        self._apply_filters()
        if hasattr(self, "tree") and self.tree.exists(dup["id"]):
            self.tree.selection_set(dup["id"])
            self.tree.see(dup["id"])
            self._on_select(None)
        self.mode_label.configure(text="سجل مكرر جديد - وضع التعديل", fg=self.settings.accent_color)
        self.status_var.set("تم تكرار السجل مع الاحتفاظ برقم العقد والتفاصيل وتمييزه داخليًا")

    def copy_selected_records(self, _event=None):
        """ينسخ السجلات المختارة كجدول يمكن لصقه في التطبيق أو Excel."""
        selected = list(self.tree.selection()) if hasattr(self, "tree") else []
        if not selected:
            messagebox.showwarning("نسخ السجلات", "حدد سجلًا أو أكثر من الجدول.", parent=self.root)
            return "break"
        fields = self.settings.fields
        selected_ids = set(selected)
        records = [record for record in self.filtered_records if str(record.get("id", "")) in selected_ids]
        buffer = io.StringIO(newline="")
        writer = csv.writer(buffer, delimiter="\t", lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
        writer.writerow([item.label for item in fields])
        for record in records:
            values = record.get("values", {})
            writer.writerow([clean_value(values.get(item.key, "")) for item in fields])
        self.root.clipboard_clear()
        self.root.clipboard_append(buffer.getvalue())
        self.root.update_idletasks()
        self.status_var.set(f"تم نسخ {len(records)} سجل؛ يمكن لصقها هنا أو في Excel")
        return "break"

    def paste_records_from_clipboard(self, _event=None):
        """يلصق جدولًا من التطبيق أو Excel كسجلات جديدة مع قبول العقود المتكررة."""
        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("لصق السجلات", "الحافظة لا تحتوي نصًا قابلًا للصق.", parent=self.root)
            return "break"
        try:
            rows = [
                [clean_value(value) for value in row]
                for row in csv.reader(io.StringIO(clipboard_text), delimiter="\t")
                if any(clean_value(value) for value in row)
            ]
        except csv.Error as exc:
            messagebox.showerror("لصق السجلات", f"تعذر قراءة محتوى الحافظة: {exc}", parent=self.root)
            return "break"
        if not rows:
            messagebox.showwarning("لصق السجلات", "لا توجد صفوف بيانات في الحافظة.", parent=self.root)
            return "break"

        fields = self.settings.fields
        normalized_fields: dict[str, str] = {}
        for item in fields:
            normalized_fields[SpreadsheetImportReader.normalize_header(item.key)] = item.key
            normalized_fields[SpreadsheetImportReader.normalize_header(item.label)] = item.key
            for alias in SpreadsheetImportReader.FIELD_ALIASES.get(item.key, []):
                normalized_fields[SpreadsheetImportReader.normalize_header(alias)] = item.key
        mapping = {
            index: normalized_fields[SpreadsheetImportReader.normalize_header(value)]
            for index, value in enumerate(rows[0])
            if SpreadsheetImportReader.normalize_header(value) in normalized_fields
        }
        if mapping:
            data_rows = rows[1:]
        elif all(len(row) == len(fields) for row in rows):
            mapping = {index: item.key for index, item in enumerate(fields)}
            data_rows = rows
        else:
            messagebox.showwarning(
                "لصق السجلات",
                "لم أتعرف على عناوين الأعمدة. انسخ صف العناوين مع البيانات من Excel، أو استخدم «نسخ السجلات» من التطبيق.",
                parent=self.root,
            )
            return "break"

        required_keys = {item.key for item in self.settings.visible_fields if item.required}
        accepted: list[dict[str, Any]] = []
        skipped = 0
        now = dt.datetime.now().isoformat(timespec="seconds")
        for row in data_rows:
            values = {item.key: "" for item in fields}
            for column_index, field_key in mapping.items():
                if column_index < len(row):
                    values[field_key] = clean_value(row[column_index])
            if not any(values.values()):
                continue
            if any(not clean_value(values.get(key, "")) for key in required_keys):
                skipped += 1
                continue
            values["delivery_status"] = clean_value(values.get("delivery_status", "")) or "جديد"
            ensure_kitchen_identity(values, [*self.records, *accepted])
            accepted.append({
                "id": uuid.uuid4().hex,
                "created_at": now,
                "updated_at": now,
                "values": values,
            })
        if not accepted:
            messagebox.showinfo(
                "لصق السجلات",
                "لم تُلصق سجلات؛ تأكد من وجود رقم العقد واسم العميل ورقم الجوال في كل صف.",
                parent=self.root,
            )
            return "break"
        if not messagebox.askyesno(
            "تأكيد لصق السجلات",
            f"سيتم إنشاء {len(accepted)} سجل جديد.\n"
            "العقود المتكررة مسموحة وستُميّز السجلات داخليًا.\n"
            f"الصفوف غير المكتملة: {skipped}\n\nهل تريد المتابعة؟",
            parent=self.root,
        ):
            return "break"
        try:
            DatabaseManager.save_records_bulk(accepted)
            self.records.extend(accepted)
            if self.current_report_record_ids is not None:
                self.current_report_record_ids.update(str(record["id"]) for record in accepted)
            self._apply_filters()
            self.status_var.set(f"تم لصق {len(accepted)} سجل جديد بنجاح")
        except Exception as exc:
            logging.exception("فشل لصق السجلات")
            messagebox.showerror("لصق السجلات", f"تعذر حفظ السجلات: {exc}", parent=self.root)
        return "break"

    def delete_selected(self):
        if not self._require_permission("records.delete", "حذف السجلات"): return
        selected = list(self.tree.selection()) if hasattr(self, "tree") else []
        if not selected:
            messagebox.showwarning("تنبيه", "حدد سجلًا أو أكثر للحذف.", parent=self.root)
            return
        if not messagebox.askyesno("تأكيد", f"نقل {len(selected)} سجل إلى سلة المحذوفات؟", parent=self.root):
            return
        try:
            DatabaseManager.delete_records(selected)
            self.records = [r for r in self.records if r.get("id") not in selected]
            if self.current_report_record_ids is not None:
                self.current_report_record_ids.difference_update(str(item) for item in selected)
            self._form_undo_stack.clear()
            self.clear_form()
            self._apply_filters()
            self.status_var.set(f"تم نقل {len(selected)} سجل إلى سلة المحذوفات")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل الحذف: {e}", parent=self.root)

    def clear_form(self):
        self._cancel_record_autosave()
        self._record_autosave_busy = True
        try:
            for key, entry in self.entries.items():
                if isinstance(entry, ttk.Combobox):
                    entry.set("جديد" if key == "delivery_status" else "")
                else:
                    entry.delete(0, tk.END)
            self.selected_record_id = None
            self._last_record_snapshot = ""
            if hasattr(self, "tree"):
                self.tree.selection_remove(self.tree.selection())
            self.mode_label.configure(text="إضافة سجل جديد", fg=self.settings.success_color)
            first = next(iter(self.entries.values()), None)
            if first:
                first.focus_set()
        finally:
            self._record_autosave_busy = False

    def _on_select(self, event):
        if self._ignore_tree_select_once:
            self._ignore_tree_select_once = False
            return
        if self._record_autosave_busy:
            return
        selected = self.tree.selection()
        if not selected:
            return
        next_record_id = selected[0]
        self._capture_form_undo(force=True)
        if self.selected_record_id and self.selected_record_id != next_record_id:
            self._cancel_record_autosave()
            self._autosave_current_record(refresh=False)
        self.selected_record_id = next_record_id
        rec = self._record_by_id(self.selected_record_id)
        if not rec:
            return
        values = rec.get("values", {})
        self._record_autosave_busy = True
        try:
            for key, entry in self.entries.items():
                value = values.get(key, "")
                if isinstance(entry, ttk.Combobox):
                    entry.set(value or ("جديد" if key == "delivery_status" else ""))
                else:
                    entry.delete(0, tk.END)
                    entry.insert(0, value)
        finally:
            self._record_autosave_busy = False
        self._last_record_snapshot = f"{rec['id']}:" + json.dumps(values, ensure_ascii=False, sort_keys=True)
        self.mode_label.configure(text="وضع التعديل", fg=self.settings.accent_color)

    def generate_contract_number(self):
        entry = self.entries.get("contract_number")
        if not entry:
            return
        current = clean_value(entry.get())
        if current and not messagebox.askyesno(
            "استبدال رقم العقد", f"استبدال الرقم الحالي «{current}» برقم تلقائي؟", parent=self.root
        ):
            return
        self._capture_form_undo(force=True)
        number = DatabaseManager.next_contract_number(self.settings.auto_contract_prefix)
        entry.delete(0, tk.END)
        entry.insert(0, number)
        self.status_var.set(f"تم توليد رقم العقد {number}")

    def open_attachments(self):
        if not self._require_permission("attachments", "المرفقات"):
            return
        if not self.selected_record_id:
            self._cancel_record_autosave()
            if not self._autosave_current_record():
                messagebox.showinfo(
                    "المرفقات",
                    "أكمل الحقول المطلوبة أولاً؛ سيُحفظ السجل تلقائيًا ثم يمكنك إضافة مرفقاته.",
                    parent=self.root,
                )
                return
        AttachmentsDialog(self, self.selected_record_id)

    def open_selected_whatsapp(self):
        record = self._record_by_id(self.selected_record_id) if self.selected_record_id else None
        values = record.get("values", {}) if record else self._entry_values()
        mobile = re.sub(r"\D", "", clean_value(values.get("mobile", "")))
        if mobile.startswith("00"):
            mobile = mobile[2:]
        elif mobile.startswith("05") and len(mobile) == 10:
            mobile = "966" + mobile[1:]
        if not 8 <= len(mobile) <= 15:
            messagebox.showwarning(
                "رسالة واتساب", "أدخل رقم جوال صحيحًا مع مفتاح الدولة أو رقمًا سعوديًا يبدأ بـ05.",
                parent=self.root,
            )
            return
        customer = clean_value(values.get("customer_name")) or "العميل الكريم"
        contract = clean_value(values.get("contract_number")) or "غير محدد"
        status = clean_value(values.get("delivery_status")) or "جديد"
        scheduled = clean_value(values.get("scheduled_delivery_date")) or "غير محدد"
        actual = clean_value(values.get("actual_delivery_date")) or "لم يسجل بعد"
        message = (
            f"مرحبًا {customer}،\n"
            f"تحديث طلب الرخام رقم {contract}:\n"
            f"الحالة: {status}\nموعد التسليم: {scheduled}\nتاريخ التركيب الفعلي: {actual}\n"
            f"مع تحيات {self.settings.company_name}"
        )
        webbrowser.open(f"https://wa.me/{mobile}?text={urllib.parse.quote(message)}")
        DatabaseManager.log_action("فتح رسالة واتساب", self.selected_record_id or "", contract)
        self.status_var.set("تم فتح رسالة واتساب الجاهزة - الإرسال النهائي بيد المستخدم")

    # ==================== دوال التصدير ====================
    def _start_export_task(self, file_type, records, recipient, selected_fields, output_path,
                           *, after_success=None, show_success=True):
        """يشغّل إنشاء التقارير بعيداً عن واجهة Tk حتى تبقى مستجيبة."""
        output_path = Path(output_path)
        records_snapshot = copy.deepcopy(list(records))
        fields_snapshot = copy.deepcopy(list(selected_fields))
        settings_snapshot = copy.deepcopy(self.settings)
        progress = ProgressDialog(self, f"جاري إنشاء تقرير {file_type.upper()}...")
        self.status_var.set("جارٍ إنشاء التقرير دون تجميد الواجهة...")

        def worker():
            try:
                exporter = ReportExporter(settings_snapshot)
                method = getattr(exporter, EXPORT_METHOD_NAMES[file_type])
                generated = method(records_snapshot, recipient, output_path, fields_snapshot)
                generated_paths = [Path(item) for item in generated] if isinstance(generated, (list, tuple)) else [output_path]
                DatabaseManager.archive_report(
                    output_path, file_type, settings_snapshot, recipient, records_snapshot, fields_snapshot
                )
                self.root.after(0, lambda paths=generated_paths: finish(None, paths))
            except Exception as exc:
                logging.exception("فشل التصدير")
                message = str(exc)
                self.root.after(0, lambda: finish(message, []))

        def finish(error_message, generated_paths):
            try:
                progress.close()
            except tk.TclError:
                pass
            if error_message:
                self.status_var.set("فشل التصدير")
                messagebox.showerror("تعذر إنشاء التقرير", error_message, parent=self.root)
                return
            DatabaseManager.log_action(
                f"تصدير {file_type.upper()}",
                details=f"{output_path.name} | عدد السجلات: {len(records_snapshot)}",
            )
            self.status_var.set(f"تم التصدير: {output_path}")
            if after_success:
                try:
                    after_success(output_path)
                except Exception as exc:
                    logging.exception("فشل الإجراء التالي للتصدير")
                    messagebox.showerror("تعذر إكمال العملية", str(exc), parent=self.root)
                    return
            if show_success:
                if len(generated_paths) > 1:
                    messagebox.showinfo(
                        "اكتمل",
                        f"تم إنشاء {len(generated_paths)} صفحات صور عالية الدقة في المجلد:\n{output_path.parent}\n\n"
                        f"الصفحة الأولى: {output_path.name}",
                        parent=self.root,
                    )
                else:
                    messagebox.showinfo("اكتمل", f"تم حفظ التقرير في:\n{output_path}", parent=self.root)

        threading.Thread(target=worker, daemon=True).start()

    def preview_current_report(self):
        """ينشئ معاينة PDF بنفس مقاس الطباعة الفعلي ويفتحها دون التأثير على السجلات."""
        if not self._require_permission("export", "معاينة التقرير"):
            return
        if not self.filtered_records:
            messagebox.showwarning("لا توجد بيانات", "لا توجد سجلات قابلة للمعاينة.", parent=self.root)
            return
        fields = self.settings.report_fields
        output = PREVIEW_DIR / f"معاينة_تقرير_تسليم_الرخام_{dt.datetime.now():%Y%m%d_%H%M%S}.pdf"
        self._start_export_task(
            "pdf", self.filtered_records, self.recipient_var.get(), fields, output,
            after_success=open_with_default_app, show_success=False,
        )

    def print_current_report(self):
        if not self._require_permission("export", "طباعة التقرير"):
            return
        if not self.filtered_records:
            messagebox.showwarning("لا توجد بيانات", "لا توجد سجلات قابلة للطباعة.", parent=self.root)
            return
        fields = self.settings.report_fields
        output = PREVIEW_DIR / f"طباعة_تقرير_تسليم_الرخام_{dt.datetime.now():%Y%m%d_%H%M%S}.pdf"

        def send_to_printer(path):
            print_pdf_file(path)
            self.status_var.set("تم إرسال التقرير إلى الطابعة")
            messagebox.showinfo("الطباعة", "تم إرسال التقرير إلى الطابعة الافتراضية.", parent=self.root)

        self._start_export_task(
            "pdf", self.filtered_records, self.recipient_var.get(), fields, output,
            after_success=send_to_printer, show_success=False,
        )

    def open_dashboard(self):
        OperationsDashboardDialog(self)

    def open_integration_center(self):
        IntegrationCenterDialog(self)

    def open_activity_log(self):
        ActivityLogDialog(self)

    def open_report_archive(self):
        ReportArchiveDialog(self)

    def _export(self, file_type, selected_fields=None, records=None, recipient=""):
        if not self._require_permission("export", "تصدير التقارير"): return
        if records is None:
            records = self.filtered_records
        if not records:
            messagebox.showwarning("لا توجد بيانات", "لا توجد سجلات للتصدير.", parent=self.root)
            return
        if recipient == "":
            recipient = self.recipient_var.get()

        now = dt.datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("_%H%M") if self.settings.export_preferences.get("include_time", False) else ""
        prefix = self.settings.export_preferences.get("file_name_prefix", "تقرير تسليم الرخام")
        file_name = f"{prefix} - {date_str}{time_str}"

        ext_map = {"xlsx": ".xlsx", "pdf": ".pdf", "png": ".png", "jpg": ".jpg", "csv": ".csv", "html": ".html"}
        filetypes = {
            "xlsx": [("Excel", "*.xlsx")],
            "pdf": [("PDF", "*.pdf")],
            "png": [("PNG", "*.png")],
            "jpg": [("JPG", "*.jpg *.jpeg")],
            "csv": [("CSV", "*.csv")],
            "html": [("HTML", "*.html")]
        }
        path = filedialog.asksaveasfilename(
            parent=self.root, title="حفظ التقرير",
            initialfile=f"{file_name}{ext_map[file_type]}",
            defaultextension=ext_map[file_type],
            filetypes=filetypes[file_type],
            initialdir=self.settings.default_export_folder if self.settings.default_export_folder and Path(self.settings.default_export_folder).is_dir() else None
        )
        if not path:
            return

        if selected_fields is None:
            selected_fields = self.settings.report_fields
        after = open_with_default_app if self.settings.export_preferences.get("open_after_export", False) else None
        self._start_export_task(file_type, records, recipient, selected_fields, Path(path), after_success=after)

    # ==================== استيراد من نسخة قديمة ====================
    def import_from_old_backup(self):
        if not self._require_permission("import", "استيراد نسخة قديمة"):
            return
        path = filedialog.askopenfilename(parent=self.root, title="اختيار نسخة احتياطية قديمة",
                                          filetypes=[("JSON", "*.json"), ("جميع", "*.*")])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if "records" not in data:
                messagebox.showerror("خطأ", "الملف المحدد ليس نسخة احتياطية صالحة.", parent=self.root)
                return
            OldImportDialog(self, Path(path))
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل فتح الملف: {e}", parent=self.root)

    # ==================== دوال أخرى ====================
    def open_settings(self):
        if not self._require_permission("*", "إعدادات البرنامج"): return
        SettingsDialog(self, self.settings, self._apply_settings)

    def _apply_settings(self, new_settings):
        if new_settings.require_login:
            users = DatabaseManager.list_users(active_only=True)
            if not any(user.get("password_hash") and user.get("salt") for user in users):
                messagebox.showerror(
                    "تفعيل الدخول",
                    "يجب تعيين كلمة مرور للمسؤول أولاً من إدارة المستخدمين قبل تفعيل تسجيل الدخول.",
                    parent=self.root,
                )
                return
        self.settings = new_settings
        try:
            DatabaseManager.save_settings(new_settings)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر حفظ الإعدادات: {e}", parent=self.root)
            return
        self._configure_styles()
        self._build_ui()
        self._apply_filters()
        self.status_var.set("تم تطبيق الإعدادات")

    def import_from_excel(self):
        if not self._require_permission("import", "استيراد Excel"): return
        path = filedialog.askopenfilename(
            parent=self.root,
            title="استيراد من Excel أو CSV",
            filetypes=[
                ("ملفات Excel الحديثة", "*.xlsx *.xlsm"),
                ("ملفات Excel القديمة", "*.xls"),
                ("ملفات CSV", "*.csv"),
                ("جميع الملفات المدعومة", "*.xlsx *.xlsm *.xls *.csv"),
            ],
        )
        if not path:
            return

        reader = SpreadsheetImportReader(self.settings)

        def do_import():
            try:
                preview = reader.read(Path(path))
                self.root.after(0, lambda p=preview: self._finish_excel_import(reader, p))
            except Exception as exc:
                message = str(exc)
                self.root.after(0, lambda m=message: self._excel_import_failed(m))

        self._show_progress("جاري تحليل الملف...")
        threading.Thread(target=do_import, daemon=True).start()

    def _excel_import_failed(self, message: str):
        self._hide_progress()
        logging.error("فشل استيراد الملف: %s", message)
        messagebox.showerror("تعذر الاستيراد", message, parent=self.root)
        self.status_var.set("تعذر الاستيراد")

    def _finish_excel_import(self, reader: SpreadsheetImportReader, preview: SpreadsheetImportPreview):
        self._hide_progress()
        if not preview.data_rows:
            messagebox.showinfo("لا توجد بيانات", "تم اكتشاف صف العناوين لكن لا توجد صفوف بيانات قابلة للاستيراد.", parent=self.root)
            return
        dialog = ManualMappingDialog(
            self,
            preview.headers,
            self.settings.fields,
            suggested=preview.suggested_mapping,
            sample_rows=preview.data_rows[:5],
        )
        if not dialog.mapping:
            self.status_var.set("تم إلغاء الاستيراد")
            return

        candidates = reader.records_from_preview(preview, dialog.mapping)
        accepted: list[dict[str, Any]] = []
        for record in candidates:
            ensure_kitchen_identity(record.get("values", {}), [*self.records, *accepted])
            accepted.append(record)

        mapped_labels = [
            item.label for item in self.settings.fields if item.key in set(dialog.mapping.values())
        ]
        summary = [
            f"الملف: {preview.source_path.name}",
            f"الورقة: {preview.sheet_name}",
            f"صف العناوين: {preview.header_row}",
            f"الحقول المرتبطة: {len(mapped_labels)} ({'، '.join(mapped_labels)})",
            f"السجلات الجاهزة: {len(accepted)}",
            "العقود المتكررة: مسموحة كسجلات مستقلة مع تمييز داخلي تلقائي",
        ]
        if not accepted:
            messagebox.showinfo("لا توجد سجلات جديدة", "\n".join(summary), parent=self.root)
            return
        if not messagebox.askyesno(
            "تأكيد استيراد البيانات",
            "\n".join(summary) + "\n\nهل تريد إضافة هذه السجلات إلى قاعدة البيانات؟",
            parent=self.root,
        ):
            self.status_var.set("تم إلغاء الاستيراد")
            return

        progress = ProgressDialog(self, "جاري حفظ السجلات المستوردة...")
        try:
            DatabaseManager.save_records_bulk(accepted)
            self.records = DatabaseManager.load_records()
            if self.current_report_record_ids is not None:
                self.current_report_record_ids.update(str(record["id"]) for record in accepted)
            self._apply_filters()
            self.status_var.set(f"تم استيراد {len(accepted)} سجل بنجاح")
            messagebox.showinfo(
                "اكتمل الاستيراد",
                f"تم استيراد {len(accepted)} سجل بنجاح.\n"
                "تم الاحتفاظ بجميع العقود المتكررة كسجلات مستقلة.",
                parent=self.root,
            )
        except Exception as exc:
            logging.exception("فشل حفظ السجلات المستوردة")
            messagebox.showerror("تعذر الحفظ", str(exc), parent=self.root)
        finally:
            progress.close()

    def _show_progress(self, text="جاري العمل..."):
        self.progress = ProgressDialog(self, text)

    def _hide_progress(self):
        if hasattr(self, "progress"):
            try:
                self.progress.close()
            except tk.TclError:
                pass

    def create_backup(self):
        if not self._require_permission("backup", "إنشاء نسخة احتياطية"): return
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
        path = filedialog.asksaveasfilename(parent=self.root, title="حفظ نسخة احتياطية",
                                            initialfile=f"نسخة_احتياطية_كاملة_{stamp}.zip",
                                            defaultextension=".zip",
                                            filetypes=[("نسخة كاملة ZIP", "*.zip"), ("JSON قديم", "*.json")])
        if not path:
            return
        try:
            self._write_backup(Path(path))
            DatabaseManager.log_action("إنشاء نسخة احتياطية", details=Path(path).name)
            messagebox.showinfo("نسخة احتياطية", f"تم حفظ النسخة في:\n{path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("خطأ", str(e), parent=self.root)

    def _backup_payload(self, include_attachments=True):
        attachment_items = []
        if include_attachments:
            for item in DatabaseManager.all_attachments():
                stored_path = Path(item["stored_path"])
                copied = dict(item)
                copied["archive_name"] = f"attachments/{item['record_id']}/{stored_path.name}"
                copied.pop("stored_path", None)
                attachment_items.append(copied)
        return {
            "app_version": APP_VERSION,
            "schema_version": SCHEMA_VERSION,
            "exported_at": dt.datetime.now().isoformat(timespec="seconds"),
            "settings": self.settings.to_dict(),
            "records": self.records,
            "attachments": attachment_items,
            "reports": DatabaseManager.report_archive_for_backup() if include_attachments else [],
            "directories": [item for kind in ("customer", "showroom", "salesperson") for item in DatabaseManager.list_directory(kind)],
            "report_templates": DatabaseManager.list_report_templates(),
            "users": DatabaseManager.list_users(),
        }

    def _write_backup(self, output: Path):
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.suffix.lower() == ".json":
            output.write_text(json.dumps(self._backup_payload(False), ensure_ascii=False, indent=2), encoding="utf-8")
            return
        payload = self._backup_payload(True)
        all_attachments = {item["id"]: item for item in DatabaseManager.all_attachments()}
        attachment_sources: dict[str, Path] = {}
        report_bytes: dict[str, bytes] = {}
        manifest: dict[str, dict[str, Any]] = {}
        for item in payload["attachments"]:
            source = Path(all_attachments[item["id"]]["stored_path"])
            if source.is_file():
                attachment_sources[item["archive_name"]] = source
                manifest[item["archive_name"]] = {
                    "sha256": sha256_file(source), "size": source.stat().st_size,
                }
        for item in payload.get("reports", []):
            archived = DatabaseManager.get_archived_report(item.get("id", ""))
            if archived and archived.get("file_data") is not None:
                data = bytes(archived["file_data"])
                report_bytes[item["archive_name"]] = data
                manifest[item["archive_name"]] = {
                    "sha256": hashlib.sha256(data).hexdigest(), "size": len(data),
                }
        payload["file_manifest"] = manifest
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            archive.writestr("backup.json", json.dumps(payload, ensure_ascii=False, indent=2))
            for archive_name, source in attachment_sources.items():
                archive.write(source, archive_name)
            for archive_name, data in report_bytes.items():
                archive.writestr(archive_name, data)

    def restore_backup(self):
        if not self._require_permission("backup", "استعادة النسخ الاحتياطية"): return
        path = filedialog.askopenfilename(parent=self.root, title="اختيار نسخة احتياطية",
                                          filetypes=[("نسخة كاملة", "*.zip"), ("JSON", "*.json"), ("جميع", "*.*")])
        if not path:
            return
        safety_copy: Path | None = None
        try:
            validation = validate_backup_archive(Path(path))
            archive = None
            if Path(path).suffix.lower() == ".zip":
                archive = zipfile.ZipFile(path, "r")
                data = json.loads(archive.read("backup.json").decode("utf-8"))
            else:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            if not isinstance(data, dict) or "records" not in data:
                raise ValueError("ملف غير صحيح")
            if not messagebox.askyesno(
                "تأكيد الاستعادة",
                "اجتازت النسخة فحص السلامة.\n"
                f"السجلات: {validation['records']} | المرفقات: {validation['attachments']} | التقارير: {validation['reports']}\n\n"
                "سيتم استبدال البيانات الحالية. هل تريد المتابعة؟",
                parent=self.root,
            ):
                if archive is not None:
                    archive.close()
                return
            safety_copy = BACKUP_DIR / f"قبل_الاستعادة_{dt.datetime.now():%Y%m%d_%H%M%S}.db"
            if DB_FILE.exists():
                shutil.copy2(DB_FILE, safety_copy)
            with get_db() as conn:
                conn.execute("DELETE FROM records")
            settings = AppSettings.from_dict(data.get("settings", {}))
            DatabaseManager.save_settings(settings)
            self.settings = settings
            records = [rec for rec in data["records"] if isinstance(rec, dict)]
            DatabaseManager.save_records_bulk(records)
            with get_db() as conn:
                conn.execute("DELETE FROM directories")
                conn.execute("DELETE FROM report_templates")
                if data.get("users"):
                    conn.execute("DELETE FROM users WHERE id <> 'builtin-admin'")
            for item in data.get("directories", []):
                if isinstance(item, dict):
                    DatabaseManager.save_directory(item.get("entity_type", ""), item.get("name", ""), item.get("mobile", ""), item.get("details", ""))
            now_restore = dt.datetime.now().isoformat(timespec="seconds")
            with get_db() as conn:
                for item in data.get("report_templates", []):
                    if not isinstance(item, dict) or not clean_value(item.get("name", "")):
                        continue
                    conn.execute(
                        "INSERT OR REPLACE INTO report_templates(id,name,created_at,updated_at,payload) VALUES (?,?,?,?,?)",
                        (clean_value(item.get("id", "")) or str(uuid.uuid4()), clean_value(item.get("name", "")),
                         clean_value(item.get("created_at", now_restore)), clean_value(item.get("updated_at", now_restore)),
                         json.dumps(item.get("payload", {}), ensure_ascii=False)),
                    )
                for item in data.get("users", []):
                    if not isinstance(item, dict):
                        continue
                    uid = clean_value(item.get("id", ""))
                    if not uid:
                        continue
                    conn.execute(
                        """INSERT OR REPLACE INTO users(id, username, display_name, role, permissions, salt, password_hash, active, created_at, updated_at)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (uid, clean_value(item.get("username", "")) or "user", clean_value(item.get("display_name", "")) or "مستخدم",
                         clean_value(item.get("role", "موظف")) or "موظف", json.dumps(item.get("permissions", []), ensure_ascii=False),
                         clean_value(item.get("salt", "")), clean_value(item.get("password_hash", "")), int(bool(item.get("active", True))),
                         clean_value(item.get("created_at", now_restore)), clean_value(item.get("updated_at", now_restore))),
                    )
            restored_attachments = 0
            restored_reports = 0
            if archive is not None:
                restored_attachments = DatabaseManager.restore_attachment_rows(data.get("attachments", []), archive)
                restored_reports = DatabaseManager.restore_report_archive(data.get("reports", []), archive)
                archive.close()
            self.records = DatabaseManager.load_records()
            self.current_report_record_ids = None
            self._form_undo_stack.clear()
            self._apply_filters()
            self._configure_styles()
            self._build_ui()
            self.status_var.set("تمت الاستعادة بنجاح")
            DatabaseManager.log_action("استعادة نسخة احتياطية", details=Path(path).name)
            messagebox.showinfo(
                "تم",
                f"تمت استعادة البيانات بنجاح.\nالمرفقات المستعادة: {restored_attachments}"
                f"\nالتقارير المستعادة: {restored_reports}",
                parent=self.root,
            )
        except Exception as e:
            try:
                if "archive" in locals() and archive is not None:
                    archive.close()
            except Exception:
                pass
            if safety_copy and safety_copy.exists():
                try:
                    shutil.copy2(safety_copy, DB_FILE)
                    self.records = DatabaseManager.load_records()
                except OSError:
                    logging.exception("تعذر التراجع عن الاستعادة الفاشلة")
            logging.exception("فشلت استعادة النسخة الاحتياطية")
            messagebox.showerror("خطأ", str(e), parent=self.root)

    def _database_health_notice(self):
        if not DB_HEALTH_ERROR:
            return
        self.status_var.set("تنبيه: قاعدة البيانات تحتاج إلى فحص")
        messagebox.showerror(
            "سلامة قاعدة البيانات",
            "اكتشف البرنامج مشكلة أثناء فحص قاعدة البيانات ولن يحذف أي بيانات تلقائياً.\n\n"
            f"التفاصيل: {DB_HEALTH_ERROR}\n\n"
            f"المسار: {DB_FILE}\n"
            "استخدم آخر نسخة احتياطية سليمة أو تواصل مع الدعم قبل متابعة إدخال البيانات.",
            parent=self.root,
        )

    def _overdue_notice(self):
        if not self.settings.overdue_warning:
            return
        overdue_count = sum(1 for record in self.records if record_is_overdue(record))
        if overdue_count:
            self.status_var.set(f"يوجد {overdue_count} تسليم متأخر يحتاج إلى متابعة")
            messagebox.showwarning(
                "متابعة التسليم",
                f"يوجد {overdue_count} سجل تجاوز موعد التسليم ولم يُسجل كمُسلّم.\n"
                "يمكن مراجعتها من «تحليلات المبيعات والتسليم».",
                parent=self.root,
            )

    def _dependency_notice(self):
        missing = []
        if Workbook is None:
            missing.append("openpyxl")
        if colors is None:
            missing.append("reportlab")
        if Image is None:
            missing.append("Pillow")
        if arabic_reshaper is None or get_display is None:
            missing.extend(["arabic-reshaper", "python-bidi"])
        if xlrd is None:
            missing.append("xlrd")
        if missing:
            messagebox.showwarning(
                "متطلبات ناقصة",
                "بعض المكتبات غير مثبتة:\n" + ", ".join(dict.fromkeys(missing)) +
                "\n\nشغّل Install_Requirements.ps1 من مجلد المصدر، "
                "أو نفّذ: py -m pip install -r requirements.txt",
                parent=self.root,
            )

    def _session_payload(self) -> dict[str, Any]:
        draft_values = {}
        for key, entry in getattr(self, "entries", {}).items():
            try:
                draft_values[key] = clean_value(entry.get())
            except tk.TclError:
                pass
        filter_field_var = getattr(self, "filter_field_var", None)
        filter_value_var = getattr(self, "filter_value_var", None)
        status_filter_var = getattr(self, "status_filter_var", None)
        filter_date_basis_var = getattr(self, "filter_date_basis_var", None)
        filter_from_var = getattr(self, "filter_from_var", None)
        filter_to_var = getattr(self, "filter_to_var", None)
        return {
            "draft_values": draft_values,
            "recipient": clean_value(self.recipient_var.get()),
            "search": self.search_var.get(),
            "filter_field": filter_field_var.get() if filter_field_var is not None else "الكل",
            "filter_value": filter_value_var.get() if filter_value_var is not None else "",
            "status_filter": status_filter_var.get() if status_filter_var is not None else "كل الحالات",
            "filter_date_basis": (
                filter_date_basis_var.get() if filter_date_basis_var is not None else "التاريخ الأنسب"
            ),
            "filter_from": filter_from_var.get() if filter_from_var is not None else "",
            "filter_to": filter_to_var.get() if filter_to_var is not None else "",
            "report_workspace_active": self.current_report_record_ids is not None,
            "current_user": clean_value(self.current_user.get("username", "")),
            "current_report_record_ids": (
                sorted(self.current_report_record_ids)
                if self.current_report_record_ids is not None else []
            ),
        }

    def _restore_report_workspace(self, payload: dict[str, Any]):
        """يحافظ على بقاء التقرير الجديد فارغًا حتى بعد إغلاق البرنامج وفتحه."""
        if not isinstance(payload, dict) or not payload.get("report_workspace_active", False):
            self.current_report_record_ids = None
            self._sync_report_workspace_ui()
            return
        raw_ids = payload.get("current_report_record_ids", [])
        requested = {str(item) for item in raw_ids} if isinstance(raw_ids, list) else set()
        existing = {str(record.get("id", "")) for record in self.records}
        self.current_report_record_ids = requested & existing
        self._sync_report_workspace_ui()

    def _restore_session_state(self, payload: dict[str, Any]) -> bool:
        if not isinstance(payload, dict):
            return False
        restored = False
        draft = payload.get("draft_values", {})
        if isinstance(draft, dict):
            for key, value in draft.items():
                entry = self.entries.get(key)
                if entry is None or not clean_value(value):
                    continue
                try:
                    if isinstance(entry, ttk.Combobox):
                        entry.set(value)
                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, value)
                    restored = True
                except tk.TclError:
                    pass
        self.recipient_var.set(clean_value(payload.get("recipient", "")))
        self.search_var.set(clean_value(payload.get("search", "")))
        if hasattr(self, "filter_field_var"):
            candidate = clean_value(payload.get("filter_field", "الكل")) or "الكل"
            allowed = ["الكل", *[field.label for field in self.settings.visible_fields]]
            self.filter_field_var.set(candidate if candidate in allowed else "الكل")
        if hasattr(self, "filter_value_var"):
            self.filter_value_var.set(clean_value(payload.get("filter_value", "")))
        if hasattr(self, "status_filter_var"):
            candidate = clean_value(payload.get("status_filter", "كل الحالات")) or "كل الحالات"
            self.status_filter_var.set(candidate if candidate in {"كل الحالات", *STATUS_OPTIONS} else "كل الحالات")
        if hasattr(self, "filter_date_basis_var"):
            candidate = clean_value(payload.get("filter_date_basis", "التاريخ الأنسب")) or "التاريخ الأنسب"
            allowed_basis = {
                "التاريخ الأنسب", "موعد التسليم", "التركيب الفعلي",
                "تاريخ إنشاء السجل", "تاريخ آخر تعديل",
            }
            self.filter_date_basis_var.set(candidate if candidate in allowed_basis else "التاريخ الأنسب")
        if hasattr(self, "filter_from_var"):
            self.filter_from_var.set(clean_value(payload.get("filter_from", "")))
        if hasattr(self, "filter_to_var"):
            self.filter_to_var.set(clean_value(payload.get("filter_to", "")))
        return restored or any([
            clean_value(payload.get("recipient", "")), clean_value(payload.get("search", "")),
            clean_value(payload.get("filter_value", "")), clean_value(payload.get("filter_from", "")),
            clean_value(payload.get("filter_to", "")),
        ])

    def _autosave_session_state(self):
        self._session_after_id = None
        try:
            DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
        finally:
            try:
                self._session_after_id = self.root.after(int(self.settings.autosave_seconds * 1000), self._autosave_session_state)
            except tk.TclError:
                self._session_after_id = None

    def logout(self):
        if not messagebox.askyesno(
            "تسجيل الخروج",
            "سيحفظ البرنامج البيانات والجلسة وينشئ نسخة احتياطية ثم يغلق. هل تريد المتابعة؟",
            parent=self.root,
        ):
            return
        DatabaseManager.log_action("تسجيل خروج", details="إغلاق آمن للجلسة")
        self.close()

    def refresh_from_database(self, _event=None):
        """تحديث آمن بـ F5 يحفظ المكتمل ويحافظ على المسودة غير المكتملة."""
        self._cancel_record_autosave()
        draft_incomplete = self._has_user_form_data() and not self._autosave_current_record(refresh=False)
        try:
            self.records = DatabaseManager.load_records()
            existing_ids = {str(record.get("id", "")) for record in self.records}
            if self.current_report_record_ids is not None:
                self.current_report_record_ids.intersection_update(existing_ids)
            if self.selected_record_id and self.selected_record_id not in existing_ids:
                self.selected_record_id = None
                self.mode_label.configure(text="مسودة جديدة", fg=self.settings.accent_color)
            self._on_filter_field_changed()
            DatabaseManager.save_session_state(self._session_payload(), clean_shutdown=False)
            if draft_incomplete:
                self.status_var.set("تم التحديث بـ F5 مع الاحتفاظ بالمسودة غير المكتملة")
            else:
                self.status_var.set(f"تم التحديث من قاعدة البيانات - {len(self._workspace_records())} سجل ظاهر في مساحة العمل")
        except Exception as exc:
            logging.exception("تعذر تحديث السجلات من قاعدة البيانات")
            messagebox.showerror("التحديث", f"تعذر تحديث البيانات: {exc}", parent=self.root)
        return "break"

    def _handle_delete_shortcut(self, _event=None):
        """يمنع حذف سجل كامل عند ضغط Delete أثناء تحرير نص داخل حقل."""
        if hasattr(self, "tree") and self.root.focus_get() is self.tree:
            self.delete_selected()
            return "break"
        return None

    def _handle_copy_shortcut(self, _event=None):
        if hasattr(self, "tree") and self.root.focus_get() is self.tree:
            return self.copy_selected_records()
        return None

    def _handle_paste_shortcut(self, _event=None):
        if hasattr(self, "tree") and self.root.focus_get() is self.tree:
            return self.paste_records_from_clipboard()
        return None

    def _bind_shortcuts(self):
        self.root.bind("<Control-n>", lambda e: self.start_new_record())
        self.root.bind("<Control-Shift-N>", lambda e: self.start_new_report())
        self.root.bind("<Control-s>", lambda e: self.save_or_update_record())
        self.root.bind("<Control-f>", lambda e: self.search_entry.focus_set() if hasattr(self, "search_entry") else None)
        self.root.bind("<Control-m>", lambda e: self.toggle_sidebar())
        self.root.bind("<Control-e>", lambda e: self._export("xlsx"))
        self.root.bind("<Control-p>", lambda e: self._export("pdf"))
        self.root.bind("<Control-i>", lambda e: self._export("png"))
        self.root.bind("<Control-Shift-C>", lambda e: self._export("csv"))
        self.root.bind("<Control-h>", lambda e: self._export("html"))
        self.root.bind("<Control-o>", lambda e: self.import_from_excel())
        self.root.bind("<Control-c>", self._handle_copy_shortcut)
        self.root.bind("<Control-v>", self._handle_paste_shortcut)
        self.root.bind("<Control-d>", lambda e: self.duplicate_record())
        self.root.bind("<Delete>", self._handle_delete_shortcut)
        self.root.bind("<Escape>", self._undo_form)
        self.root.bind("<F2>", lambda e: self.toggle_form_panel())
        self.root.bind("<F3>", lambda e: self.toggle_filters_panel())
        self.root.bind("<F4>", lambda e: self.toggle_stats_panel())
        self.root.bind("<F9>", lambda e: self.toggle_report_focus_mode())
        self.root.bind("<F11>", lambda e: self.toggle_window_fullscreen())
        self.root.bind("<F5>", self.refresh_from_database)

    def _update_clock(self):
        now = dt.datetime.now()
        days = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
        self.clock_var.set(f"{days[now.weekday()]}\n{now:%Y-%m-%d  |  %H:%M}")
        self.root.after(30000, self._update_clock)

    def close(self):
        self._cancel_record_autosave()
        record_saved = self._autosave_current_record()
        has_unfinished_draft = any(
            clean_value(entry.get()) and not (
                key == "delivery_status" and clean_value(entry.get()) == "جديد"
            )
            for key, entry in getattr(self, "entries", {}).items()
        ) and not record_saved
        if self._session_after_id is not None:
            try:
                self.root.after_cancel(self._session_after_id)
            except tk.TclError:
                pass
            self._session_after_id = None
        try:
            backup_path = BACKUP_DIR / f"backup_automatic_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            self._write_backup(backup_path)
            validate_backup_archive(backup_path)
            automatic = sorted(BACKUP_DIR.glob("backup_automatic_*.zip"), key=lambda item: item.stat().st_mtime, reverse=True)
            for old_backup in automatic[self.settings.backup_keep_count:]:
                old_backup.unlink(missing_ok=True)
            if self.settings.sync_folder:
                sync_dir = Path(self.settings.sync_folder)
                if sync_dir.is_dir():
                    shutil.copy2(backup_path, sync_dir / backup_path.name)
        except:
            logging.exception("تعذر إنشاء النسخة الاحتياطية التلقائية")
        DatabaseManager.save_session_state(
            self._session_payload(), clean_shutdown=not has_unfinished_draft
        )
        self.root.destroy()

# ============================================================================
# نقطة الدخول
# ============================================================================
def run_packaged_self_test() -> int:
    """اختبار داخلي صامت للحزمة التنفيذية دون لمس سجلات المستخدم."""
    try:
        settings = AppSettings()
        settings.print_friendly_reports = True
        settings.report_page_size = "A4"
        settings.report_orientation = "landscape"
        settings.png_dpi = 300
        now = dt.datetime.now().isoformat(timespec="seconds")
        records = [{
            "id": "self-test",
            "created_at": now,
            "updated_at": now,
            "values": {
                "material_estimate": "1", "contract_number": "SELF-TEST-001", "kitchen_number": "1",
                "customer_name": "اختبار داخلي", "mobile": "0500000000",
                "marble_color": "أبيض", "sink_type": "علوي", "marble_edge": "5 سم",
                "showroom": "الاختبار", "salesperson": "محلي", "delivery_status": "جديد",
                "scheduled_delivery_date": dt.date.today().isoformat(), "actual_delivery_date": "",
            },
        }, {
            "id": "self-test-kitchen-2",
            "created_at": now,
            "updated_at": now,
            "values": {
                "material_estimate": "1", "contract_number": "SELF-TEST-001", "kitchen_number": "2",
                "customer_name": "اختبار داخلي", "mobile": "0500000000",
                "marble_color": "أبيض", "sink_type": "علوي", "marble_edge": "5 سم",
                "showroom": "الاختبار", "salesperson": "محلي", "delivery_status": "جديد",
                "scheduled_delivery_date": dt.date.today().isoformat(), "actual_delivery_date": "",
            },
        }]
        with tempfile.TemporaryDirectory(prefix="marble_delivery_self_test_") as temp_name:
            temp_dir = Path(temp_name)
            exporter = ReportExporter(settings)
            outputs = {
                "pdf": temp_dir / "test.pdf", "xlsx": temp_dir / "test.xlsx",
                "png": temp_dir / "test.png", "jpg": temp_dir / "test.jpg", "csv": temp_dir / "test.csv",
                "html": temp_dir / "test.html",
            }
            for file_type, output in outputs.items():
                method = getattr(exporter, EXPORT_METHOD_NAMES[file_type])
                method(records, "اختبار", output, settings.visible_fields)
                if not output.is_file() or output.stat().st_size < 100:
                    raise RuntimeError(f"فشل اختبار {file_type.upper()}")
            if Image is not None:
                with Image.open(outputs["png"]) as png_test:
                    expected_width = int(round(297 / 25.4 * settings.png_dpi))
                    expected_height = int(round(210 / 25.4 * settings.png_dpi))
                    if abs(png_test.width - expected_width) > 2 or abs(png_test.height - expected_height) > 2:
                        raise RuntimeError("فشل اختبار مقاس ودقة PNG الحقيقيين")
            preview = SpreadsheetImportReader(settings).read(outputs["xlsx"])
            if len(preview.data_rows) != 2:
                raise RuntimeError("فشل اختبار إعادة استيراد Excel")
            metrics = report_record_metrics(records)
            if metrics["contracts"] != 1 or metrics["records"] != 2 or next_kitchen_number(records, "SELF-TEST-001") != "3":
                raise RuntimeError("فشل اختبار تعدد المطابخ للعقد الواحد")
        return 0
    except Exception:
        logging.exception("فشل الاختبار الذاتي للحزمة")
        return 2


def main():
    if "--self-test" in sys.argv:
        return run_packaged_self_test()
    root = tk.Tk()
    settings = DatabaseManager.load_settings()
    current_user = next((u for u in DatabaseManager.list_users(active_only=True) if u.get("id") == "builtin-admin"), dict(CURRENT_USER))
    if settings.require_login:
        root.withdraw()
        login = LoginDialog(root)
        root.wait_window(login)
        if not login.result:
            root.destroy(); return 0
        current_user = login.result
        root.deiconify()
    app = MarbleDeliveryApp(root, current_user=current_user)
    root.mainloop()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
