"""Static + isolated first-run privacy checks; never accesses real user data."""
from __future__ import annotations
import importlib.util
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKIP = {".git", ".build", ".build-venv", "build", "dist", ".venv", "__pycache__", ".pytest_cache"}
PRIVATE_DIRS = {"attachments", "backups", "previews", "exports", "user_data", "data"}
BLOCK_EXT = {".db", ".sqlite", ".sqlite3", ".bak", ".backup", ".wal", ".shm", ".log", ".exe", ".zip", ".csv", ".xlsx", ".xls"}
SOURCE = next(ROOT.glob("*.py"))
checked = 0
for path in ROOT.rglob("*"):
    rel = path.relative_to(ROOT)
    if any(part.lower() in SKIP for part in rel.parts):
        continue
    if any(part.lower() in PRIVATE_DIRS for part in rel.parts):
        raise SystemExit(f"FAIL: private folder: {rel}")
    if path.is_file():
        checked += 1
        if path.suffix.lower() in BLOCK_EXT or path.name.lower().endswith(("-wal", "-shm")):
            raise SystemExit(f"FAIL: generated/private file in source tree: {rel}")
text = SOURCE.read_text(encoding="utf-8-sig")
if "يونيفرسال توب" in text or 'company_name: str = ""' not in text:
    raise SystemExit("FAIL: private/default company name")
if 'APP_VERSION = "10.0.5"' not in text:
    raise SystemExit("FAIL: version mismatch")
with tempfile.TemporaryDirectory(prefix="marble-public-first-run-") as folder:
    os.environ["MARBLE_DELIVERY_DATA_DIR"] = folder
    spec = importlib.util.spec_from_file_location("marble_public_first_run", SOURCE)
    app = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = app
    spec.loader.exec_module(app)
    if app.DB_HEALTH_ERROR:
        raise SystemExit(f"FAIL: DB unhealthy: {app.DB_HEALTH_ERROR}")
    if app.AppSettings().company_name:
        raise SystemExit("FAIL: company name populated")
    with app.get_db() as conn:
        for table in ("records", "attachments", "audit_log", "report_archive", "integration_outbox", "deleted_records", "directories", "report_templates", "session_state", "settings"):
            if conn.execute(f"SELECT count(*) FROM {table}").fetchone()[0]:
                raise SystemExit(f"FAIL: {table} has private data")
        users = conn.execute("SELECT username, salt, password_hash FROM users").fetchall()
        if users != [("admin", "", "")]:
            raise SystemExit(f"FAIL: unexpected accounts: {users!r}")
    if app.DatabaseManager.authenticate("admin", "") or app.password_matches("", "", ""):
        raise SystemExit("FAIL: passwordless login")
    # The app installs a logging.FileHandler; release it before deleting temp DB files.
    import logging
    logging.shutdown()
print(f"PASS: files={checked}; database clean; company blank; passwordless authentication denied")
