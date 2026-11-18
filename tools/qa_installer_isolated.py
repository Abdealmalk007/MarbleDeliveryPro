# -*- coding: utf-8 -*-
"""Safely test a duplicate installer with a distinct AppId and no user shortcuts."""
from __future__ import annotations
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]
ISS = next(ROOT.glob("*.iss"))
QA_ROOT = ROOT / ".build" / "qa_installer"
QA_ROOT.mkdir(parents=True, exist_ok=True)
qa_script = ROOT / "MarbleDeliveryPro_v10.0.7_qa.iss"
script = ISS.read_text(encoding="utf-8-sig")
def replace_one(regex: str, value: str):
    global script
    script, count = re.subn(regex, lambda _match: value, script, count=1, flags=re.MULTILINE)
    if count != 1:
        raise RuntimeError("Installer QA script missing: " + regex)
replace_one(r"^AppId=.*$", "AppId={{CD77E66A-59CB-4EDA-9B0A-A7BC4B6070A6}")
replace_one(r"^DefaultDirName=.*$", r"DefaultDirName={localappdata}\Programs\MarbleDeliveryPro-ISOLATED-QA")
replace_one(r"^DefaultGroupName=.*$", "DefaultGroupName=MarbleDeliveryPro ISOLATED QA")
replace_one(r"^OutputDir=.*$", "OutputDir=.build\\qa_installer")
replace_one(r"^OutputBaseFilename=.*$", "OutputBaseFilename=Setup_MarbleDeliveryPro_v10.0.7_ISOLATED_QA")
replace_one(r"^CloseApplications=.*$", "CloseApplications=no")
for section in ("Icons", "Tasks", "Run", "UninstallRun"):
    script = re.sub(
        r"(?ms)^\[" + re.escape(section) + r"\]\r?\n.*?(?=^\[|\Z)",
        f"[{section}]\n", script
    )
qa_script.write_text(script, encoding="utf-8-sig")

iscc = Path(os.environ["LOCALAPPDATA"]) / "Programs" / "Inno Setup 6" / "ISCC.exe"
if not iscc.is_file():
    raise RuntimeError(f"Missing Inno Setup compiler: {iscc}")
compile_log = QA_ROOT / "compile.log"
with compile_log.open("w", encoding="utf-8") as out:
    completed = subprocess.run([str(iscc), str(qa_script)], cwd=ROOT,
                               stdout=out, stderr=subprocess.STDOUT, timeout=180)
if completed.returncode:
    raise RuntimeError(f"QA Inno compile failure: {completed.returncode}")

installer = QA_ROOT / "Setup_MarbleDeliveryPro_v10.0.7_ISOLATED_QA.exe"
assert installer.is_file() and installer.stat().st_size > 1_000_000
install_to = QA_ROOT / "installed"
if install_to.exists():
    raise RuntimeError("QA install path already exists; not overwriting")
install_log = QA_ROOT / "install.log"
command = [
    str(installer), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART",
    "/SP-", f"/DIR={install_to}", f"/LOG={install_log}",
]
print("Installing isolated QA instance", flush=True)
result = subprocess.run(command, cwd=ROOT, timeout=180)
print("INSTALL_EXIT=", result.returncode, flush=True)
if result.returncode:
    raise RuntimeError(f"QA install failed: {result.returncode}")

exe = install_to / "MarbleDeliveryPro.exe"
if not exe.is_file():
    raise RuntimeError("QA app missing after installation")
with tempfile.TemporaryDirectory(prefix="marble-qa-v1006-") as test_dir:
    env = os.environ.copy()
    env["MARBLE_DELIVERY_DATA_DIR"] = test_dir
    env["PYTHONUTF8"] = "1"
    print("Testing installed application with isolated data", flush=True)
    result = subprocess.run([str(exe), "--self-test"], env=env,
                            timeout=120, cwd=install_to)
    print("INSTALLED_EXE_SELFTEST_EXIT=", result.returncode, flush=True)
    if result.returncode:
        raise RuntimeError(f"Installed EXE self test failed: {result.returncode}")
uninstaller = install_to / "unins000.exe"
if not uninstaller.is_file():
    raise RuntimeError("QA uninstaller missing")
print("Uninstalling isolated QA instance", flush=True)
result = subprocess.run([str(uninstaller), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],
                        timeout=120, cwd=ROOT)
print("UNINSTALL_EXIT=", result.returncode, flush=True)
if result.returncode:
    raise RuntimeError(f"QA uninstall failed: {result.returncode}")
for _ in range(15):
    if not exe.exists():
        break
    time.sleep(0.4)
if exe.exists():
    raise RuntimeError("QA app remained after uninstall")
qa_script.unlink(missing_ok=True)
print("QA_INSTALL_UNINSTALL_PASS: separate AppId, separate path, separate test DB", flush=True)
