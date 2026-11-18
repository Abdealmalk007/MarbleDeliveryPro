# -*- coding: utf-8 -*-
"""Stage only allowlisted clean source into an isolated GitHub checkout."""
from __future__ import annotations
import shutil
from pathlib import Path

source=Path(__file__).resolve().parents[1]
repo=Path(r"E:\ProTech\MarbleDeliveryPro_GitHub_v10.0.7")
assert (repo/".git").is_dir()
allowed_ext={".py",".iss",".ps1",".cmd",".md",".txt",".png",".ico",".svg"}
root_names={
    "Build_All.ps1","Build_Setup.cmd","Install_Build_Tools.ps1",
    "Install_Requirements.ps1","PUBLISH_GITHUB.ps1",
    "README.md","README_AR.md","README_INSTALL_AR.txt",
    "RELEASE_NOTES_AR.md","requirements.txt","requirements-build.txt",
    "SECURITY.md","PROTECH_FUTURE_RIGHTS_POLICY.md","app_icon.ico",
}
root_files=[item for item in source.iterdir() if item.is_file()
            and (item.name in root_names or item.suffix in {".py",".iss"})]
for file in root_files:
    destination=repo/file.name
    raw=None
    if file.suffix in {".md",".txt",".ps1",".cmd",".iss"}:
        raw=file.read_text(encoding="utf-8-sig")
        if file.name in {"PUBLISH_GITHUB.ps1","README.md","README_AR.md"}:
            raw=raw.replace("10.0.5","10.0.7").replace("10.0.6","10.0.7")
    if raw is not None:
        destination.write_text(raw,encoding="utf-8-sig" if file.suffix==".ps1" else "utf-8")
    else:
        shutil.copy2(file,destination)

for dirname in ("assets","tests","docs","tools"):
    srcdir=source/dirname
    for file in srcdir.rglob("*"):
        if not file.is_file() or file.suffix not in allowed_ext:
            continue
        if "__pycache__" in file.parts:
            continue
        if any(word in file.name.lower() for word in (".bak",".log","baseline","backup","private")):
            continue
        if dirname=="tools" and file.name in {"upgrade_v1006.py","package_local_release_v1006.py"}:
            continue
        relative=file.relative_to(source)
        destination=repo/relative
        destination.parent.mkdir(parents=True,exist_ok=True)
        shutil.copy2(file,destination)

# Historical license and legal decisions are preserved as stored in the repo.
assert (repo/"LICENSE").is_file()
for extension in (".db",".sqlite",".sqlite3",".exe",".zip",".log",".bak"):
    assert not any(file.suffix.lower()==extension for file in (repo/"tests").rglob("*"))
print("STAGED_CLEAN_SOURCE")
print("ROOT_FILES",len(root_files))
print("VERSION",next(repo.glob("*.py")).name)
