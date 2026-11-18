from pathlib import Path
root = Path(__file__).resolve().parents[1]
files = [root / "Build_All.ps1", *root.glob("*.iss")]
for file in files:
    data = file.read_text(encoding="utf-8-sig")
    file.write_text(data, encoding="utf-8-sig")
    print("UTF8 BOM fixed:", file.name)
