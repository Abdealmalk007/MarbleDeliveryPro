# MarbleDeliveryPro

**Version:** 10.0.5 — Public clean release  
**Developer:** Abu Omar / ProTech  
**Platform:** Windows x64 (Python 3.11+ to run source)

برنامج عربي لإدارة تسليم الرخام ومتابعة الطلبات والعقود، والمرفقات، والاستيراد من Excel/CSV، وإصدار التقارير PDF / Excel / CSV / HTML / PNG / JPG.

## Privacy and first run

- No customer records, phone numbers, completed transactions, private attachments, historical reports, backups, or database files are included in this repository or release.
- A new database is initialized in `%LOCALAPPDATA%\MarbleDeliveryPro` for each Windows user, not in the installation folder.
- Existing local data are **not deleted** by installing or upgrading. To test a genuinely blank first run, use a new Windows profile or set `MARBLE_DELIVERY_DATA_DIR` to an empty folder.
- Company name is blank by default. Set the company name, business fields, users and reports in Settings after launch.
- Login enforcement is initially off. Before enabling it, configure a password for an administrator in user management. Passwordless accounts cannot authenticate.
- Integrations are opt-in. Do not share API tokens or user databases on GitHub.

## Run from source

Install Python 3.11+ and run PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\تشغيل_المصدر.ps1
```

Alternatively, run the supplied Windows Setup or extract the Portable ZIP.

## Build and verify

With Python and Inno Setup 6 installed:

```powershell
.\Build_All.ps1
```

This runs the unit/regression suite and source/EXE self-tests, then produces `dist/Setup_MarbleDeliveryPro_v10.0.5.exe`, the portable ZIP, and a SHA256 manifest. A code-signing certificate is required to publish a signed binary; unsigned builds may display Windows SmartScreen warnings.

Before publishing the source archive, run:

```powershell
py -3 tools\verify_public_release.py
```

## Publish to GitHub

After building and reviewing the release assets, install GitHub CLI and run `gh auth login`. Then run `powershell -ExecutionPolicy Bypass -File .\PUBLISH_GITHUB.ps1`. The script requires you to type `PUBLISH` before it changes GitHub.

## License

MIT — see `LICENSE`. A contribution must not contain user data, employee names, API tokens, or licensed third-party assets.

**More details:** [README_AR.md](README_AR.md) | [SECURITY.md](SECURITY.md) | [RELEASE_NOTES_AR.md](RELEASE_NOTES_AR.md)
