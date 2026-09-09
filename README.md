# Vision Issue Tracker

Windows desktop issue tracker for mass-production vision inspection instruments.

The current source supports both:

- local SQLite mode for one shared PC
- online Google Sheets + Apps Script mode with a local cache for multiple PCs

## Safety

Do not commit real deployment secrets. Keep these local only:

- `config.json`
- real Google Sheet IDs
- deployed Apps Script `/exec` URLs
- API tokens
- generated `.exe` / `.zip` release files
- local `data/` databases

Use `config.example.json` and `apps_script/Code.gs` as sanitized templates.

## Main Features

- Issue Board with Action Required and Monitoring columns
- Create/Edit issue workflow with multi-line and multi-vision selection
- Search / Report with Excel export
- Version History dashboard and SW/Algo version descriptions
- Korean/English UI selector
- Online sync status with local cache fallback

## Online Setup

See `docs/ONLINE_SETUP.md` for the Google Sheets / Apps Script setup flow.

## Run From Source

```powershell
python app.py
```

Install dependencies first if needed:

```powershell
pip install -r requirements.txt
```

## Build EXE

```powershell
python -m PyInstaller --noconfirm VisionIssueTracker.spec
```

When publishing a new exe, commit and push the source changes, but keep generated release files and real deployment credentials out of GitHub.