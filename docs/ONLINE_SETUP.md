# Vision Issue Tracker Online Setup

Google Sheet DB:

```text
PASTE_GOOGLE_SHEET_ID_HERE
```

## Files Created

- `C:\KWAK\6. Vision History Online\apps_script\Code.gs`
- `C:\KWAK\6. Vision History Online\config.example.json`
- `C:\KWAK\4. ESHG\release_online\config.json`

## Apps Script Deployment Steps

1. Open the Google Sheet.
2. Go to `Extensions > Apps Script`.
3. Delete the default code and paste the full contents of:

```text
C:\KWAK\6. Vision History Online\apps_script\Code.gs
```

4. In `Code.gs`, replace:

```javascript
const API_TOKEN = 'PASTE_SHARED_API_TOKEN_HERE';
```

with the `api_token` value from:

```text
C:\KWAK\4. ESHG\release_online\config.json
```

5. Click `Run` on any function or deploy and approve permissions.
6. Click `Deploy > New deployment`.
7. Select type: `Web app`.
8. Suggested settings:
   - Execute as: `Me`
   - Who has access: `Anyone with the link` or your company domain if available
9. Copy the deployed `/exec` URL.
10. Paste that URL into `apps_script_url` in:

```text
C:\KWAK\4. ESHG\release_online\config.json
```

## Important

Do not commit `release_online/config.json` to public GitHub. It contains the shared API token.

The source code should only keep `config.example.json` with placeholders.

## Online 1.1 Performance Notes

- The EXE now keeps a local per-user cache at `%LOCALAPPDATA%\VisionIssueTracker\online_cache.db` so the window can open from cached data before the network refresh finishes, even when the EXE is run from a USB drive or protected folder.
- Apps Script now supports `bootstrap`, `changesSince`, `searchIssues`, soft deletes, and `updated_at` / `deleted_at` metadata.
- To get the optimized sync path, paste the latest `apps_script/Code.gs` into Apps Script and deploy a new Web App version. If the old Apps Script code is still deployed, the EXE falls back to the slower legacy full refresh.
- Keep `config.json` next to the EXE. Do not commit `config.json` or `apps_script/Code.gs` to GitHub.
