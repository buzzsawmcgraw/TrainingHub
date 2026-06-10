# TrainingHub

Paste-only UI for SharePoint **Modern Script Editor** (PnP).

## Deploy (recommended — small paste)

Large all-in-one pastes (~130 KB) can **corrupt the SharePoint page** after save: the hub may work once, then the page shows “Cannot edit at this time” and stays locked. Use the **small loader** instead.

### Step 1 — Upload Site Assets files

On the **same site** as your page (e.g. `https://usaf.dps.mil/sites/88thSFS`), upload to **Site Assets**:

| File | Purpose |
|------|---------|
| `training-hub-styles.txt` | Full hub CSS |
| `training-hub-script.txt` | Full hub JavaScript |

Download both from this repo (or run `build-sharepoint-deploy.ps1` after editing the source).

Default URLs the loader expects:

- `/sites/88thSFS/SiteAssets/training-hub-styles.txt`
- `/sites/88thSFS/SiteAssets/training-hub-script.txt`

If your site path differs, edit `SITE_ASSETS_BASE` in `sharepoint-script-editor-paste.html` before pasting.

### Step 2 — Paste into Script Editor

1. Open **`sharepoint-script-editor-paste.html`** (~5 KB).
2. Select all, paste into Modern Script Editor.
3. Save the web part, **publish** the page.
4. View the **published** page to load the roster (not page edit mode).

### Step 3 — Regenerate after code changes

Edit `sharepoint-modern-script-editor-ui.html`, then run:

```powershell
.\build-sharepoint-deploy.ps1
```

Re-upload the two `.txt` files to Site Assets and re-publish the page. You usually **do not** need to change the Script Editor paste unless the HTML shell changed.

## If the page is already locked

Symptoms: worked once, then **Cannot edit at this time**, comment panel on reload.

1. **Page version history** — Site Pages library → your page → Version history → restore a version from **before** the large paste.
2. Or delete the broken page and recreate it with the **small loader** paste only.
3. Remove any old ~130 KB inline script from Script Editor; do not re-paste the all-in-one file.

## Features

- **Personnel roster** from SharePoint **Personnel** list (REST)
- **Record** → view full row; **Edit record** → **Save changes**
- **Recall roster** PDF export
- Hub **pauses during page edit** so Script Editor stays usable

## Local preview

Open `sharepoint-modern-script-editor-ui.html` in a browser (full HTML document with demo data on localhost).

## Lists required

- **Personnel** (title must match `LIST_PERSONNEL` in the script, default `Personnel`)

Hub uses `/_api` on the signed-in SharePoint session. Set `PERSONNEL_SITE_ROOT_URL` in the script if the list is on a known site root.
