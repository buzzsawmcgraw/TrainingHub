# TrainingHub

PIP-Boy style personnel UI for SharePoint **Modern Script Editor** (PnP).

## Recommended deploy (small paste + Site Assets)

Keeps the Script Editor web part small (~7 KB) so the SharePoint page stays editable as the hub grows.

### Step 1 — Get the files from GitHub

Download these three files from the repo (or run `build-sharepoint-deploy.ps1` locally after edits):

| File | Purpose |
|------|---------|
| `training-hub-styles.txt` | All CSS (~20 KB) |
| `training-hub-script.txt` | All JavaScript (~56 KB) |
| `sharepoint-script-editor-loader.html` | Small paste block (~7 KB) |

`.txt` extensions are used because some networks block `.js` downloads from Site Assets.

### Step 2 — Upload to Site Assets

1. Open your SharePoint site (default: `https://usaf.dps.mil/sites/88thSFS`).
2. Go to **Site contents** → **Site Assets** (or your site’s document library used for static files).
3. Upload **both** files to the **root** of that library:
   - `training-hub-styles.txt`
   - `training-hub-script.txt`
4. If a file already exists, **replace** it when you update the hub.

After upload, the files should be reachable at:

- `/sites/88thSFS/SiteAssets/training-hub-styles.txt`
- `/sites/88thSFS/SiteAssets/training-hub-script.txt`

If your library path is different, change `SITE_ASSETS_BASE` inside `sharepoint-script-editor-loader.html` before pasting.

### Step 3 — Paste the loader into Script Editor

1. Open `sharepoint-script-editor-loader.html` in a text editor.
2. **Select all** (Ctrl+A) and **copy**.
3. Edit your SharePoint page → open the **Modern Script Editor** web part.
4. **Delete** any old hub code in the web part (if replacing an earlier version).
5. **Paste** the loader.
6. **Save** the web part, then **publish** the page.

### Step 4 — Test

1. Open the **published** page (not edit mode).
2. You should briefly see “Loading Training Hub…”, then the password screen.
3. Enter the password (default `Training2026`) and confirm the roster loads.

If you see a red error about uploading files, the `.txt` files are missing or `SITE_ASSETS_BASE` is wrong.

### Updating the hub later

1. Edit `sharepoint-modern-script-editor-ui.html` (or pull the latest from GitHub).
2. Run `build-sharepoint-deploy.ps1`.
3. Re-upload `training-hub-styles.txt` and `training-hub-script.txt` to Site Assets.
4. You usually **do not** need to re-paste the loader unless the HTML shell changed.

---

## Fallback deploy (all-in-one paste)

Use only if Site Assets upload is not possible.

1. Open **`sharepoint-script-editor-paste.html`**.
2. Select all → paste into Modern Script Editor.
3. Save and publish.

This is ~78 KB in one block. It works today but may cause page edit issues as the hub grows.

---

## Configuration

Edit constants in **`sharepoint-modern-script-editor-ui.html`**, then run `build-sharepoint-deploy.ps1`:

- **`HUB_ACCESS_PASSWORD`** — page password (default `Training2026`). Set `""` to disable. Unlock lasts for the browser tab (`sessionStorage`).
- **`PERSONNEL_SITE_ROOT_URL`** — site root for REST calls (default `https://usaf.dps.mil/sites/88thSFS`).
- **`LIST_PERSONNEL`** — list title (default `Personnel`).

**Record review:** click **Record** on a roster row to open the full personnel view. **Back to roster** returns without reloading the page.

Password and list settings live in the **script** file (`training-hub-script.txt` after build), not in the loader.

## Local preview

Open `sharepoint-modern-script-editor-ui.html` in a browser.

## Lists required

- **Personnel** list on the same site (`LIST_PERSONNEL`, default title `Personnel`)
