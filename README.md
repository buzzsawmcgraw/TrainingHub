# TrainingHub

Paste-only UI for SharePoint **Modern Script Editor** (PnP).

## Deploy

1. Open **`sharepoint-script-editor-paste.html`** (or run `build-sharepoint-deploy.ps1` after editing the source).
2. Select **all**, paste into Modern Script Editor.
3. Save the web part and **publish** the page.

One self-contained block: HTML, CSS, and JavaScript together. No Site Assets upload required.

**Personnel records:** **Record** opens the full row; **Edit record** then **Save changes** updates SharePoint via REST.

After editing the hub, run `build-sharepoint-deploy.ps1` to regenerate `sharepoint-script-editor-paste.html`.

## Page edit mode

The hub **pauses while the SharePoint page is in edit mode** (shows a placeholder, no REST calls). That keeps the page editor and Script Editor usable. View the **published** page to load the roster.

If the page says **Cannot edit at this time**, restore an older version from **Version history** (before the bad save), then paste the latest `sharepoint-script-editor-paste.html`.

## Local preview

Open `sharepoint-modern-script-editor-ui.html` in a browser (demo data on localhost).

## Lists required

- **Personnel** list on the same site (`LIST_PERSONNEL`, default title `Personnel`)

Set `PERSONNEL_SITE_ROOT_URL` in the script if needed (default `https://usaf.dps.mil/sites/88thSFS`).
