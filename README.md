# TrainingHub

This repo contains a **paste-only UI** intended for use with a SharePoint **Modern Script Editor** web part.

## What to paste into SharePoint

1. Open **`sharepoint-script-editor-paste.html`** on GitHub (or locally).
2. Select **all** of it and paste into your Modern Script Editor web part.
3. Save and publish the page.

That file is **one self-contained block** — HTML, CSS, and JavaScript together. You do **not** need to upload anything to Site Assets.

**Personnel records:** Open a row with **Record**, then **Edit record** to change fields (status, phone, address, etc.) and **Save changes**. Updates go to the SharePoint Personnel list via REST.

**Alternative:** You can paste from `sharepoint-modern-script-editor-ui.html` instead (same content, wrapped in a full HTML document for local browser preview).

After editing the hub, run `build-sharepoint-deploy.ps1` to regenerate `sharepoint-script-editor-paste.html`.

### Size note

The paste file is about **110 KB**. Some SharePoint tenants truncate very large Script Editor pastes, which can blank the page or lock it on save. If that happens, use the optional split files (`training-hub-styles.txt` + `training-hub-script.txt` in Site Assets) — ask for help switching back to that layout.

## SharePoint lists required

Minimum:

- **Personnel** list (title must be `Personnel` on the same site as the page)
- **Certifiers** list (title `Certifiers`; names read from field `Certifier`, then `Title` if empty)

List titles and the certifier field key are set in the snippet source (`LIST_PERSONNEL`, `LIST_CERTIFIERS`, `CERTIFIERS_NAME_FIELD`). Change those constants if your site uses different names.

For **Personnel**, `PERSONNEL_IGNORE_SP_TITLE` (default `true`) means the Hub does not send SharePoint’s built-in **Title**; it uses **LastName**, **FirstName**, **MiddleInitial** instead. Set it to `false` if **Save** on a new row fails because the list requires Title (then Title is filled from last/first name or DoD ID).

**Hub Save** writes these fields (REST internal keys; override with `COLUMN_MAP` if yours differ): **RecordDate** (default internal `Record_x0020_Date` for a “Record Date” column), **LastName**, **FirstName**, **MiddleInitial**, **DoDID**, **Address**, **WorkPhone**, **CellPhone**, **Status**, **OfficeSymbol**, **Squadron**, **Rank**, **Notes**. System columns (**Modified**, **Created**, **Created by**, **Modified by**) are not sent.

**Choice dropdowns:** Edit `STATUS_OPTIONS`, `OFFICE_SYMBOLS`, `RANKS`, and `SQUADRONS` in the snippet so every option string **exactly matches** the choices configured on the SharePoint columns.

**Existing rows:** Put the web part on a page in the **same site** as the list. On load (and Refresh), the Hub calls SharePoint REST and lists every item in **Personnel** (up to 5000). If the roster is empty, the list title usually does not match `LIST_PERSONNEL`, the list is on another site, or the browser user lacks **Read** on that list.

The UI uses `/_api` calls and expects to run on the SharePoint site so your signed-in session provides auth cookies.
