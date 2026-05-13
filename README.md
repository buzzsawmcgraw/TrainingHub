# TrainingHub

This repo contains a **paste-only UI** intended for use with a SharePoint **Modern Script Editor** web part.

It is based on requirements gathered from prior agent conversations:

- Paste-only SharePoint UI constraints: `3bfc4a27-4b7a-4f2c-a443-7b178d95bb0a`
- List-backed CRUD pattern over SharePoint REST: `83d6288b-4e99-48ca-8607-724e3b411498`
- Personnel / Phase 1 / Certificates / Certifiers scope: `33d5f28e-e092-49e4-a9ea-8a650ef6ac95` and `8dc4d435-3539-464b-aaf4-7a7b53a442bb`

## What to paste into SharePoint

Open `sharepoint-modern-script-editor-ui.html` and paste the contents into your Modern Script Editor web part.

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

