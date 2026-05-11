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

- **Personnel** list
- **Certifiers** list (with a column internal name `certifier`)

The UI uses `/_api` calls and expects to run on the SharePoint site so your signed-in session provides auth cookies.

