(function () {
        // ============ EDIT THESE (URLs remove ambiguity) ============
        /**
         * 88thSFS deployment reference:
         * - Site page: https://usaf.dps.mil/sites/88thSFS/SitePages/S3T.aspx
         * - List view: https://usaf.dps.mil/sites/88thSFS/Lists/Personnel.AllItems.aspx
         *
         * (A) Site root — no trailing slash. REST is always: SITE_ROOT + "/_api/web/..."
         */
        const PERSONNEL_SITE_ROOT_URL = "https://usaf.dps.mil/sites/88thSFS";

        /**
         * (B) Optional fallback: paste any URL under the site; derives …/sites/88thSFS. Ignored when (A) is set.
         * Example: https://usaf.dps.mil/sites/88thSFS/Lists/Personnel.AllItems.aspx
         */
        const PERSONNEL_SHAREPOINT_URL = "";

        /** @deprecated Same effect as PERSONNEL_SITE_ROOT_URL; kept for older copies of this file. */
        const PERSONNEL_REST_WEB_URL = "";

        /** Must match Site contents list title (URL …/Lists/Personnel… usually means title "Personnel"). */
        const LIST_PERSONNEL = "Personnel";
        /** Optional: List GUID from List settings → use if getbytitle fails. */
        const LIST_PERSONNEL_GUID = "";

        /**
         * Circular profile image to the left of the squadron header (default: **S3T Files** / **Fallout Guy.png**).
         * `HUB_PROFILE_IMAGE_URL` — full URL override. `HUB_LOGO_URL` — legacy override if profile URL is empty.
         * Otherwise built from `PERSONNEL_SITE_ROOT_URL` + `HUB_PROFILE_SITE_RELATIVE_PATH`.
         */
        const HUB_PROFILE_IMAGE_URL = "";
        /** Path from site root (spaces as %20). Same folder convention as the banner. */
        const HUB_PROFILE_SITE_RELATIVE_PATH = "S3T%20Files/Fallout%20Guy.png";
        const HUB_PROFILE_ALT = "Fallout guy";

        /**
         * Optional square/rect masthead image (legacy). If set and `HUB_PROFILE_IMAGE_URL` is empty, used as profile
         * source before the Fallout Guy default path.
         */
        const HUB_LOGO_URL = "";
        /** Alt text fallback when the profile image is shown. */
        const HUB_LOGO_ALT = "88th Security Forces Squadron";

        /**
         * Full-width image across the top of the web part (default: **S3T Files** / **banner.png** on this site).
         * Paste a full file URL here to override. Leave "" to build from `PERSONNEL_SITE_ROOT_URL` +
         * `HUB_BANNER_SITE_RELATIVE_PATH`. If the image does not load, open the file in SharePoint, copy the browser
         * address, and set `HUB_TOP_BANNER_URL` to that value (or fix `HUB_BANNER_SITE_RELATIVE_PATH` — e.g. add
         * `Shared%20Documents/` if **S3T Files** is a folder inside Documents, not the library root).
         */
        const HUB_TOP_BANNER_URL = "";
        /** Path from site root to `banner.png` (spaces as %20). */
        const HUB_BANNER_SITE_RELATIVE_PATH = "S3T%20Files/banner.png";
        const HUB_TOP_BANNER_ALT = "88th Security Forces Squadron banner";

        /**
         * Roster column `key` values that render as dropdowns. Options are loaded from SharePoint field **Choices**
         * (same internal name as used for save — see `resolveWriteKey` / first data row).
         */
        const DROPDOWN_COLUMN_KEYS = ["Status", "OfficeSymbol", "Squadron", "Rank"];

        /** Send SharePoint **Title** on create (built from Last + First, else DoD ID, else "New personnel"). */
        const SET_TITLE_ON_CREATE = true;

        /**
         * OData `$orderby` for list items. Rows are always sorted again client-side by LastName, then FirstName, so the
         * roster follows A–Z by name even if this clause is rejected (400).
         */
        const PERSONNEL_ITEMS_ORDERBY = "LastName asc,FirstName asc";

        /**
         * When false, the SharePoint list item `Id` is never shown as a roster column (covers stale Script Editor copies
         * that still list `Id` / "Personnel ID" in ROSTER_COLUMNS). Delete still uses `item.Id`. Set true to allow an Id column.
         */
        const SHOW_LIST_ITEM_ID_IN_ROSTER = false;

        /**
         * **Record** button opens the full personnel view via `?personId=` on this page (or PERSONNEL_DETAIL_PAGE_URL).
         */
        const PERSONNEL_DETAIL_QUERY_PARAM = "personId";
        /** Optional SharePoint page URL for person detail (e.g. same S3T.aspx page). Empty = current page. */
        const PERSONNEL_DETAIL_PAGE_URL = "";
        /** When roster row count exceeds this, the table body scrolls instead of growing the page. */
        const ROSTER_SCROLL_AFTER_ROWS = 10;

        /** Recall roster PDF — header text and download file name. */
        const RECALL_ROSTER_ORG_TITLE = "88th Security Forces Squadron";
        const RECALL_ROSTER_DOC_TITLE = "Recall Roster";
        const RECALL_ROSTER_FILENAME = "88SFS-Recall-Roster.pdf";

        /**
         * Local demo when not on SharePoint. `"auto"` = on for localhost / 127.0.0.1 / file://; `true` = always; `false` = never.
         * Demo uses in-memory sample rows — roster, detail pages, add, and delete all work without REST.
         */
        const LOCAL_DEMO_MODE = "auto";

        /** Choice values for dropdown columns in local demo (Status, Rank, etc.). */
        const DEMO_FIELD_CHOICES = {
          Status: ["Active", "Pending", "Inactive", "TDY"],
          OfficeSymbol: ["S3T", "SFOPS", "CP", "CC"],
          Squadron: ["88 SFS", "88 OSS", "88 MDG"],
          Rank: ["AB", "Amn", "A1C", "SrA", "SSgt", "TSgt", "MSgt", "SMSgt", "CMSgt", "2d Lt", "1st Lt", "Capt", "Maj"],
        };

        /**
         * Roster columns — fixed order for 88thSFS Personnel. List item Id is hidden unless SHOW_LIST_ITEM_ID_IN_ROSTER is true
         * and you include `{ key: "Id", label: "…" }` here.
         * If a column is always blank, open Diagnostics and change that row’s `key` / add `altKeys` (see Record Date).
         */
        const ROSTER_COLUMNS = [
          {
            key: "RecordDate",
            label: "Record Date",
            altKeys: ["Record_x0020_Date", "Record0", "DateArrived", "Date_x0020_Arrived"],
          },
          { key: "DoDID", label: "DoDID" },
          { key: "Rank", label: "Rank" },
          { key: "LastName", label: "LastName" },
          { key: "FirstName", label: "FirstName" },
          { key: "MiddleInitial", label: "Middle Initial" },
          { key: "Address", label: "Address" },
          { key: "CellPhone", label: "CellPhone" },
          { key: "WorkPhone", label: "WorkPhone" },
          { key: "Status", label: "Status" },
          { key: "OfficeSymbol", label: "OfficeSymbol" },
          { key: "Squadron", label: "Squadron" },
          { key: "Notes", label: "Notes" },
        ];

        /**
         * Training Hub form (`Id` omitted). Identity | Contact. Contact: Address / phones, then Status · OfficeSymbol ·
         * Squadron on one row, then Notes (height trimmed to align column bottoms with Identity).
         */
        const ADD_FORM_FIELD_GROUPS = [
          { title: "Identity & record", keys: ["RecordDate", "DoDID", "Rank", "LastName", "FirstName", "MiddleInitial"], split: "left", layout: "stack" },
          {
            title: "Contact",
            keys: ["Address", "CellPhone", "WorkPhone", "Status", "OfficeSymbol", "Squadron", "Notes"],
            split: "right",
            layout: "contact-stack",
          },
        ];

        // ============================================================

        function resolveTopBannerImageUrl() {
          const explicit = typeof HUB_TOP_BANNER_URL === "string" ? HUB_TOP_BANNER_URL.trim() : "";
          if (explicit) return explicit;
          const root = (typeof PERSONNEL_SITE_ROOT_URL === "string" ? PERSONNEL_SITE_ROOT_URL : "")
            .trim()
            .replace(/\/$/, "");
          const rel = (typeof HUB_BANNER_SITE_RELATIVE_PATH === "string" ? HUB_BANNER_SITE_RELATIVE_PATH : "")
            .trim()
            .replace(/^\//, "");
          if (!root || !rel) return "";
          return root + "/" + rel;
        }

        function resolveProfileImageUrl() {
          const profileExplicit =
            typeof HUB_PROFILE_IMAGE_URL === "string" ? HUB_PROFILE_IMAGE_URL.trim() : "";
          if (profileExplicit) return profileExplicit;
          const logoLegacy = typeof HUB_LOGO_URL === "string" ? HUB_LOGO_URL.trim() : "";
          if (logoLegacy) return logoLegacy;
          const root = (typeof PERSONNEL_SITE_ROOT_URL === "string" ? PERSONNEL_SITE_ROOT_URL : "")
            .trim()
            .replace(/\/$/, "");
          const rel = (typeof HUB_PROFILE_SITE_RELATIVE_PATH === "string" ? HUB_PROFILE_SITE_RELATIVE_PATH : "")
            .trim()
            .replace(/^\//, "");
          if (!root || !rel) return "";
          return root + "/" + rel;
        }

        const out = document.getElementById("probeOut");
        const btn = document.getElementById("probeRun");
        const recallRosterBtn = document.getElementById("recallRosterBtn");
        const rosterReadState = document.getElementById("rosterReadState");
        const rosterTableBody = document.getElementById("rosterTableBody");
        const personDetailSection = document.getElementById("personDetailSection");
        const personDetailTitle = document.getElementById("personDetailTitle");
        const personDetailContent = document.getElementById("personDetailContent");
        const personDetailReadState = document.getElementById("personDetailReadState");
        const personDetailBackLink = document.getElementById("personDetailBackLink");
        const personDetailEditBtn = document.getElementById("personDetailEditBtn");
        const personDetailSaveBtn = document.getElementById("personDetailSaveBtn");
        const personDetailCancelBtn = document.getElementById("personDetailCancelBtn");
        const recallRosterSection = document.querySelector(".hub-section--recall");
        const hubBootOverlay = document.getElementById("hubBootOverlay");
        const hubRoot = document.getElementById("sp-pip-ui");

        /** Cached roster + REST context so Record / Back never reload the SharePoint page. */
        let hubSession = {
          rows: null,
          meta: null,
          pw: null,
          seg: null,
          sampleRow: null,
        };

        /** Active personnel record view (read / edit). */
        let personDetailSession = {
          item: null,
          editing: false,
          meta: null,
          pw: null,
          seg: null,
          sampleRow: null,
        };

        function dismissHubBootOverlay() {
          if (hubBootOverlay) hubBootOverlay.hidden = true;
          if (hubRoot) hubRoot.classList.add("hub-ready");
        }

        /** True when SharePoint page canvas is in authoring/edit — hub must not run REST or heavy DOM work. */
        function isSharePointAuthoringContext() {
          const urlPatterns = [/mode=edit/i, /action=edit/i, /layoutview=edit/i, /preview=1/i, /[?&]e=[a-f0-9-]{8,}/i];
          const wins = [];
          try {
            wins.push(window);
          } catch (_) {}
          try {
            if (window.parent && window.parent !== window) wins.push(window.parent);
          } catch (_) {}
          try {
            if (window.top && window.top !== window) wins.push(window.top);
          } catch (_) {}
          for (let w = 0; w < wins.length; w++) {
            const win = wins[w];
            try {
              const combined = String((win.location && win.location.href) || "") + String((win.location && win.location.hash) || "");
              for (let p = 0; p < urlPatterns.length; p++) {
                if (urlPatterns[p].test(combined)) return true;
              }
            } catch (_) {}
          }
          const docs = [];
          try {
            docs.push(document);
          } catch (_) {}
          try {
            if (window.parent && window.parent.document) docs.push(window.parent.document);
          } catch (_) {}
          try {
            if (window.top && window.top.document) docs.push(window.top.document);
          } catch (_) {}
          for (let d = 0; d < docs.length; d++) {
            const doc = docs[d];
            if (!doc || !doc.querySelector) continue;
            if (doc.querySelector('[data-automation-id="publishButton"]')) return true;
            if (doc.querySelector('[data-automation-id="SaveAndCloseButton"]')) return true;
            if (doc.querySelector('[data-automation-id="pageCommandBar"]')) return true;
            if (doc.querySelector('[data-automation-id="CanvasControl"]')) return true;
            if (doc.querySelector(".CanvasZone--edit")) return true;
            if (doc.querySelector(".CanvasZoneEdit") || doc.querySelector(".CanvasSection--edit")) return true;
            if (doc.querySelector("#spPageCanvasContent[contenteditable='true']")) return true;
          }
          try {
            if (window.frameElement && window.parent && window.parent !== window && window.parent.document) {
              const pdoc = window.parent.document;
              if (pdoc.querySelector('[data-automation-id="publishButton"]')) return true;
              if (pdoc.querySelector('[data-automation-id="pageCommandBar"]')) return true;
            }
          } catch (_) {}
          const ctxs = [];
          try {
            if (window._spPageContextInfo) ctxs.push(window._spPageContextInfo);
          } catch (_) {}
          try {
            if (window.parent && window.parent._spPageContextInfo) ctxs.push(window.parent._spPageContextInfo);
          } catch (_) {}
          try {
            if (window.top && window.top._spPageContextInfo) ctxs.push(window.top._spPageContextInfo);
          } catch (_) {}
          for (let c = 0; c < ctxs.length; c++) {
            const ctx = ctxs[c];
            if (ctx && (ctx.isWebEditor || ctx.isSiteEditor || ctx.inWikiEditMode)) return true;
          }
          return false;
        }

        function enterAuthoringSafeMode() {
          const boot = document.getElementById("hubBootOverlay");
          const placeholder = document.getElementById("hubAuthoringPlaceholder");
          const root = document.getElementById("sp-pip-ui");
          if (boot) boot.hidden = true;
          if (placeholder) placeholder.hidden = false;
          if (root) root.classList.add("hub-ready", "hub-authoring-mode");
        }

        function initSharePointAuthoringSafeMode() {
          if (!isSharePointAuthoringContext()) return false;
          enterAuthoringSafeMode();
          return true;
        }

        function rosterTableIsRendered() {
          return !!(rosterTableBody && rosterTableBody.querySelector("tr"));
        }

        async function ensureRosterViewRendered() {
          if (!hubSession.rows || !hubSession.meta || !hubSession.pw || !hubSession.seg) return;
          if (rosterTableIsRendered() && document.getElementById("newPersonForm")) return;
          await renderAddPersonForm(hubSession.meta, hubSession.pw, hubSession.seg, hubSession.sampleRow);
          renderRosterTable(hubSession.rows, hubSession.meta, hubSession.pw, hubSession.seg);
        }

        function updateHubHistoryForPerson(itemId, replace) {
          /* No-op: changing the URL breaks SharePoint page save/edit. */
        }

        function navigateToPersonDetail(itemId) {
          void showPersonDetailById(itemId, hubSession.meta, hubSession.pw, hubSession.seg, hubSession.rows);
        }

        async function navigateToRoster() {
          if (personDetailSection) personDetailSection.hidden = true;
          personDetailSession = { item: null, editing: false, meta: null, pw: null, seg: null, sampleRow: null };
          setPersonDetailEditMode(false);
          setHubListViewVisible(true);
          await ensureRosterViewRendered();
        }

        function hubPageBaseUrl() {
          const explicit = String(PERSONNEL_DETAIL_PAGE_URL || "").trim();
          const raw = explicit || String(location.href).split("#")[0].split("?")[0];
          try {
            return new URL(raw, location.href).href.split("#")[0].split("?")[0];
          } catch (_) {
            return raw;
          }
        }

        function rosterPageUrl() {
          return hubPageBaseUrl();
        }

        function personDetailUrl(itemId) {
          const param = String(PERSONNEL_DETAIL_QUERY_PARAM || "personId").trim() || "personId";
          const u = new URL(hubPageBaseUrl(), location.href);
          u.searchParams.set(param, String(itemId));
          return u.href;
        }

        function getPersonIdFromUrl() {
          const param = String(PERSONNEL_DETAIL_QUERY_PARAM || "personId").trim() || "personId";
          try {
            const v = new URL(location.href).searchParams.get(param);
            if (v != null && /^\d+$/.test(String(v).trim())) return String(v).trim();
          } catch (_) {}
          return null;
        }

        function setHubListViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = !visible;
          });
          if (personDetailSection) personDetailSection.hidden = visible;
          if (recallRosterSection) recallRosterSection.hidden = !visible;
        }

        function setPersonDetailEditMode(editing) {
          personDetailSession.editing = !!editing;
          if (personDetailEditBtn) personDetailEditBtn.hidden = !!editing;
          if (personDetailSaveBtn) personDetailSaveBtn.hidden = !editing;
          if (personDetailCancelBtn) personDetailCancelBtn.hidden = !editing;
        }

        function setPersonDetailState(kind, message) {
          if (!personDetailReadState) return;
          if (!message) {
            personDetailReadState.hidden = true;
            personDetailReadState.textContent = "";
            return;
          }
          personDetailReadState.hidden = false;
          personDetailReadState.className = "read-state " + kind;
          personDetailReadState.textContent = message;
        }

        function itemFieldText(item, fieldKey) {
          const cols = normalizedRosterColumns();
          const col = cols.find(function (c) {
            return c.key === fieldKey;
          });
          const keys = col ? col.tryKeys || [col.key] : sortKeysForPersonnelName(fieldKey);
          const raw = valueFromItemByKeys(item, keys);
          return raw !== undefined && raw !== null ? formatCellValue(raw) : "";
        }

        function formatPersonDisplayName(item) {
          const rank = itemFieldText(item, "Rank");
          const last = itemFieldText(item, "LastName");
          const first = itemFieldText(item, "FirstName");
          const mi = itemFieldText(item, "MiddleInitial");
          let name = [last, first].filter(Boolean).join(", ");
          if (mi) name = name ? name + " " + mi + "." : mi + ".";
          if (rank && name) return rank + " " + name;
          return rank || name || "Personnel record";
        }

        function escHtml(text) {
          return String(text != null ? text : "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
        }

        function recallRosterFieldText(item, fieldKey) {
          const text = itemFieldText(item, fieldKey);
          return text === "" ? "—" : text;
        }

        function buildRecallRosterTableRows(personnelRows) {
          const sorted = sortPersonnelRowsAlphabetical(Array.isArray(personnelRows) ? personnelRows.slice() : []);
          return sorted.map(function (item) {
            return {
              name: formatPersonDisplayName(item),
              address: recallRosterFieldText(item, "Address"),
              cell: recallRosterFieldText(item, "CellPhone"),
              work: recallRosterFieldText(item, "WorkPhone"),
            };
          });
        }

        async function getRecallRosterPersonnelRows() {
          if (Array.isArray(hubSession.rows) && hubSession.rows.length) {
            return sortPersonnelRowsAlphabetical(hubSession.rows.slice());
          }
          const pw = hubSession.pw || resolvedPersonnelRestBase();
          const seg = hubSession.seg || personnelListApiPath();
          if (!pw) throw new Error("Personnel list has not loaded yet. Wait for the roster, then try again.");
          const items = await spFetch(`/_api/web/${seg}/items?$top=500`, {}, pw);
          const rows = sortPersonnelRowsAlphabetical((items && items.value) || []);
          hubSession.rows = rows;
          hubSession.pw = pw;
          hubSession.seg = seg;
          return rows;
        }

        function loadExternalScript(src) {
          return new Promise(function (resolve, reject) {
            const existing = document.querySelector('script[data-recall-src="' + src + '"]');
            if (existing) {
              if (existing.getAttribute("data-loaded") === "1") resolve();
              else existing.addEventListener("load", function () { resolve(); }, { once: true });
              return;
            }
            const s = document.createElement("script");
            s.src = src;
            s.async = true;
            s.setAttribute("data-recall-src", src);
            s.onload = function () {
              s.setAttribute("data-loaded", "1");
              resolve();
            };
            s.onerror = function () {
              reject(new Error("Could not load script: " + src));
            };
            const host = document.getElementById("sp-pip-ui") || document.body;
            host.appendChild(s);
          });
        }

        async function generateRecallRosterPdf(personnelRows) {
          await loadExternalScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");
          await loadExternalScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js");
          if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF library unavailable");
          const recallRows = buildRecallRosterTableRows(personnelRows);
          const doc = new window.jspdf.jsPDF({ orientation: "landscape", unit: "pt", format: "letter" });
          const generated = new Date().toLocaleString();
          doc.setFontSize(14);
          doc.text(RECALL_ROSTER_ORG_TITLE, 40, 36);
          doc.setFontSize(11);
          doc.text(RECALL_ROSTER_DOC_TITLE, 40, 54);
          doc.setFontSize(9);
          doc.text("Generated: " + generated + "   ·   Personnel: " + recallRows.length, 40, 70);
          doc.autoTable({
            startY: 82,
            head: [["Name", "Address", "Cell phone", "Work phone"]],
            body: recallRows.map(function (r) {
              return [r.name, r.address, r.cell, r.work];
            }),
            styles: { fontSize: 8, cellPadding: 4, overflow: "linebreak" },
            headStyles: { fillColor: [14, 36, 26], textColor: [255, 204, 102] },
            columnStyles: {
              0: { cellWidth: 130 },
              1: { cellWidth: 250 },
              2: { cellWidth: 90 },
              3: { cellWidth: 90 },
            },
            margin: { left: 40, right: 40 },
          });
          doc.save(RECALL_ROSTER_FILENAME);
        }

        function printRecallRosterFallback(personnelRows) {
          const recallRows = buildRecallRosterTableRows(personnelRows);
          const generated = new Date().toLocaleString();
          const rowsHtml = recallRows
            .map(function (r) {
              return (
                "<tr><td>" +
                escHtml(r.name) +
                "</td><td>" +
                escHtml(r.address) +
                "</td><td>" +
                escHtml(r.cell) +
                "</td><td>" +
                escHtml(r.work) +
                "</td></tr>"
              );
            })
            .join("");
          const html =
            "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\" /><title>" +
            escHtml(RECALL_ROSTER_DOC_TITLE) +
            "</title><style>" +
            "body{font-family:Arial,sans-serif;margin:24px;color:#111;}" +
            "h1{font-size:18px;margin:0 0 4px;}" +
            "h2{font-size:14px;margin:0 0 12px;font-weight:normal;}" +
            ".meta{font-size:11px;color:#444;margin-bottom:16px;}" +
            "table{width:100%;border-collapse:collapse;font-size:11px;}" +
            "th,td{border:1px solid #333;padding:6px 8px;text-align:left;vertical-align:top;}" +
            "th{background:#1e4a36;color:#ffcc66;}" +
            "tr:nth-child(even) td{background:#f4f4f4;}" +
            "</style></head><body>" +
            "<h1>" +
            escHtml(RECALL_ROSTER_ORG_TITLE) +
            "</h1>" +
            "<h2>" +
            escHtml(RECALL_ROSTER_DOC_TITLE) +
            "</h2>" +
            "<p class=\"meta\">Generated: " +
            escHtml(generated) +
            " · Personnel: " +
            recallRows.length +
            "</p>" +
            "<table><thead><tr><th>Name</th><th>Address</th><th>Cell phone</th><th>Work phone</th></tr></thead><tbody>" +
            rowsHtml +
            "</tbody></table></body></html>";
          const w = window.open("", "_blank");
          if (!w) {
            throw new Error("Pop-up blocked. Allow pop-ups, then try again.");
          }
          w.document.open();
          w.document.write(html);
          w.document.close();
          w.focus();
          w.onload = function () {
            w.print();
          };
        }

        async function handleRecallRosterDownload() {
          if (!recallRosterBtn) return;
          const labelDefault = "Save recall roster (PDF)";
          recallRosterBtn.disabled = true;
          recallRosterBtn.textContent = "Building PDF…";
          try {
            const rows = await getRecallRosterPersonnelRows();
            if (!rows.length) {
              setReadState("warn", "No personnel on the list. Add people first, then try again.");
              return;
            }
            try {
              await generateRecallRosterPdf(rows);
            } catch (pdfErr) {
              log("PDF library unavailable; opening print view:\n" + (pdfErr.message || String(pdfErr)), "err");
              printRecallRosterFallback(rows);
            }
          } catch (e) {
            setReadState("err", "Recall roster failed: " + (e.message || String(e)).slice(0, 280));
          } finally {
            recallRosterBtn.disabled = false;
            recallRosterBtn.textContent = labelDefault;
          }
        }

        function buildDetailFieldWrap(col, item) {
          const tryKeys = col.tryKeys || [col.key];
          const raw = valueFromItemByKeys(item, tryKeys);
          const text = raw !== undefined && raw !== null ? formatCellValue(raw) : "";
          const fwrap = document.createElement("div");
          fwrap.className = "add-field person-detail-field";
          if (col.key === "Address") fwrap.classList.add("add-field--address");
          if (col.key === "Notes") fwrap.classList.add("add-field--notes");
          const lab = document.createElement("div");
          lab.className = "person-detail-label";
          lab.textContent = col.label;
          const val = document.createElement("div");
          val.className = "person-detail-value";
          val.textContent = text === "" ? "—" : text;
          fwrap.appendChild(lab);
          fwrap.appendChild(val);
          return fwrap;
        }

        function buildDetailGroupFieldset(group, cols, item) {
          if (!group || !group.title || !Array.isArray(group.keys)) return null;

          if (group.layout === "contact-stack") {
            const keys = group.keys;
            if (keys.length < 7) return null;
            const topKeys = keys.slice(0, 3);
            const statusKeys = keys.slice(3, 6);
            const notesKey = keys[6];
            const bundle = document.createElement("div");
            bundle.className = "add-form-contact-bundle";

            const topGrid = document.createElement("div");
            topGrid.className = "add-form-grid add-form-grid--stack add-form-contact-stack-top";
            topKeys.forEach(function (key) {
              const col = cols.find(function (c) {
                return c.key === key;
              });
              if (!col) return;
              const f = buildDetailFieldWrap(col, item);
              if (f) topGrid.appendChild(f);
            });

            const statusRow = document.createElement("div");
            statusRow.className = "add-form-contact-status-row";
            statusKeys.forEach(function (key) {
              const col = cols.find(function (c) {
                return c.key === key;
              });
              if (!col) return;
              const f = buildDetailFieldWrap(col, item);
              if (f) statusRow.appendChild(f);
            });

            const notesCol = cols.find(function (c) {
              return c.key === notesKey;
            });
            const notesWrap = notesCol ? buildDetailFieldWrap(notesCol, item) : null;

            bundle.appendChild(topGrid);
            bundle.appendChild(statusRow);
            if (notesWrap) bundle.appendChild(notesWrap);

            const fs = document.createElement("fieldset");
            fs.className = "add-form-group";
            const leg = document.createElement("legend");
            leg.className = "add-form-group-legend";
            leg.textContent = group.title;
            fs.appendChild(leg);
            fs.appendChild(bundle);
            return fs;
          }

          const grid = document.createElement("div");
          grid.className = "add-form-grid";
          if (group.layout === "stack") grid.classList.add("add-form-grid--stack");
          group.keys.forEach(function (key) {
            const col = cols.find(function (c) {
              return c.key === key;
            });
            if (!col) return;
            const fwrap = buildDetailFieldWrap(col, item);
            if (fwrap) grid.appendChild(fwrap);
          });
          if (!grid.childElementCount) return null;
          const fs = document.createElement("fieldset");
          fs.className = "add-form-group";
          const leg = document.createElement("legend");
          leg.className = "add-form-group-legend";
          leg.textContent = group.title;
          fs.appendChild(leg);
          fs.appendChild(grid);
          return fs;
        }

        function setFieldValueFromItem(el, col, item) {
          if (!el || !item) return;
          const tryKeys = col.tryKeys || [col.key];
          const raw = valueFromItemByKeys(item, tryKeys);
          if (raw === undefined || raw === null || raw === "") return;
          const text = formatCellValue(raw);
          if (el.type === "date") {
            const m = text.match(/^(\d{4}-\d{2}-\d{2})/);
            if (m) el.value = m[1];
            return;
          }
          el.value = text;
        }

        function buildEditGroupFieldset(group, cols, sampleRow, item) {
          if (!group || !group.title || !Array.isArray(group.keys)) return null;
          const fieldOpt = { idPrefix: "pf_", sampleRow: sampleRow };

          if (group.layout === "contact-stack") {
            const keys = group.keys;
            if (keys.length < 7) return null;
            const topKeys = keys.slice(0, 3);
            const statusKeys = keys.slice(3, 6);
            const notesKey = keys[6];
            const bundle = document.createElement("div");
            bundle.className = "add-form-contact-bundle";

            const topGrid = document.createElement("div");
            topGrid.className = "add-form-grid add-form-grid--stack add-form-contact-stack-top";
            topKeys.forEach(function (key) {
              const col = cols.find(function (c) {
                return c.key === key;
              });
              if (!col) return;
              const f = buildAddFieldWrap(col, sampleRow, fieldOpt);
              if (!f) return;
              const input = f.querySelector("input, select, textarea");
              if (input) setFieldValueFromItem(input, col, item);
              topGrid.appendChild(f);
            });

            const statusRow = document.createElement("div");
            statusRow.className = "add-form-contact-status-row";
            statusKeys.forEach(function (key) {
              const col = cols.find(function (c) {
                return c.key === key;
              });
              if (!col) return;
              const f = buildAddFieldWrap(col, sampleRow, Object.assign({ fullWidthSelect: true }, fieldOpt));
              if (!f) return;
              const input = f.querySelector("input, select, textarea");
              if (input) setFieldValueFromItem(input, col, item);
              statusRow.appendChild(f);
            });

            const notesCol = cols.find(function (c) {
              return c.key === notesKey;
            });
            let notesWrap = null;
            if (notesCol) {
              notesWrap = buildAddFieldWrap(notesCol, sampleRow, fieldOpt);
              if (notesWrap) {
                const input = notesWrap.querySelector("textarea");
                if (input) setFieldValueFromItem(input, notesCol, item);
              }
            }

            bundle.appendChild(topGrid);
            bundle.appendChild(statusRow);
            if (notesWrap) bundle.appendChild(notesWrap);

            const fs = document.createElement("fieldset");
            fs.className = "add-form-group";
            const leg = document.createElement("legend");
            leg.className = "add-form-group-legend";
            leg.textContent = group.title;
            fs.appendChild(leg);
            fs.appendChild(bundle);
            return fs;
          }

          const grid = document.createElement("div");
          grid.className = "add-form-grid";
          if (group.layout === "stack") grid.classList.add("add-form-grid--stack");
          group.keys.forEach(function (key) {
            const col = cols.find(function (c) {
              return c.key === key;
            });
            if (!col) return;
            const fwrap = buildAddFieldWrap(col, sampleRow, fieldOpt);
            if (!fwrap) return;
            const input = fwrap.querySelector("input, select, textarea");
            if (input) setFieldValueFromItem(input, col, item);
            grid.appendChild(fwrap);
          });
          if (!grid.childElementCount) return null;
          const fs = document.createElement("fieldset");
          fs.className = "add-form-group";
          const leg = document.createElement("legend");
          leg.className = "add-form-group-legend";
          leg.textContent = group.title;
          fs.appendChild(leg);
          fs.appendChild(grid);
          return fs;
        }

        async function populatePersonEditDropdowns(seg, pw, sampleRow, root) {
          const cols = normalizedRosterColumns();
          for (let i = 0; i < cols.length; i++) {
            const col = cols[i];
            if (col.key === "Id") continue;
            if (DROPDOWN_COLUMN_KEYS.indexOf(col.key) === -1) continue;
            const writeKey = resolveWriteKey(col, sampleRow);
            const sel = root.querySelector("#pf_" + col.key);
            if (!sel || sel.tagName !== "SELECT") continue;
            const keepValue = sel.value;
            while (sel.options.length > 1) sel.remove(1);
            try {
              const r = await fetchChoiceOptions(seg, pw, writeKey);
              if (r.choices && r.choices.length) {
                r.choices.forEach(function (c) {
                  const o = document.createElement("option");
                  o.value = c;
                  o.textContent = c;
                  sel.appendChild(o);
                });
              }
            } catch (_) {}
            if (keepValue) sel.value = keepValue;
            applySelectAutosize(sel);
          }
          applyAllSelectAutosizes(root);
        }

        function buildPersonDetailFields(item, editing) {
          const cols = normalizedRosterColumns().filter(function (c) {
            return c.key !== "Id";
          });
          const splitWrap = document.createElement("div");
          splitWrap.className = "add-form-split add-form-split--primary";
          const sampleRow = personDetailSession.sampleRow || item;

          if (Array.isArray(ADD_FORM_FIELD_GROUPS)) {
            ADD_FORM_FIELD_GROUPS.forEach(function (group) {
              const fs = editing
                ? buildEditGroupFieldset(group, cols, sampleRow, item)
                : buildDetailGroupFieldset(group, cols, item);
              if (fs) splitWrap.appendChild(fs);
            });
          }

          if (!splitWrap.childElementCount) {
            const grid = document.createElement("div");
            grid.className = "add-form-grid add-form-grid--stack";
            cols.forEach(function (col) {
              if (editing) {
                const f = buildAddFieldWrap(col, sampleRow, { idPrefix: "pf_" });
                if (!f) return;
                const input = f.querySelector("input, select, textarea");
                if (input) setFieldValueFromItem(input, col, item);
                grid.appendChild(f);
              } else {
                grid.appendChild(buildDetailFieldWrap(col, item));
              }
            });
            return grid;
          }

          return splitWrap;
        }

        async function renderPersonDetailView(item, editing) {
          if (!personDetailContent || !item) return;
          personDetailContent.innerHTML = "";
          const cols = normalizedRosterColumns().filter(function (c) {
            return c.key !== "Id";
          });
          if (!cols.length) {
            setPersonDetailState("warn", "No columns configured to display.");
            return;
          }
          setPersonDetailEditMode(!!editing);
          setPersonDetailState("", "");
          if (personDetailTitle) personDetailTitle.textContent = formatPersonDisplayName(item);

          const fields = buildPersonDetailFields(item, editing);
          if (editing) {
            const form = document.createElement("div");
            form.id = "personEditForm";
            form.className = "add-form-stack";
            form.setAttribute("role", "form");
            form.appendChild(fields);
            personDetailContent.appendChild(form);
            await populatePersonEditDropdowns(personDetailSession.seg, personDetailSession.pw, personDetailSession.sampleRow, form);
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                balanceContactNotesToIdentityColumn(form);
              });
            });
          } else {
            personDetailContent.appendChild(fields);
          }
        }

        function syncPersonInHubSession(updatedItem) {
          if (!updatedItem || updatedItem.Id == null) return;
          if (!Array.isArray(hubSession.rows)) hubSession.rows = [];
          const idx = hubSession.rows.findIndex(function (r) {
            return r && String(r.Id) === String(updatedItem.Id);
          });
          if (idx === -1) hubSession.rows.push(updatedItem);
          else hubSession.rows[idx] = Object.assign({}, hubSession.rows[idx], updatedItem);
          hubSession.rows = sortPersonnelRowsAlphabetical(hubSession.rows);
          if (rosterTableIsRendered() && hubSession.meta && hubSession.pw && hubSession.seg) {
            renderRosterTable(hubSession.rows, hubSession.meta, hubSession.pw, hubSession.seg);
          }
        }

        async function savePersonDetailEdits() {
          const s = personDetailSession;
          if (!s.item || s.item.Id == null || !s.pw || !s.seg) return;
          const cols = normalizedRosterColumns();
          let lastName = "";
          let firstName = "";
          let dod = "";
          cols.forEach(function (col) {
            if (col.key === "Id") return;
            const el = document.getElementById("pf_" + col.key);
            if (!el) return;
            const v = String(el.value || "").trim();
            if (col.key === "LastName") lastName = v;
            if (col.key === "FirstName") firstName = v;
            if (col.key === "DoDID") dod = v;
          });
          if (!lastName || !firstName) {
            setPersonDetailState("err", "Last name and First name are required.");
            return;
          }

          const payload = {};
          cols.forEach(function (col) {
            if (col.key === "Id") return;
            const el = document.getElementById("pf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWriteKey(col, s.sampleRow);
            let v = String(el.value || "").trim();
            if (v === "") return;
            if (el.type === "date") {
              payload[writeKey] = v + "T00:00:00Z";
              return;
            }
            payload[writeKey] = v;
          });

          if (SET_TITLE_ON_CREATE) {
            payload.Title = [lastName, firstName].filter(Boolean).join(", ") || dod || "Personnel record";
          }

          try {
            setPersonDetailState("loading", "Saving changes…");
            if (personDetailSaveBtn) personDetailSaveBtn.disabled = true;
            await spFetch(`/_api/web/${s.seg}/items(${s.item.Id})`, { method: "MERGE", body: payload }, s.pw);
            const updated = Object.assign({}, s.item, demoPayloadToItem(payload));
            updated.Id = s.item.Id;
            personDetailSession.item = updated;
            syncPersonInHubSession(updated);
            if (personDetailTitle) personDetailTitle.textContent = formatPersonDisplayName(updated);
            await renderPersonDetailView(updated, false);
            setPersonDetailState("ok", "Record saved.");
            window.setTimeout(function () {
              setPersonDetailState("", "");
            }, 2500);
          } catch (e) {
            setPersonDetailState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
          } finally {
            if (personDetailSaveBtn) personDetailSaveBtn.disabled = false;
          }
        }

        function renderPersonDetailContent(item) {
          void renderPersonDetailView(item, false);
        }

        async function openPersonDetailRecord(itemId, meta, pw, seg, cachedRows, sampleRow) {
          if (!personDetailSection || !personDetailContent) return;
          setHubListViewVisible(false);
          personDetailSection.hidden = false;
          setPersonDetailEditMode(false);

          const rows = Array.isArray(cachedRows) ? cachedRows : hubSession.rows;
          let item =
            Array.isArray(rows) &&
            rows.find(function (r) {
              return r && String(r.Id) === String(itemId);
            });

          if (!item) {
            personDetailContent.innerHTML = "";
            setPersonDetailState("loading", "Loading personnel record…");
            if (personDetailTitle) personDetailTitle.textContent = "Personnel record";
            try {
              item = await spFetch(`/_api/web/${seg}/items(${itemId})`, {}, pw);
            } catch (e) {
              setPersonDetailState("err", "Cannot load personnel record: " + (e.message || String(e)).slice(0, 280));
              return;
            }
          }

          personDetailSession = {
            item: item,
            editing: false,
            meta: meta || hubSession.meta,
            pw: pw || hubSession.pw,
            seg: seg || hubSession.seg,
            sampleRow: sampleRow || (Array.isArray(rows) && rows.length ? rows[0] : item),
          };

          await renderPersonDetailView(item, false);
        }

        async function showPersonDetailById(itemId, meta, pw, seg, cachedRows) {
          const sampleRow =
            (Array.isArray(cachedRows) && cachedRows.length ? cachedRows[0] : null) || hubSession.sampleRow;
          await openPersonDetailRecord(itemId, meta, pw, seg, cachedRows, sampleRow);
        }

        if (personDetailBackLink) {
          personDetailBackLink.href = "#";
          personDetailBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (personDetailEditBtn) {
          personDetailEditBtn.addEventListener("click", function () {
            if (!personDetailSession.item) return;
            void renderPersonDetailView(personDetailSession.item, true);
          });
        }

        if (personDetailSaveBtn) {
          personDetailSaveBtn.addEventListener("click", function () {
            void savePersonDetailEdits();
          });
        }

        if (personDetailCancelBtn) {
          personDetailCancelBtn.addEventListener("click", function () {
            if (!personDetailSession.item) return;
            void renderPersonDetailView(personDetailSession.item, false);
          });
        }

        let demoPersonnelStore = null;

        function isLocalDemoMode() {
          const mode = LOCAL_DEMO_MODE;
          if (mode === true || String(mode).toLowerCase() === "true") return true;
          if (mode === false || String(mode).toLowerCase() === "false") return false;
          const h = String(location.hostname || "").toLowerCase();
          if (h === "localhost" || h === "127.0.0.1" || h === "[::1]") return true;
          if (location.protocol === "file:") return true;
          try {
            if (new URL(location.href).searchParams.get("demo") === "1") return true;
          } catch (_) {}
          return false;
        }

        function updateDemoModeBanner() {
          const banner = document.getElementById("demoModeBanner");
          if (banner) banner.hidden = !isLocalDemoMode();
        }

        function createDemoPersonnelSeed() {
          return [
            {
              Id: 1,
              RecordDate: "2024-03-12",
              DoDID: "1234567890",
              Rank: "TSgt",
              LastName: "Anderson",
              FirstName: "James",
              MiddleInitial: "R",
              Address: "123 Main St, Wright-Patterson AFB, OH 45433",
              CellPhone: "937-555-0101",
              WorkPhone: "937-555-1001",
              Status: "Active",
              OfficeSymbol: "S3T",
              Squadron: "88 SFS",
              Notes: "Training coordinator. Primary POC for annual certs.",
              Title: "Anderson, James",
            },
            {
              Id: 2,
              RecordDate: "2023-11-02",
              DoDID: "2345678901",
              Rank: "SSgt",
              LastName: "Baker",
              FirstName: "Maria",
              MiddleInitial: "L",
              Address: "456 Oak Ave, Dayton, OH 45431",
              CellPhone: "937-555-0102",
              WorkPhone: "937-555-1002",
              Status: "Active",
              OfficeSymbol: "SFOPS",
              Squadron: "88 SFS",
              Notes: "Range safety certified.",
              Title: "Baker, Maria",
            },
            {
              Id: 3,
              RecordDate: "2025-01-08",
              DoDID: "3456789012",
              Rank: "SrA",
              LastName: "Chen",
              FirstName: "David",
              MiddleInitial: "K",
              Address: "789 Elm Dr, Beavercreek, OH 45440",
              CellPhone: "937-555-0103",
              WorkPhone: "937-555-1003",
              Status: "Pending",
              OfficeSymbol: "S3T",
              Squadron: "88 SFS",
              Notes: "Awaiting final background check.",
              Title: "Chen, David",
            },
            {
              Id: 4,
              RecordDate: "2022-06-20",
              DoDID: "4567890123",
              Rank: "MSgt",
              LastName: "Davis",
              FirstName: "Robert",
              MiddleInitial: "A",
              Address: "321 Pine Rd, Fairborn, OH 45324",
              CellPhone: "937-555-0104",
              WorkPhone: "937-555-1004",
              Status: "Active",
              OfficeSymbol: "CP",
              Squadron: "88 SFS",
              Notes: "NCOIC, patrol section.",
              Title: "Davis, Robert",
            },
            {
              Id: 5,
              RecordDate: "2024-09-15",
              DoDID: "5678901234",
              Rank: "A1C",
              LastName: "Evans",
              FirstName: "Sarah",
              MiddleInitial: "M",
              Address: "654 Maple Ln, Huber Heights, OH 45424",
              CellPhone: "937-555-0105",
              WorkPhone: "937-555-1005",
              Status: "TDY",
              OfficeSymbol: "S3T",
              Squadron: "88 OSS",
              Notes: "Deployed — return date TBD.",
              Title: "Evans, Sarah",
            },
            {
              Id: 6,
              RecordDate: "2021-02-28",
              DoDID: "6789012345",
              Rank: "Capt",
              LastName: "Foster",
              FirstName: "Michael",
              MiddleInitial: "J",
              Address: "987 Cedar Ct, Wright-Patterson AFB, OH 45433",
              CellPhone: "937-555-0106",
              WorkPhone: "937-555-1006",
              Status: "Active",
              OfficeSymbol: "CC",
              Squadron: "88 SFS",
              Notes: "Commander, 88 SFS.",
              Title: "Foster, Michael",
            },
            {
              Id: 7,
              RecordDate: "2023-07-04",
              DoDID: "7890123456",
              Rank: "Amn",
              LastName: "Garcia",
              FirstName: "Elena",
              MiddleInitial: "S",
              Address: "147 Birch Way, Kettering, OH 45429",
              CellPhone: "937-555-0107",
              WorkPhone: "937-555-1007",
              Status: "Active",
              OfficeSymbol: "S3T",
              Squadron: "88 SFS",
              Notes: "New hire — in-process training.",
              Title: "Garcia, Elena",
            },
            {
              Id: 8,
              RecordDate: "2020-12-01",
              DoDID: "8901234567",
              Rank: "SMSgt",
              LastName: "Harris",
              FirstName: "William",
              MiddleInitial: "T",
              Address: "258 Spruce St, Centerville, OH 45458",
              CellPhone: "937-555-0108",
              WorkPhone: "937-555-1008",
              Status: "Inactive",
              OfficeSymbol: "SFOPS",
              Squadron: "88 SFS",
              Notes: "Retired — archive record only.",
              Title: "Harris, William",
            },
          ];
        }

        function initDemoPersonnelStore() {
          if (!demoPersonnelStore) {
            demoPersonnelStore = createDemoPersonnelSeed();
          }
          return demoPersonnelStore;
        }

        function demoChoicesForField(fieldName) {
          const name = String(fieldName || "").trim();
          if (DEMO_FIELD_CHOICES[name]) return DEMO_FIELD_CHOICES[name].slice();
          const cols = normalizedRosterColumns();
          const col = cols.find(function (c) {
            const keys = c.tryKeys || [c.key];
            return keys.indexOf(name) !== -1 || c.key === name;
          });
          if (col && DEMO_FIELD_CHOICES[col.key]) return DEMO_FIELD_CHOICES[col.key].slice();
          return [];
        }

        function demoPayloadToItem(body) {
          const item = {};
          const cols = normalizedRosterColumns();
          Object.keys(body || {}).forEach(function (key) {
            if (key === "Title") {
              item.Title = body[key];
              return;
            }
            const col = cols.find(function (c) {
              const writeKey = c.saveKey || c.key;
              const tryKeys = c.tryKeys || [c.key];
              return writeKey === key || tryKeys.indexOf(key) !== -1 || c.key === key;
            });
            const targetKey = col ? col.key : key;
            let val = body[key];
            if (typeof val === "string" && /T00:00:00Z$/.test(val)) val = val.slice(0, 10);
            item[targetKey] = val;
          });
          if (!item.Title && item.LastName && item.FirstName) {
            item.Title = item.LastName + ", " + item.FirstName;
          }
          return item;
        }

        function demoCloneStore() {
          return JSON.parse(JSON.stringify(initDemoPersonnelStore()));
        }

        async function demoSpFetch(path, opts, method) {
          opts = opts || {};
          method = method || (opts.method || "GET").toUpperCase();
          const store = initDemoPersonnelStore();
          await new Promise(function (resolve) {
            setTimeout(resolve, 120);
          });

          if (method === "GET" && /\/fields\/getbyinternalnameortitle\('/.test(path)) {
            const m = path.match(/getbyinternalnameortitle\('([^']*)'\)/);
            const fieldName = m ? m[1] : "";
            return {
              Choices: demoChoicesForField(fieldName),
              Title: fieldName,
              InternalName: fieldName,
              TypeAsString: "Choice",
            };
          }

          if (method === "GET" && /\/lists\?\$select=Title,Id,ItemCount,Hidden/.test(path)) {
            return {
              value: [{ Title: personnelListTitle(), Id: 101, ItemCount: store.length, Hidden: false }],
            };
          }

          if (method === "GET" && /\/lists\/(guid'[^']+'|getbytitle\([^)]+\))\?\$select=/.test(path)) {
            return {
              Id: 101,
              Title: personnelListTitle(),
              ItemCount: store.length,
              ListItemEntityTypeFullName: "SP.Data.PersonnelListItem",
            };
          }

          const itemGet = path.match(/\/items\((\d+)\)(?:\?|$)/);
          if (method === "GET" && itemGet) {
            const hit = store.find(function (r) {
              return String(r.Id) === itemGet[1];
            });
            if (!hit) throw new Error("404 Not Found The specified item does not exist.");
            return JSON.parse(JSON.stringify(hit));
          }

          if (method === "GET" && /\/items(?:\?|$)/.test(path)) {
            return { value: demoCloneStore() };
          }

          if (method === "POST" && /\/items(?:\?|$)/.test(path) && opts.body) {
            const body = typeof opts.body === "object" ? opts.body : JSON.parse(opts.body);
            const nextId =
              store.reduce(function (max, r) {
                return Math.max(max, Number(r.Id) || 0);
              }, 0) + 1;
            const item = Object.assign({ Id: nextId }, demoPayloadToItem(body));
            store.push(item);
            return JSON.parse(JSON.stringify(item));
          }

          if (method === "MERGE" && itemGet && opts.body) {
            const idx = store.findIndex(function (r) {
              return String(r.Id) === itemGet[1];
            });
            if (idx === -1) throw new Error("404 Not Found The specified item does not exist.");
            const body = typeof opts.body === "object" ? opts.body : JSON.parse(opts.body);
            Object.assign(store[idx], demoPayloadToItem(body));
            return null;
          }

          if (method === "DELETE" && itemGet) {
            const idx = store.findIndex(function (r) {
              return String(r.Id) === itemGet[1];
            });
            if (idx === -1) throw new Error("404 Not Found The specified item does not exist.");
            store.splice(idx, 1);
            return null;
          }

          throw new Error("Demo mode: unhandled " + method + " " + path);
        }

        function normalizedRosterColumns() {
          const arr = Array.isArray(ROSTER_COLUMNS) ? ROSTER_COLUMNS : [];
          const result = [];
          arr.forEach((entry) => {
            if (entry == null) return;
            if (typeof entry === "string") {
              const key = entry.trim();
              if (key) result.push({ key: key, label: key, tryKeys: [key], saveKey: null });
              return;
            }
            if (typeof entry === "object") {
              const key = String(entry.key != null ? entry.key : entry.field != null ? entry.field : "").trim();
              if (!key) return;
              const label = String(entry.label != null ? entry.label : "").trim() || key;
              const alt = Array.isArray(entry.altKeys)
                ? entry.altKeys.map((x) => String(x).trim()).filter(Boolean)
                : [];
              const tryKeys = [key];
              alt.forEach((a) => {
                if (a && tryKeys.indexOf(a) === -1) tryKeys.push(a);
              });
              const saveKey = entry.saveKey != null ? String(entry.saveKey).trim() : "";
              result.push({ key: key, label: label, tryKeys: tryKeys, saveKey: saveKey || null });
            }
          });
          return result;
        }

        /** Explicit plan from ROSTER_COLUMNS, or auto from row keys. */
        function resolveRosterColumnPlan(rows) {
          const explicit = normalizedRosterColumns();
          if (explicit.length > 0) return { mode: "explicit", columns: explicit };
          const keys = collectColumnKeys(rows);
          return { mode: "auto", columns: keys.map((k) => ({ key: k, label: k, tryKeys: [k], saveKey: null })) };
        }

        /** Columns actually drawn in the roster table (optionally strips list item Id). */
        function rosterColumnsForDisplay(plan) {
          if (!plan || !plan.columns) return [];
          if (SHOW_LIST_ITEM_ID_IN_ROSTER) return plan.columns.slice();
          return plan.columns.filter(function (c) {
            return c.key !== "Id";
          });
        }

        /** Skip OData / navigation noise; every other property becomes a table column. */
        function isODataOrMetaKey(k) {
          if (!k || typeof k !== "string") return true;
          if (k.indexOf("@") === 0) return true;
          if (k.indexOf("odata.") === 0) return true;
          return k === "odata.type" || k === "odata.id" || k === "odata.editLink" || k === "odata.metadata";
        }

        /** Union of field names across all loaded rows (order: Title, then A–Z). List item `Id` is omitted from the table. */
        function collectColumnKeys(rows) {
          const set = new Set();
          rows.forEach((item) => {
            if (!item || typeof item !== "object") return;
            Object.keys(item).forEach((k) => {
              if (!isODataOrMetaKey(k) && k !== "Id") set.add(k);
            });
          });
          const keys = Array.from(set);
          const priority = ["Title"];
          keys.sort((a, b) => {
            const ia = priority.indexOf(a);
            const ib = priority.indexOf(b);
            if (ia !== -1 || ib !== -1) {
              if (ia === -1) return 1;
              if (ib === -1) return -1;
              return ia - ib;
            }
            return a.localeCompare(b, undefined, { sensitivity: "base" });
          });
          return keys;
        }

        function formatCellValue(val) {
          if (val === null || val === undefined) return "";
          if (typeof val === "boolean") return val ? "Yes" : "No";
          if (typeof val === "number") return String(val);
          if (typeof val === "string") {
            const m = val.match(/^\/Date\((\d+)\)\/$/);
            if (m) {
              const d = new Date(parseInt(m[1], 10));
              if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
            }
            const isoDay = val.match(/^(\d{4}-\d{2}-\d{2})(?:T|$)/);
            if (isoDay) return isoDay[1];
            return val;
          }
          try {
            return JSON.stringify(val);
          } catch (_) {
            return String(val);
          }
        }

        function valueFromItemByKeys(item, tryKeys) {
          if (!item || !tryKeys || !tryKeys.length) return undefined;
          for (let i = 0; i < tryKeys.length; i++) {
            const k = tryKeys[i];
            if (!k || !Object.prototype.hasOwnProperty.call(item, k)) continue;
            const v = item[k];
            if (v !== null && v !== undefined && v !== "") return v;
          }
          for (let i = 0; i < tryKeys.length; i++) {
            const k = tryKeys[i];
            if (k && Object.prototype.hasOwnProperty.call(item, k)) return item[k];
          }
          return undefined;
        }

        function sortKeysForPersonnelName(fieldKey) {
          if (fieldKey === "LastName") return ["LastName", "Last_x0020_Name", "LastName0"];
          if (fieldKey === "FirstName") return ["FirstName", "First_x0020_Name", "FirstName0"];
          return [fieldKey];
        }

        function sortPersonnelRowsAlphabetical(rows) {
          if (!Array.isArray(rows) || rows.length < 2) return rows;
          const cols = normalizedRosterColumns();
          function columnSpecForSort(fieldKey) {
            const fromRoster = cols.find(function (c) {
              return c.key === fieldKey;
            });
            if (fromRoster) return fromRoster;
            return {
              key: fieldKey,
              label: fieldKey,
              tryKeys: sortKeysForPersonnelName(fieldKey),
              saveKey: null,
            };
          }
          const lastCol = columnSpecForSort("LastName");
          const firstCol = columnSpecForSort("FirstName");
          function nameVal(item, col) {
            if (!col) return "";
            const keys = col.tryKeys || [col.key];
            const raw = valueFromItemByKeys(item, keys);
            return String(raw != null ? raw : "")
              .trim()
              .toLowerCase();
          }
          return rows.slice().sort(function (a, b) {
            const la = nameVal(a, lastCol);
            const lb = nameVal(b, lastCol);
            if (la !== lb) return la < lb ? -1 : la > lb ? 1 : 0;
            const fa = nameVal(a, firstCol);
            const fb = nameVal(b, firstCol);
            if (fa !== fb) return fa < fb ? -1 : fa > fb ? 1 : 0;
            const ida = a.Id != null ? Number(a.Id) : 0;
            const idb = b.Id != null ? Number(b.Id) : 0;
            return ida - idb;
          });
        }

        function resolveWriteKey(col, sampleRow) {
          if (col.saveKey) return String(col.saveKey).trim();
          if (sampleRow) {
            const tryKeys = col.tryKeys || [col.key];
            const hit = tryKeys.find((k) => sampleRow && Object.prototype.hasOwnProperty.call(sampleRow, k));
            if (hit) return hit;
          }
          return col.key;
        }

        function setReadState(kind, message) {
          if (!rosterReadState) return;
          if (!message || kind === "ok" || kind === "loading") {
            rosterReadState.hidden = true;
            rosterReadState.textContent = "";
            return;
          }
          rosterReadState.hidden = false;
          rosterReadState.className = "read-state " + kind;
          rosterReadState.textContent = message;
        }

        function clearRosterTable() {
          const thead = document.getElementById("rosterThead");
          if (thead) thead.innerHTML = "";
          if (rosterTableBody) rosterTableBody.innerHTML = "";
        }

        async function deletePersonnelRow(id, pw, seg) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid)) {
            setReadState("err", "Invalid row Id for delete.");
            return;
          }
          if (!confirm("Delete personnel record Id " + sid + "? This cannot be undone.")) return;
          try {
            setReadState("loading", "Deleting record…");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            log("DELETE Id " + sid + " OK", "ok");
            await runProbe();
            setReadState("ok", "Deleted record Id " + sid + ".");
          } catch (e) {
            setReadState("err", "Delete failed: " + (e.message || String(e)).slice(0, 420));
            log("DELETE failed:\n" + (e.message || String(e)), "err");
          }
        }

        function renderRosterTable(rows, meta, pw, seg) {
          clearRosterTable();
          const thead = document.getElementById("rosterThead");
          const rosterWrap = document.querySelector("#sp-pip-ui .roster-wrap");
          if (!rosterTableBody || !thead) return;
          const plan = resolveRosterColumnPlan(rows);
          const columns = rosterColumnsForDisplay(plan);
          const showRowActions = pw && seg;
          const trHead = document.createElement("tr");
          columns.forEach((col) => {
            const th = document.createElement("th");
            th.textContent = col.label;
            th.title = (col.tryKeys || [col.key]).join(" · ");
            trHead.appendChild(th);
          });
          if (showRowActions) {
            const thx = document.createElement("th");
            thx.className = "roster-actions";
            thx.textContent = " ";
            thx.title = "Record / delete";
            trHead.appendChild(thx);
          }
          thead.appendChild(trHead);

          const frag = document.createDocumentFragment();
          rows.forEach((item) => {
            const tr = document.createElement("tr");
            columns.forEach((col) => {
              const td = document.createElement("td");
              const tryKeys = col.tryKeys || [col.key];
              const raw = valueFromItemByKeys(item, tryKeys);
              const text = raw !== undefined && raw !== null ? formatCellValue(raw) : "";
              td.textContent = text === "" ? "—" : text;
              tr.appendChild(td);
            });
            if (showRowActions && item.Id != null) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              const inner = document.createElement("div");
              inner.className = "roster-actions-inner";
              const recordBtn = document.createElement("button");
              recordBtn.type = "button";
              recordBtn.className = "btn-record";
              recordBtn.textContent = "Record";
              recordBtn.title = "View full record for " + formatPersonDisplayName(item);
              recordBtn.addEventListener("click", () => {
                navigateToPersonDetail(item.Id);
              });
              inner.appendChild(recordBtn);
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              delBtn.title = "Delete list item Id " + item.Id;
              delBtn.addEventListener("click", () => {
                void deletePersonnelRow(item.Id, pw, seg);
              });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            } else if (showRowActions) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              tdAct.textContent = "—";
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          rosterTableBody.appendChild(frag);

          if (rosterWrap) {
            const scrollAfter = Number(ROSTER_SCROLL_AFTER_ROWS) > 0 ? Number(ROSTER_SCROLL_AFTER_ROWS) : 10;
            if (rows.length > scrollAfter) {
              rosterWrap.classList.add("roster-wrap--scroll");
              const headerRow = thead.querySelector("tr");
              const sampleRow = rosterTableBody.querySelector("tr");
              const headerH = headerRow ? headerRow.getBoundingClientRect().height : 42;
              const rowH = sampleRow ? sampleRow.getBoundingClientRect().height : 42;
              rosterWrap.style.maxHeight = Math.ceil(headerH + rowH * scrollAfter + 4) + "px";
            } else {
              rosterWrap.classList.remove("roster-wrap--scroll");
              rosterWrap.style.maxHeight = "";
            }
          }

        }

        function log(msg, cls) {
          if (!out) return;
          const line = typeof msg === "string" ? msg : JSON.stringify(msg, null, 2);
          const span = document.createElement("span");
          if (cls) span.className = cls;
          span.textContent = line + "\n\n";
          out.appendChild(span);
        }

        function escListTitle(title) {
          return String(title || "").replace(/'/g, "''");
        }

        function personnelListTitle() {
          return String(LIST_PERSONNEL || "").trim();
        }

        function personnelGuidRaw() {
          return String(LIST_PERSONNEL_GUID || "").trim();
        }

        function personnelListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(personnelGuidRaw());
        }

        function personnelListApiPath() {
          if (personnelListUsesGuid()) return "lists(guid'" + personnelGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(personnelListTitle()) + "')";
        }

        function getCtx() {
          try {
            if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) return window._spPageContextInfo;
          } catch (_) {}
          if (isSharePointAuthoringContext()) return null;
          try {
            if (window.parent && window.parent._spPageContextInfo && window.parent._spPageContextInfo.webAbsoluteUrl)
              return window.parent._spPageContextInfo;
          } catch (_) {}
          try {
            if (window.top && window.top._spPageContextInfo && window.top._spPageContextInfo.webAbsoluteUrl) return window.top._spPageContextInfo;
          } catch (_) {}
          return null;
        }

        function getDigest() {
          const el = document.getElementById("__REQUESTDIGEST");
          if (el && el.value) return el.value;
          if (isSharePointAuthoringContext()) return "";
          try {
            const parentEl = window.parent && window.parent.document && window.parent.document.getElementById("__REQUESTDIGEST");
            if (parentEl && parentEl.value) return parentEl.value;
          } catch (_) {}
          try {
            const topEl = window.top && window.top.document && window.top.document.getElementById("__REQUESTDIGEST");
            if (topEl && topEl.value) return topEl.value;
          } catch (_) {}
          return "";
        }

        function deriveSiteRootFromSharePointUrl(raw) {
          const s = String(raw || "").trim();
          if (!s) return "";
          try {
            const u = new URL(s);
            const path = u.pathname || "";
            const sites = path.match(/^(\/sites\/[^/]+)/i);
            if (sites) return (u.origin + sites[1]).replace(/\/$/, "");
            const teams = path.match(/^(\/teams\/[^/]+)/i);
            if (teams) return (u.origin + teams[1]).replace(/\/$/, "");
            return u.origin.replace(/\/$/, "");
          } catch (_) {
            return "";
          }
        }

        function resolvedPersonnelRestBase() {
          const site = String(PERSONNEL_SITE_ROOT_URL || "").trim().replace(/\/$/, "");
          if (site) return site;
          const legacy = String(PERSONNEL_REST_WEB_URL || "").trim().replace(/\/$/, "");
          if (legacy) return legacy;
          const derived = deriveSiteRootFromSharePointUrl(PERSONNEL_SHAREPOINT_URL);
          if (derived) return derived;
          const ctx = getCtx();
          if (ctx && ctx.webAbsoluteUrl) return String(ctx.webAbsoluteUrl).replace(/\/$/, "");
          return "";
        }

        /**
         * Fresh X-RequestDigest for a given web URL. Required when POST/MERGE targets a web different from the
         * Script Editor page, or when the hidden __REQUESTDIGEST is stale ("security validation ... invalid").
         */
        async function fetchFormDigestForWeb(webAbsoluteUrl) {
          const base = String(webAbsoluteUrl || "").trim().replace(/\/$/, "");
          if (!base) return "";
          const res = await fetch(base + "/_api/contextinfo", {
            method: "POST",
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
            },
            credentials: "include",
            cache: "no-store",
          });
          const text = await res.text();
          let data = null;
          try {
            data = text ? JSON.parse(text) : null;
          } catch (_) {}
          if (!res.ok || !data) return "";
          if (data.FormDigestValue) return String(data.FormDigestValue);
          if (data.d && data.d.GetContextWebInformation && data.d.GetContextWebInformation.FormDigestValue) {
            return String(data.d.GetContextWebInformation.FormDigestValue);
          }
          return "";
        }

        async function spFetch(path, opts, baseUrlOverride) {
          opts = opts || {};
          const method = (opts.method || "GET").toUpperCase();
          if (isLocalDemoMode()) {
            return demoSpFetch(path, opts, method);
          }
          const base =
            baseUrlOverride != null && String(baseUrlOverride).trim()
              ? String(baseUrlOverride).trim().replace(/\/$/, "")
              : resolvedPersonnelRestBase();
          const headers = Object.assign({}, opts.headers || {});
          headers["Accept"] = "application/json;odata=nometadata";
          const methodUpper = method;
          if (methodUpper !== "GET" && methodUpper !== "DELETE") {
            headers["Content-Type"] = "application/json;odata=nometadata";
          }

          const isPostWithBody = methodUpper === "POST" && opts.body != null;
          const needsWriteDigest =
            !opts.skipAntiForgery && (methodUpper === "MERGE" || methodUpper === "DELETE" || isPostWithBody);

          let digest = "";
          if (needsWriteDigest) {
            try {
              digest = await fetchFormDigestForWeb(base);
            } catch (_) {
              digest = "";
            }
            if (!digest) digest = getDigest();
            if (digest) headers["X-RequestDigest"] = digest;
          }

          let realMethod = methodUpper;
          if (methodUpper === "MERGE" || methodUpper === "DELETE") {
            headers["X-HTTP-Method"] = methodUpper;
            headers["IF-MATCH"] = "*";
            realMethod = "POST";
          }
          const res = await fetch(base + path, {
            method: realMethod,
            headers,
            body: opts.body ? JSON.stringify(opts.body) : undefined,
            credentials: "include",
            cache: "no-store",
          });
          if (res.status === 204) return null;
          const text = await res.text();
          let data = null;
          try {
            data = text ? JSON.parse(text) : null;
          } catch (_) {}
          if (!res.ok) {
            let msg = "";
            if (data && data.error) {
              msg =
                (typeof data.error.message === "string" && data.error.message) ||
                (data.error.message && data.error.message.value) ||
                "";
            }
            if (!msg && data && data["odata.error"] && data["odata.error"].message) {
              msg = data["odata.error"].message.value || data["odata.error"].message || "";
            }
            if (!msg) msg = text || "";
            throw new Error((res.status + " " + res.statusText + " " + msg).slice(0, 1200));
          }
          return data;
        }

        async function enumerateListTitles(pw) {
          log("--- Listing lists on this web (first 200, not hidden) — find exact Title for LIST_PERSONNEL ---");
          const data = await spFetch(
            "/_api/web/lists?$select=Title,Id,ItemCount,Hidden&$filter=Hidden eq false&$orderby=Title&$top=200",
            {},
            pw,
          );
          const rows = (data && data.value) || [];
          rows.forEach((r) => {
            log("  • " + (r.Title || "") + "  (Id=" + r.Id + ", ItemCount=" + (r.ItemCount != null ? r.ItemCount : "?") + ")");
          });
        }

        async function fetchChoiceOptions(seg, pw, fieldInternalOrTitle) {
          const esc = String(fieldInternalOrTitle || "").replace(/'/g, "''");
          if (!esc) return { choices: [], error: "empty field name" };
          const data = await spFetch(
            `/_api/web/${seg}/fields/getbyinternalnameortitle('${esc}')?$select=Choices,Title,InternalName,TypeAsString`,
            {},
            pw,
          );
          let choices = [];
          if (Array.isArray(data.Choices)) choices = data.Choices;
          else if (data.Choices && typeof data.Choices === "object" && Array.isArray(data.Choices.results))
            choices = data.Choices.results;
          choices = choices.map((c) => String(c).trim()).filter(Boolean);
          return { choices: choices, field: data };
        }

        function primeRecordDateDefault() {
          const el = document.getElementById("nf_RecordDate");
          if (el && el.tagName === "INPUT" && el.type === "date" && !el.value) {
            el.value = new Date().toISOString().slice(0, 10);
          }
        }

        function parseCreatedItemId(created) {
          if (!created || typeof created !== "object") return null;
          if (created.Id != null && created.Id !== "") return String(created.Id);
          const oid = created["odata.id"] || created["@odata.id"];
          if (oid && typeof oid === "string") {
            const m = oid.match(/\/items\((\d+)\)/i);
            if (m) return m[1];
          }
          return null;
        }

        function addFormGroupedKeySet() {
          const s = new Set();
          if (!Array.isArray(ADD_FORM_FIELD_GROUPS)) return s;
          ADD_FORM_FIELD_GROUPS.forEach((g) => {
            if (!g || !Array.isArray(g.keys)) return;
            g.keys.forEach((k) => {
              if (k) s.add(String(k));
            });
          });
          return s;
        }

        let pipMeasureCanvas = null;
        function measureCanvasTextWidth(text, font) {
          if (!pipMeasureCanvas) pipMeasureCanvas = document.createElement("canvas");
          const ctx = pipMeasureCanvas.getContext("2d");
          if (!ctx) return (String(text).length + 2) * 8;
          ctx.font = font;
          return Math.ceil(ctx.measureText(text || " ").width);
        }

        function applySelectAutosize(sel) {
          if (!sel || sel.tagName !== "SELECT") return;
          const cs = window.getComputedStyle(sel);
          const font =
            (cs.fontWeight || "400") + " " + (cs.fontSize || "13px") + " " + (cs.fontFamily || "monospace");
          const arrowPad = 40;
          let maxText = 0;
          for (let i = 0; i < sel.options.length; i++) {
            const w = measureCanvasTextWidth(sel.options[i].textContent || "", font);
            if (w > maxText) maxText = w;
          }
          const minW = measureCanvasTextWidth("(select)", font) + arrowPad;
          let px = Math.max(maxText + arrowPad, minW);
          const cap = 400;
          px = Math.min(px, cap);
          sel.style.width = px + "px";
          sel.style.maxWidth = "100%";
        }

        function applyAllSelectAutosizes(root) {
          if (!root || !root.querySelectorAll) return;
          root.querySelectorAll("select.add-field-select--autosize").forEach((el) => applySelectAutosize(el));
        }

        function buildAddFieldWrap(col, sampleRow, opt) {
          opt = opt || {};
          if (col.key === "Id") return null;
          const idPrefix = opt.idPrefix || "nf_";
          const writeKey = resolveWriteKey(col, sampleRow);
          const fwrap = document.createElement("div");
          fwrap.className = "add-field";
          if (col.key === "Address") fwrap.classList.add("add-field--address");
          if (col.key === "Notes") fwrap.classList.add("add-field--notes");
          const lab = document.createElement("label");
          lab.setAttribute("for", idPrefix + col.key);
          lab.textContent = col.label;
          lab.title = "REST write key: " + writeKey;
          fwrap.appendChild(lab);

          const isDrop = DROPDOWN_COLUMN_KEYS.indexOf(col.key) !== -1;
          let input;
          if (isDrop) {
            input = document.createElement("select");
            if (!opt.fullWidthSelect) input.className = "add-field-select--autosize";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
            const opt0 = document.createElement("option");
            opt0.value = "";
            opt0.textContent = "(select)";
            input.appendChild(opt0);
          } else if (col.key === "Address" || col.key === "Notes") {
            input = document.createElement("textarea");
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
            input.rows = col.key === "Address" ? 2 : 2;
          } else if (col.key === "RecordDate") {
            input = document.createElement("input");
            input.type = "date";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
          } else {
            input = document.createElement("input");
            input.type = "text";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
          }
          fwrap.appendChild(input);
          return fwrap;
        }

        function balanceContactNotesToIdentityColumn(form) {
          const split = form.querySelector(".add-form-split--primary");
          if (!split) return;
          const fsets = split.querySelectorAll(":scope > fieldset.add-form-group");
          if (fsets.length < 2) return;
          const leftCol = fsets[0];
          const rightCol = fsets[1];
          const notesTa = form.querySelector(".add-field--notes textarea");
          if (!notesTa || notesTa.tagName !== "TEXTAREA") return;
          notesTa.style.maxHeight = "";
          void split.offsetHeight;
          let leftH = leftCol.getBoundingClientRect().height;
          let rightH = rightCol.getBoundingClientRect().height;
          if (rightH <= leftH + 2) return;
          let guard = 0;
          while (rightH > leftH + 2 && guard < 14) {
            guard++;
            const diff = rightH - leftH;
            const nh = notesTa.getBoundingClientRect().height;
            const next = Math.max(40, nh - diff - 8);
            notesTa.style.maxHeight = next + "px";
            void rightCol.offsetHeight;
            leftH = leftCol.getBoundingClientRect().height;
            rightH = rightCol.getBoundingClientRect().height;
          }
        }

        function buildGroupFieldset(group, cols, sampleRow) {
          if (!group || !group.title || !Array.isArray(group.keys)) return null;

          if (group.layout === "contact-stack") {
            const keys = group.keys;
            if (keys.length < 7) return null;
            const topKeys = keys.slice(0, 3);
            const statusKeys = keys.slice(3, 6);
            const notesKey = keys[6];
            const bundle = document.createElement("div");
            bundle.className = "add-form-contact-bundle";

            const topGrid = document.createElement("div");
            topGrid.className = "add-form-grid add-form-grid--stack add-form-contact-stack-top";
            topKeys.forEach((key) => {
              const col = cols.find((c) => c.key === key);
              if (!col) return;
              const f = buildAddFieldWrap(col, sampleRow);
              if (f) topGrid.appendChild(f);
            });

            const statusRow = document.createElement("div");
            statusRow.className = "add-form-contact-status-row";
            statusKeys.forEach((key) => {
              const col = cols.find((c) => c.key === key);
              if (!col) return;
              const f = buildAddFieldWrap(col, sampleRow, { fullWidthSelect: true });
              if (f) statusRow.appendChild(f);
            });

            const notesCol = cols.find((c) => c.key === notesKey);
            const notesWrap = notesCol ? buildAddFieldWrap(notesCol, sampleRow) : null;

            bundle.appendChild(topGrid);
            bundle.appendChild(statusRow);
            if (notesWrap) bundle.appendChild(notesWrap);

            const fs = document.createElement("fieldset");
            fs.className = "add-form-group";
            if (group.compact) fs.classList.add("add-form-group--compact");
            const leg = document.createElement("legend");
            leg.className = "add-form-group-legend";
            leg.textContent = group.title;
            fs.appendChild(leg);
            fs.appendChild(bundle);
            return fs;
          }

          const grid = document.createElement("div");
          grid.className = "add-form-grid";
          if (group.layout === "stack") grid.classList.add("add-form-grid--stack");
          group.keys.forEach((key) => {
            const col = cols.find((c) => c.key === key);
            if (!col) return;
            const fwrap = buildAddFieldWrap(col, sampleRow);
            if (fwrap) grid.appendChild(fwrap);
          });
          if (!grid.childElementCount) return null;
          const fs = document.createElement("fieldset");
          fs.className = "add-form-group";
          if (group.compact) fs.classList.add("add-form-group--compact");
          const leg = document.createElement("legend");
          leg.className = "add-form-group-legend";
          leg.textContent = group.title;
          fs.appendChild(leg);
          fs.appendChild(grid);
          return fs;
        }

        async function renderAddPersonForm(meta, pw, seg, sampleRow) {
          const panel = document.getElementById("addPersonPanel");
          if (!panel) return;
          panel.innerHTML = "";
          const cols = normalizedRosterColumns();
          if (!cols.length) return;

          const h = document.createElement("h2");
          h.className = "hub-subtitle";
          h.textContent = "Training Hub";
          panel.appendChild(h);

          const form = document.createElement("div");
          form.id = "newPersonForm";
          form.className = "add-form-stack";
          form.setAttribute("role", "form");

          const groupedKeys = addFormGroupedKeySet();

          let splitWrap = null;
          if (Array.isArray(ADD_FORM_FIELD_GROUPS)) {
            ADD_FORM_FIELD_GROUPS.forEach((group) => {
              const fs = buildGroupFieldset(group, cols, sampleRow);
              if (!fs) return;
              if (group.split === "left") {
                splitWrap = document.createElement("div");
                splitWrap.className = "add-form-split add-form-split--primary";
                splitWrap.appendChild(fs);
                form.appendChild(splitWrap);
                return;
              }
              if (group.split === "right") {
                if (!splitWrap) {
                  splitWrap = document.createElement("div");
                  splitWrap.className = "add-form-split";
                  form.appendChild(splitWrap);
                }
                splitWrap.appendChild(fs);
                return;
              }
              form.appendChild(fs);
            });
          }

          const leftovers = cols.filter((c) => c.key !== "Id" && !groupedKeys.has(c.key));
          if (leftovers.length) {
            const grid = document.createElement("div");
            grid.className = "add-form-grid";
            leftovers.forEach((col) => {
              const fwrap = buildAddFieldWrap(col, sampleRow);
              if (fwrap) grid.appendChild(fwrap);
            });
            if (grid.childElementCount) {
              const fs = document.createElement("fieldset");
              fs.className = "add-form-group";
              const leg = document.createElement("legend");
              leg.className = "add-form-group-legend";
              leg.textContent = "Additional fields";
              fs.appendChild(leg);
              fs.appendChild(grid);
              form.appendChild(fs);
            }
          }

          const bar = document.createElement("div");
          bar.className = "add-form-actions";
          const submitBtn = document.createElement("button");
          submitBtn.type = "button";
          submitBtn.className = "btn";
          submitBtn.textContent = "Submit";
          submitBtn.addEventListener("click", function () {
            void submitNewPerson(meta, pw, seg, sampleRow);
          });
          bar.appendChild(submitBtn);
          const clearBtn = document.createElement("button");
          clearBtn.type = "button";
          clearBtn.className = "btn-secondary";
          clearBtn.textContent = "Clear form";
          clearBtn.addEventListener("click", () => {
            resetHubForm(form);
          });
          bar.appendChild(clearBtn);
          form.appendChild(bar);
          panel.appendChild(form);

          for (let i = 0; i < cols.length; i++) {
            const col = cols[i];
            if (col.key === "Id") continue;
            if (DROPDOWN_COLUMN_KEYS.indexOf(col.key) === -1) continue;
            const writeKey = resolveWriteKey(col, sampleRow);
            const sel = document.getElementById("nf_" + col.key);
            if (!sel || sel.tagName !== "SELECT") continue;
            try {
              const r = await fetchChoiceOptions(seg, pw, writeKey);
              if (r.choices && r.choices.length) {
                r.choices.forEach((c) => {
                  const o = document.createElement("option");
                  o.value = c;
                  o.textContent = c;
                  sel.appendChild(o);
                });
              } else {
                const o = document.createElement("option");
                o.value = "";
                o.textContent = "(no choices — check field internal name: " + writeKey + ")";
                o.disabled = true;
                sel.appendChild(o);
              }
            } catch (e) {
              log("Choices for " + writeKey + ": " + (e.message || String(e)), "err");
              const o = document.createElement("option");
              o.value = "";
              o.textContent = "(error loading choices)";
              o.disabled = true;
              sel.appendChild(o);
            }
            applySelectAutosize(sel);
          }
          applyAllSelectAutosizes(form);
          requestAnimationFrame(function () {
            requestAnimationFrame(function () {
              balanceContactNotesToIdentityColumn(form);
            });
          });
          primeRecordDateDefault();
        }

        async function submitNewPerson(meta, pw, seg, sampleRow) {
          /** With `application/json;odata=nometadata`, do not send `__metadata` — SharePoint treats it as a list column and returns 400. */
          const cols = normalizedRosterColumns();
          let lastName = "";
          let firstName = "";
          let dod = "";
          cols.forEach((col) => {
            if (col.key === "Id") return;
            const el = document.getElementById("nf_" + col.key);
            if (!el) return;
            const v = String(el.value || "").trim();
            if (col.key === "LastName") lastName = v;
            if (col.key === "FirstName") firstName = v;
            if (col.key === "DoDID") dod = v;
          });
          if (!lastName || !firstName) {
            setReadState("err", "Last name and First name are required before saving.");
            return;
          }

          const payload = {};
          cols.forEach((col) => {
            if (col.key === "Id") return;
            const el = document.getElementById("nf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWriteKey(col, sampleRow);
            let v = String(el.value || "").trim();
            if (v === "") return;
            if (el.type === "date") {
              payload[writeKey] = v + "T00:00:00Z";
              return;
            }
            payload[writeKey] = v;
          });

          if (SET_TITLE_ON_CREATE) {
            const title = [lastName, firstName].filter(Boolean).join(", ") || dod || "New personnel";
            payload.Title = title;
          }

          try {
            setReadState("loading", "Submitting…");
            const created = await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            const newId = parseCreatedItemId(created);
            log("POST created item Id: " + (newId || "?"), "ok");
            const form = document.getElementById("newPersonForm");
            if (form) resetHubForm(form);
            primeRecordDateDefault();
            await runProbe();
            setReadState("ok", newId ? "Submitted (Id " + newId + "). Table refreshed." : "Submitted. Table refreshed.");
          } catch (e) {
            setReadState("err", "Submit failed: " + (e.message || String(e)).slice(0, 420));
            log("POST failed:\n" + (e.message || String(e)), "err");
          }
        }

        function resetHubForm(root) {
          if (!root) return;
          root.querySelectorAll("input, select, textarea").forEach(function (el) {
            if (el.tagName === "SELECT") el.selectedIndex = 0;
            else el.value = "";
          });
          primeRecordDateDefault();
        }

        async function runProbe() {
          if (isSharePointAuthoringContext()) {
            initSharePointAuthoringSafeMode();
            return;
          }
          try {
          if (out) out.innerHTML = "";
          clearRosterTable();
          updateDemoModeBanner();
          const addPanel = document.getElementById("addPersonPanel");
          if (addPanel) addPanel.innerHTML = "";
          const demo = isLocalDemoMode();
          setReadState("loading", demo ? "Loading demo personnel roster…" : "Loading Personnel list from SharePoint…");

          const ctx = getCtx();
          const pageWeb = ctx && ctx.webAbsoluteUrl ? String(ctx.webAbsoluteUrl).replace(/\/$/, "") : "";
          const pw = resolvedPersonnelRestBase();
          if (!pw) {
            setReadState(
              "err",
              "No REST base URL. Set PERSONNEL_SITE_ROOT_URL (or PERSONNEL_SHAREPOINT_URL), or open this on a SharePoint page.",
            );
            log(
              "FAIL: No REST base URL. Set PERSONNEL_SITE_ROOT_URL to your site root, or set PERSONNEL_SHAREPOINT_URL to any page/list URL under that site, then reload.",
              "err",
            );
            return;
          }
          const seg = personnelListApiPath();

          const baseSource = String(PERSONNEL_SITE_ROOT_URL || "").trim()
            ? "PERSONNEL_SITE_ROOT_URL"
            : String(PERSONNEL_REST_WEB_URL || "").trim()
              ? "PERSONNEL_REST_WEB_URL"
              : String(PERSONNEL_SHAREPOINT_URL || "").trim()
                ? "PERSONNEL_SHAREPOINT_URL (derived site root)"
                : pageWeb
                  ? "_spPageContextInfo.webAbsoluteUrl (page)"
                  : "unknown";
          log("How REST base was chosen: " + baseSource);
          log("location.href:\n" + String(location.href));
          log("Page web from context (informational only):\n" + (pageWeb || "(no _spPageContextInfo — OK if explicit URLs are set)"));
          log(
            "REST base used for ALL API calls below:\n" +
              pw +
              (pageWeb && pw === pageWeb
                ? "  (same as page — from context or your URL matched the page site)"
                : "  (explicit URL / derived from PERSONNEL_SHAREPOINT_URL — not necessarily this page’s site)"),
          );
          log("List API segment:\n/_api/web/" + seg);
          log("LIST_PERSONNEL (trimmed):\n" + JSON.stringify(personnelListTitle()));
          if (demo) {
            log("LOCAL DEMO MODE — in-memory sample data (hostname: " + location.hostname + ")", "ok");
          }

          let meta = null;
          try {
            log("--- Request: GET list metadata ---", "ok");
            meta = await spFetch(`/_api/web/${seg}?$select=Id,Title,ItemCount,ListItemEntityTypeFullName`, {}, pw);
            log(meta);
          } catch (e) {
            log("LIST METADATA FAILED:\n" + (e.message || String(e)), "err");
            setReadState("err", "Cannot read this list (metadata request failed). Check list name and permissions.");
            try {
              await enumerateListTitles(pw);
            } catch (e2) {
              log("Could not enumerate lists either:\n" + (e2.message || String(e2)), "err");
            }
            return;
          }

          const seg2 = personnelListApiPath();
          let items = null;
          const orderByClause = String(PERSONNEL_ITEMS_ORDERBY || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          try {
            log(
              "--- Request: GET items (up to 500" +
                (orderByClause ? ", orderby " + orderByClause : "") +
                ") ---",
              "ok",
            );
            items = await spFetch(`/_api/web/${seg2}/items?$top=500` + orderByQs, {}, pw);
          } catch (e0) {
            if (/\b400\b/.test(String(e0.message || ""))) {
              log("orderby failed with 400; retry without orderby…", "err");
              try {
                items = await spFetch(`/_api/web/${seg2}/items?$top=500`, {}, pw);
              } catch (e1) {
                setReadState("err", "Cannot read list items: " + (e1.message || String(e1)).slice(0, 280));
                log("ITEMS FAILED (after retry):\n" + (e1.message || String(e1)), "err");
                return;
              }
            } else {
              setReadState("err", "Cannot read list items: " + (e0.message || String(e0)).slice(0, 280));
              log("ITEMS FAILED:\n" + (e0.message || String(e0)), "err");
              return;
            }
          }

          let rows = (items && items.value) || [];
          rows = sortPersonnelRowsAlphabetical(rows);
          log("Row count returned: " + rows.length, rows.length ? "ok" : "err");
          const sampleRow = rows.length > 0 ? rows[0] : null;
          const plan = resolveRosterColumnPlan(rows);

          if (rows.length === 0) {
            setReadState(
              "warn",
              "REST succeeded but returned 0 rows in the default view (SharePoint ItemCount=" +
                (meta && meta.ItemCount != null ? meta.ItemCount : "?") +
                "). You can still add a person below; refresh after the first item exists.",
            );
            log("List is empty (ItemCount may still be > 0 if filtered view — this is default view items).");
          } else {
            if (plan.mode === "explicit") {
              const dispCols = rosterColumnsForDisplay(plan);
              log(
                "Roster mode: **explicit** ROSTER_COLUMNS (" +
                  dispCols.length +
                  " displayed):\n" +
                  dispCols.map((c) => c.key + (c.label !== c.key ? "  (header: " + c.label + ")" : "")).join("\n"),
                "ok",
              );
            } else {
              const tableKeys = collectColumnKeys(rows);
              log(
                "Roster mode: **auto** (union of fields on loaded rows). Copy names into ROSTER_COLUMNS to fix layout:\n" +
                  tableKeys.join(", "),
                "ok",
              );
            }
            log("--- First row (raw JSON) ---", "ok");
            log(rows[0]);
            if (rows.length > 1) {
              log("--- Second row (raw JSON) ---", "ok");
              log(rows[1]);
            }
            log(
              plan.mode === "explicit"
                ? "DONE. Table uses your ROSTER_COLUMNS order; wrong cells mean a `key` does not match REST (check spelling)."
                : "DONE. Set ROSTER_COLUMNS in the script to control column order and labels; auto mode uses every field returned.",
              "ok",
            );
          }

          const personIdFromUrl = getPersonIdFromUrl();
          hubSession = {
            rows: rows,
            meta: meta,
            pw: pw,
            seg: seg2,
            sampleRow: sampleRow,
          };

          if (personIdFromUrl) {
            await showPersonDetailById(personIdFromUrl, meta, pw, seg2, rows);
          } else {
            await renderAddPersonForm(meta, pw, seg2, sampleRow);
            renderRosterTable(rows, meta, pw, seg2);
            setHubListViewVisible(true);
          }
          setReadState("", "");
          } finally {
            dismissHubBootOverlay();
          }
        }

        btn.addEventListener("click", () => {
          if (isSharePointAuthoringContext()) {
            setReadState("warn", "Publish the page first, then click Refresh list.");
            return;
          }
          runProbe().catch((e) => {
            setReadState("err", "Unexpected error: " + (e.message || String(e)).slice(0, 280));
            log("UNCAUGHT:\n" + (e.message || String(e)), "err");
          });
        });

        if (recallRosterBtn) {
          recallRosterBtn.addEventListener("click", function () {
            void handleRecallRosterDownload();
          });
        }

        function startHubApp() {
          runProbe().catch((e) => {
            setReadState("err", "Unexpected error: " + (e.message || String(e)).slice(0, 280));
            log("UNCAUGHT:\n" + (e.message || String(e)), "err");
          });
        }

        /** Defer boot so SharePoint edit chrome can mount before we run REST (avoids locking the page editor). */
        const HUB_BOOT_DEFER_MS = 450;

        function bootHub() {
          (function initHubTopBanner() {
            const wrap = document.getElementById("hubTopBannerWrap");
            const img = document.getElementById("hubTopBannerImg");
            const url = resolveTopBannerImageUrl();
            if (!wrap || !img || !url) return;
            img.src = url;
            img.alt =
              typeof HUB_TOP_BANNER_ALT === "string" && HUB_TOP_BANNER_ALT.trim()
                ? HUB_TOP_BANNER_ALT.trim()
                : "Squadron banner";
            wrap.hidden = false;
          })();
          (function initHubLogo() {
            const wrap = document.getElementById("hubLogoWrap");
            const img = document.getElementById("hubLogoImg");
            const url = resolveProfileImageUrl();
            if (!wrap || !img || !url) return;
            img.src = url;
            const altProfile = typeof HUB_PROFILE_ALT === "string" ? HUB_PROFILE_ALT.trim() : "";
            const altLogo = typeof HUB_LOGO_ALT === "string" ? HUB_LOGO_ALT.trim() : "";
            img.alt = altProfile || altLogo || "Profile image";
            wrap.hidden = false;
          })();
          updateDemoModeBanner();
          startHubApp();
        }

        function scheduleBootHub() {
          if (isSharePointAuthoringContext()) {
            enterAuthoringSafeMode();
            updateDemoModeBanner();
            return;
          }
          setTimeout(function () {
            if (isSharePointAuthoringContext()) {
              enterAuthoringSafeMode();
              updateDemoModeBanner();
              return;
            }
            bootHub();
          }, HUB_BOOT_DEFER_MS);
        }

        scheduleBootHub();
      })();
