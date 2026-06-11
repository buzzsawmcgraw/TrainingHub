(function () {
        // ============ EDIT THESE (URLs remove ambiguity) ============
        /**
         * 88thSFS deployment reference:
         * - Site page: https://usaf.dps.mil/sites/88thSFS/SitePages/S3T.aspx
         * - List view: https://usaf.dps.mil/sites/88thSFS/Lists/Personnel.AllItems.aspx
         *
         * (A) Site root - no trailing slash. REST is always: SITE_ROOT + "/_api/web/..."
         */
        const PERSONNEL_SITE_ROOT_URL = "https://usaf.dps.mil/sites/88thSFS";

        /**
         * (B) Optional fallback: paste any URL under the site; derives .../sites/88thSFS. Ignored when (A) is set.
         * Example: https://usaf.dps.mil/sites/88thSFS/Lists/Personnel.AllItems.aspx
         */
        const PERSONNEL_SHAREPOINT_URL = "";

        /** @deprecated Same effect as PERSONNEL_SITE_ROOT_URL; kept for older copies of this file. */
        const PERSONNEL_REST_WEB_URL = "";

        /**
         * Page access gate - users must enter this password before the hub loads. Set "" to disable.
         */
        const HUB_ACCESS_PASSWORD = "Training2026";
        const HUB_ACCESS_STORAGE_KEY = "trainingHubAccessGranted";

        /** Must match Site contents list title (URL .../Lists/Personnel... usually means title "Personnel"). */
        const LIST_PERSONNEL = "Personnel";
        /** Optional: List GUID from List settings -> use if getbytitle fails. */
        const LIST_PERSONNEL_GUID = "";

        /**
         * Weapons certifications list (SharePoint title **WeaponsCertifications**). Rows are filtered by PersonnelId =
         * personnel list item Id. PersonnelId is used for filtering only (not shown in the table on the record page).
         */
        const LIST_WEAPONS_CERTIFICATIONS = "WeaponsCertifications";
        const LIST_WEAPONS_CERTIFICATIONS_GUID = "";
        const WEAPONS_CERT_PERSON_FIELD = "PersonnelId";
        /** Additional REST filter field names to try if the primary name fails (lookup / casing variants). */
        const WEAPONS_CERT_PERSON_FIELD_ALT = [
          "PersonnelID",
          "PersonnelIdId",
          "Personnel_x0020_Id",
          "Personnel_x0020_ID",
          "Personnel/Id",
        ];
        const WEAPONS_CERT_ITEMS_ORDERBY = "QualDate desc";
        const WEAPONS_CERT_COLUMNS = [
          {
            key: "Weapon",
            label: "Weapon",
            altKeys: ["WeaponName", "Weapon_x0020_Name", "WeaponType", "Title"],
          },
          {
            key: "QualDate",
            label: "Certification Date",
            altKeys: ["QualificationDate", "Qualification_x0020_Date", "CertificationDate", "Certification_x0020_Date", "CertDate"],
          },
          {
            key: "ExpirationDate",
            label: "Expiration Date",
            altKeys: ["ExpiryDate", "Expiry_x0020_Date", "ExpiresOn", "Expiration_x0020_Date", "ExpDate"],
          },
          { key: "Status", label: "Status", computed: true },
        ];
        /** Choice/lookup columns on the add-weapons form (Weapon included so list choices load when available). */
        const WEAPONS_CERT_DROPDOWN_KEYS = ["Weapon"];
        /** Set SharePoint Title on new weapons cert rows from the weapon name when true. */
        const WEAPONS_CERT_SET_TITLE = true;

        const LIST_BYLAW_TRAINING = "ByLawTraining";
        const LIST_BYLAW_TRAINING_GUID = "";
        const BYLAW_TRAINING_PERSON_FIELD = "PersonnelId";
        const BYLAW_TRAINING_PERSON_FIELD_ALT = [
          "PersonnelID",
          "PersonnelIdId",
          "Personnel_x0020_Id",
          "Personnel_x0020_ID",
          "Personnel/Id",
        ];
        const BYLAW_TRAINING_ITEMS_ORDERBY = "QualDate desc";
        /** ByLawTraining list: PersonnelId * Item * QualDate * ExpirationDate * Certifier (+ computed Status in hub). */
        const BYLAW_TRAINING_COLUMNS = [
          {
            key: "Item",
            label: "Item",
            altKeys: ["Item0", "Title", "Training", "ByLaw", "TrainingName", "Training_x0020_Name"],
          },
          {
            key: "QualDate",
            label: "Certification Date",
            altKeys: ["QualificationDate", "Qualification_x0020_Date", "CertificationDate", "Certification_x0020_Date", "CertDate"],
          },
          {
            key: "ExpirationDate",
            label: "Expiration Date",
            altKeys: ["ExpiryDate", "Expiry_x0020_Date", "ExpiresOn", "Expiration_x0020_Date", "ExpDate"],
          },
          {
            key: "Certifier",
            label: "Certifier",
            altKeys: ["CertifierId", "Certifier0", "CertifierName", "Certifier_x0020_Name", "Instructor", "Trainer"],
          },
          { key: "Status", label: "Status", computed: true },
        ];
        const BYLAW_TRAINING_DROPDOWN_KEYS = ["Item", "Certifier"];
        /** SharePoint Certifiers list (lookup source for the Certifier field on ByLawTraining). */
        const LIST_CERTIFIERS = "Certifiers";
        const LIST_CERTIFIERS_GUID = "";
        const CERTIFIERS_PERSON_FIELD = "PersonnelId";
        const CERTIFIERS_LIST_DISPLAY_FIELD = "Certifier";
        /** Set SharePoint Title on new rows from the Item value when the list still requires Title. */
        const BYLAW_TRAINING_SET_TITLE = true;
        const BYLAW_TRAINING_ITEM_SORT_KEYS = ["Item", "Item0", "Title", "Training", "ByLaw", "TrainingName", "Training_x0020_Name"];

        /** Hub Reports menu (Reports nav). Set implemented:true when the report renderer is wired. */
        const HUB_REPORT_TYPES = [
          {
            id: "status-of-training",
            title: "Status of Training",
            subtitle: "Squadron-wide snapshot of Weapons Qualifications and By-Law Training.",
            badge: "SOT",
            implemented: true,
            printable: true,
          },
          {
            id: "mql-report",
            title: "MQL Report",
            subtitle: "Minimum Qualification Level listing for review and tracking.",
            badge: "MQL",
            implemented: false,
            placeholder:
              "MQL Report will list minimum qualification levels across the Personnel Roster. This view is queued for a follow-up build.",
          },
          {
            id: "mql-pdf",
            title: "MQL PDF",
            subtitle: "Printable MQL export suitable for signatures and filing.",
            badge: "PDF",
            implemented: false,
            printable: true,
            placeholder:
              "MQL PDF will generate a print-ready minimum qualification document. Use Print report here once the export is wired.",
          },
          {
            id: "heavy-weapons",
            title: "Heavy Weapons Report",
            subtitle: "Heavy weapons qualification status and expirations.",
            badge: "HVY",
            implemented: false,
            placeholder:
              "Heavy Weapons Report will filter Weapons Qualifications to heavy-weapon entries. Define the weapon list in a future hub update.",
          },
          {
            id: "augmentee",
            title: "Augmentee Report",
            subtitle: "Augmentee Personnel training and qualification status.",
            badge: "AUG",
            implemented: false,
            placeholder:
              "Augmentee Report will highlight augmentee Personnel and their training posture. Augmentee identification rules will be added next.",
          },
        ];

        /**
         * Circular profile image to the left of the squadron header (default: **S3T Files** / **Fallout Guy.png**).
         * `HUB_PROFILE_IMAGE_URL` - full URL override. `HUB_LOGO_URL` - legacy override if profile URL is empty.
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
         * address, and set `HUB_TOP_BANNER_URL` to that value (or fix `HUB_BANNER_SITE_RELATIVE_PATH` - e.g. add
         * `Shared%20Documents/` if **S3T Files** is a folder inside Documents, not the library root).
         */
        const HUB_TOP_BANNER_URL = "";
        /** Path from site root to `banner.png` (spaces as %20). */
        const HUB_BANNER_SITE_RELATIVE_PATH = "S3T%20Files/banner.png";
        const HUB_TOP_BANNER_ALT = "88th Security Forces Squadron banner";

        /**
         * Roster column `key` values that render as dropdowns. Options are loaded from SharePoint field **Choices**
         * (same internal name as used for save - see `resolveWriteKey` / first data row).
         */
        const DROPDOWN_COLUMN_KEYS = ["Status", "OfficeSymbol", "Squadron", "Rank"];

        /** Send SharePoint **Title** on create (built from Last + First, else DoD ID, else "New personnel"). */
        const SET_TITLE_ON_CREATE = true;

        /**
         * OData `$orderby` for list items. Rows are always sorted again client-side by LastName, then FirstName, so the
         * roster follows A-Z by name even if this clause is rejected (400).
         */
        const PERSONNEL_ITEMS_ORDERBY = "LastName asc,FirstName asc";

        /** Roster table scrolls after this many rows (0 = no limit). */
        const ROSTER_SCROLL_AFTER_ROWS = 10;

        /**
         * When false, the SharePoint list item `Id` is never shown as a roster column (covers stale Script Editor copies
         * that still list `Id` / "Personnel ID" in ROSTER_COLUMNS). Delete still uses `item.Id`. Set true to allow an Id column.
         */
        const SHOW_LIST_ITEM_ID_IN_ROSTER = false;

        /**
         * Roster columns - fixed order for 88thSFS Personnel. List item Id is hidden unless SHOW_LIST_ITEM_ID_IN_ROSTER is true
         * and you include `{ key: "Id", label: "..." }` here.
         * If a column is always blank, open Diagnostics and change that row's `key` / add `altKeys` (see Record Date).
         */
        const ROSTER_COLUMNS = [
          {
            key: "RecordDate",
            label: "Record Date",
            altKeys: ["Record_x0020_Date", "Record0", "DateArrived", "Date_x0020_Arrived"],
          },
          { key: "DoDID", label: "DoDID" },
          { key: "Rank", label: "Rank", altKeys: ["Rank0", "PayGrade"] },
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
         * Training Hub form (`Id` omitted). Identity | Contact. Contact: Address / phones, then Status * OfficeSymbol *
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

        (function initHubTopBanner() {
          const wrap = document.getElementById("hubTopBannerWrap");
          const img = document.getElementById("hubTopBannerImg");
          const url = resolveTopBannerImageUrl();
          if (!wrap || !img) return;
          if (!url) return;
          img.src = url;
          img.alt =
            typeof HUB_TOP_BANNER_ALT === "string" && HUB_TOP_BANNER_ALT.trim()
              ? HUB_TOP_BANNER_ALT.trim()
              : "Squadron banner";
          wrap.hidden = false;
        })();

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

        const out = document.getElementById("probeOut");
        const btn = document.getElementById("probeRun");
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
        const personWeaponsWrap = document.getElementById("personWeaponsWrap");
        const personWeaponsThead = document.getElementById("personWeaponsThead");
        const personWeaponsBody = document.getElementById("personWeaponsBody");
        const personWeaponsEmpty = document.getElementById("personWeaponsEmpty");
        const personWeaponsState = document.getElementById("personWeaponsState");
        const personWeaponsAddBtn = document.getElementById("personWeaponsAddBtn");
        const personWeaponsAddPanel = document.getElementById("personWeaponsAddPanel");
        const personWeaponsAddForm = document.getElementById("personWeaponsAddForm");
        const personWeaponsAddFields = document.getElementById("personWeaponsAddFields");
        const personWeaponsAddCancelBtn = document.getElementById("personWeaponsAddCancelBtn");
        const personWeaponsSaveBtn = document.getElementById("personWeaponsSaveBtn");
        const personWeaponsFormTitle = document.getElementById("personWeaponsFormTitle");
        const personBylawWrap = document.getElementById("personBylawWrap");
        const personBylawThead = document.getElementById("personBylawThead");
        const personBylawBody = document.getElementById("personBylawBody");
        const personBylawEmpty = document.getElementById("personBylawEmpty");
        const personBylawState = document.getElementById("personBylawState");
        const personBylawAddBtn = document.getElementById("personBylawAddBtn");
        const personBylawAddPanel = document.getElementById("personBylawAddPanel");
        const personBylawAddForm = document.getElementById("personBylawAddForm");
        const personBylawAddFields = document.getElementById("personBylawAddFields");
        const personBylawAddCancelBtn = document.getElementById("personBylawAddCancelBtn");
        const personBylawSaveBtn = document.getElementById("personBylawSaveBtn");
        const personBylawFormTitle = document.getElementById("personBylawFormTitle");
        const hubNavInstructors = document.getElementById("hubNavInstructors");
        const instructorsSection = document.getElementById("instructorsSection");
        const instructorsThead = document.getElementById("instructorsThead");
        const instructorsTableBody = document.getElementById("instructorsTableBody");
        const instructorsEmpty = document.getElementById("instructorsEmpty");
        const instructorsReadState = document.getElementById("instructorsReadState");
        const instructorsBackLink = document.getElementById("instructorsBackLink");
        const instructorsAddBtn = document.getElementById("instructorsAddBtn");
        const instructorsAddPanel = document.getElementById("instructorsAddPanel");
        const instructorsAddForm = document.getElementById("instructorsAddForm");
        const instructorsPersonSelect = document.getElementById("instructorsPersonSelect");
        const instructorsAddCancelBtn = document.getElementById("instructorsAddCancelBtn");
        const instructorsSaveBtn = document.getElementById("instructorsSaveBtn");
        const hubNavReports = document.getElementById("hubNavReports");
        const reportsSection = document.getElementById("reportsSection");
        const reportsReadState = document.getElementById("reportsReadState");
        const reportsBackLink = document.getElementById("reportsBackLink");
        const reportsHubPanel = document.getElementById("reportsHubPanel");
        const reportsHubGrid = document.getElementById("reportsHubGrid");
        const reportsDetailPanel = document.getElementById("reportsDetailPanel");
        const reportsDetailBackLink = document.getElementById("reportsDetailBackLink");
        const reportsPrintBtn = document.getElementById("reportsPrintBtn");
        const reportsDetailTitle = document.getElementById("reportsDetailTitle");
        const reportsDetailSubtitle = document.getElementById("reportsDetailSubtitle");
        const reportsDetailBody = document.getElementById("reportsDetailBody");

        let hubSession = {
          rows: null,
          meta: null,
          pw: null,
          seg: null,
          sampleRow: null,
          weaponsPersonFilterField: null,
          weaponsPersonFilterFields: null,
          weaponsCertRows: null,
          bylawTrainingSampleRow: null,
          bylawTrainingRows: null,
          bylawPersonFilterField: null,
          bylawPersonFilterFields: null,
          bylawPersonPostKey: null,
          certifiersPersonPostKey: undefined,
          certifiersPersonDisplayKey: null,
          certifiersSampleRow: null,
        };

        let instructorsSession = {
          rows: null,
        };

        let reportsSession = {
          activeReportId: null,
          weaponsRows: null,
          bylawRows: null,
        };

        let weaponsCertEditSession = {
          item: null,
        };

        let bylawTrainingEditSession = {
          item: null,
        };

        let personDetailSession = {
          item: null,
          editing: false,
          meta: null,
          pw: null,
          seg: null,
          sampleRow: null,
        };

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

        /** Union of field names across all loaded rows (order: Title, then A-Z). List item `Id` is omitted from the table. */
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

        function formatSharePointLookupDisplay(val) {
          if (val === null || val === undefined || val === "") return "";
          if (typeof val === "object") {
            const label = val.Title != null ? val.Title : val.Name != null ? val.Name : val.Label;
            if (label != null && String(label).trim() !== "") return String(label).trim();
            if (val.Id != null) return String(val.Id);
          }
          return formatCellValue(val);
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

        /** Plain ASCII placeholder for empty table/detail cells (avoids em-dash mojibake in SharePoint). */
        function displayCellText(text) {
          const s = text == null ? "" : String(text);
          return s === "" ? "-" : s;
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
          if (!confirm("Delete Personnel Record Id " + sid + "? This cannot be undone.")) return;
          try {
            setReadState("loading", "Deleting record...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            log("DELETE Id " + sid + " OK", "ok");
            await runProbe();
            setReadState("ok", "Deleted record Id " + sid + ".");
          } catch (e) {
            setReadState("err", "Delete failed: " + (e.message || String(e)).slice(0, 420));
            log("DELETE failed:\n" + (e.message || String(e)), "err");
          }
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

        function navigateToPersonDetail(itemId) {
          void showPersonDetailById(itemId, hubSession.meta, hubSession.pw, hubSession.seg, hubSession.rows);
        }

        async function navigateToRoster() {
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          setInstructorsAddPanelVisible(false);
          showReportsHub();
          personDetailSession = { item: null, editing: false, meta: null, pw: null, seg: null, sampleRow: null };
          clearPersonWeaponsCertSection();
          clearPersonBylawTrainingSection();
          setPersonDetailEditMode(false);
          setHubListViewVisible(true);
          await ensureRosterViewRendered();
        }

        async function navigateToReports() {
          setReportsViewVisible(true);
          if (!hubSession.pw) {
            setReportsState("warn", "Load the Personnel Roster first (wait for list to load or click Refresh list).");
            showReportsHub();
            return;
          }
          showReportsHub();
          setReportsState("", "");
        }

        async function navigateToInstructors() {
          setInstructorsViewVisible(true);
          if (!hubSession.pw || !certifiersListApiSegment()) {
            setInstructorsState(
              "warn",
              !certifiersListApiSegment()
                ? "Certifiers list is not configured."
                : "Load the Personnel Roster first (wait for list to load or click Refresh list).",
            );
            return;
          }
          await loadInstructorsList();
        }

        function normalizedWeaponsCertColumns() {
          const arr = Array.isArray(WEAPONS_CERT_COLUMNS) ? WEAPONS_CERT_COLUMNS : [];
          return arr
            .map(function (entry) {
              if (!entry || !entry.key) return null;
              const key = String(entry.key).trim();
              const label = String(entry.label || key).trim() || key;
              const alt = Array.isArray(entry.altKeys) ? entry.altKeys.map(function (x) { return String(x).trim(); }).filter(Boolean) : [];
              const tryKeys = [key];
              alt.forEach(function (a) {
                if (tryKeys.indexOf(a) === -1) tryKeys.push(a);
              });
              return { key: key, label: label, tryKeys: tryKeys, computed: !!entry.computed };
            })
            .filter(Boolean);
        }

        function weaponsCertFormColumns() {
          return normalizedWeaponsCertColumns().filter(function (col) {
            return !col.computed;
          });
        }

        function weaponsCertExpiryDateKeys() {
          const expiryCol = normalizedWeaponsCertColumns().find(function (col) {
            return col.key === "ExpirationDate" || col.key === "ExpiryDate";
          });
          return expiryCol && expiryCol.tryKeys ? expiryCol.tryKeys.slice() : ["ExpirationDate", "ExpiryDate"];
        }

        const WEAPONS_CERT_MONTH_ABBR = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];

        function parseWeaponsCertCalendarDate(val) {
          if (val === null || val === undefined || val === "") return null;
          const formatted = formatCellValue(val);
          const iso = String(formatted).match(/^(\d{4})-(\d{2})-(\d{2})$/);
          if (iso) {
            const d = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
            return isNaN(d.getTime()) ? null : d;
          }
          const d = new Date(val);
          if (isNaN(d.getTime())) return null;
          return new Date(d.getFullYear(), d.getMonth(), d.getDate());
        }

        function isWeaponsCertDateColumn(col) {
          if (!col) return false;
          return col.key === "QualDate" || col.key === "ExpirationDate" || col.key === "ExpiryDate";
        }

        function isoDateFromCalendarDate(d) {
          const y = d.getFullYear();
          const m = String(d.getMonth() + 1).padStart(2, "0");
          const day = String(d.getDate()).padStart(2, "0");
          return y + "-" + m + "-" + day;
        }

        function isoDateForDateInput(val) {
          const d = parseWeaponsCertCalendarDate(val);
          return d ? isoDateFromCalendarDate(d) : "";
        }

        function formatWeaponsCertDisplayDate(val) {
          if (val === null || val === undefined || val === "") return "";
          const d = parseWeaponsCertCalendarDate(val);
          if (!d) return formatCellValue(val);
          const day = String(d.getDate()).padStart(2, "0");
          const mon = WEAPONS_CERT_MONTH_ABBR[d.getMonth()] || "";
          const yr = String(d.getFullYear()).slice(-2);
          return day + "-" + mon + "-" + yr;
        }

        function expirationDateFromQualDate(qualDate) {
          const year = qualDate.getFullYear() + 1;
          const month = qualDate.getMonth();
          return new Date(year, month + 1, 0);
        }

        function applyCertExpirationFromQual(formIdPrefix) {
          const qualEl = document.getElementById(formIdPrefix + "QualDate");
          const expEl = document.getElementById(formIdPrefix + "ExpirationDate");
          if (!qualEl || !expEl) return;
          const qual = parseWeaponsCertCalendarDate(qualEl.value);
          if (!qual) return;
          expEl.value = isoDateFromCalendarDate(expirationDateFromQualDate(qual));
        }

        function wireCertQualDateAutoExpiry(form, formIdPrefix, wiredAttr) {
          if (!form || form.dataset[wiredAttr] === "1") return;
          form.dataset[wiredAttr] = "1";
          const qualFieldId = formIdPrefix + "QualDate";
          form.addEventListener("change", function (ev) {
            if (ev.target && ev.target.id === qualFieldId) applyCertExpirationFromQual(formIdPrefix);
          });
          form.addEventListener("input", function (ev) {
            if (ev.target && ev.target.id === qualFieldId) applyCertExpirationFromQual(formIdPrefix);
          });
        }

        function applyWeaponsCertExpirationFromQual() {
          applyCertExpirationFromQual("wf_");
        }

        function wireWeaponsCertQualDateAutoExpiry(form) {
          wireCertQualDateAutoExpiry(form, "wf_", "weaponsAutoExpiryWired");
        }

        function applyBylawTrainingExpirationFromQual() {
          applyCertExpirationFromQual("bf_");
        }

        function wireBylawTrainingQualDateAutoExpiry(form) {
          wireCertQualDateAutoExpiry(form, "bf_", "bylawAutoExpiryWired");
        }

        function calendarDaysBetween(fromDate, toDate) {
          const from = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
          const to = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
          return Math.round((to.getTime() - from.getTime()) / 86400000);
        }

        function certQualDateKeysFromColumns(columns) {
          const qualCol = (Array.isArray(columns) ? columns : []).find(function (col) {
            return col && col.key === "QualDate";
          });
          return qualCol && qualCol.tryKeys ? qualCol.tryKeys.slice() : ["QualDate"];
        }

        function resolveCertExpirationDate(item, expiryKeys, qualKeys) {
          let raw = valueFromItemByKeys(item, expiryKeys);
          let expiry = parseWeaponsCertCalendarDate(raw);
          if (expiry) return expiry;
          if (!qualKeys || !qualKeys.length) return null;
          raw = valueFromItemByKeys(item, qualKeys);
          const qual = parseWeaponsCertCalendarDate(raw);
          if (!qual) return null;
          return expirationDateFromQualDate(qual);
        }

        function computeCertStatusFromExpiryKeys(item, expiryKeys, qualKeys) {
          const expiry = resolveCertExpirationDate(item, expiryKeys, qualKeys);
          if (!expiry) return { text: "-", tone: "unknown" };

          const today = new Date();
          const daysLeft = calendarDaysBetween(today, expiry);

          if (daysLeft < 0) return { text: "Expired", tone: "expired" };
          if (daysLeft <= 30) return { text: "Qualified", tone: "urgent" };
          if (daysLeft <= 60) return { text: "Qualified", tone: "warn" };
          return { text: "Qualified", tone: "ok" };
        }

        function computeWeaponsCertStatus(item) {
          const columns = normalizedWeaponsCertColumns();
          return computeCertStatusFromExpiryKeys(item, weaponsCertExpiryDateKeys(), certQualDateKeysFromColumns(columns));
        }

        function normalizedBylawTrainingColumns() {
          const arr = Array.isArray(BYLAW_TRAINING_COLUMNS) ? BYLAW_TRAINING_COLUMNS : [];
          return arr
            .map(function (entry) {
              if (!entry || !entry.key) return null;
              const key = String(entry.key).trim();
              const label = String(entry.label || key).trim() || key;
              const alt = Array.isArray(entry.altKeys) ? entry.altKeys.map(function (x) { return String(x).trim(); }).filter(Boolean) : [];
              const tryKeys = [key];
              alt.forEach(function (a) {
                if (tryKeys.indexOf(a) === -1) tryKeys.push(a);
              });
              return { key: key, label: label, tryKeys: tryKeys, computed: !!entry.computed };
            })
            .filter(Boolean);
        }

        function bylawTrainingFormColumns() {
          return normalizedBylawTrainingColumns().filter(function (col) {
            return !col.computed;
          });
        }

        function bylawTrainingExpiryDateKeys() {
          const expiryCol = normalizedBylawTrainingColumns().find(function (col) {
            return col.key === "ExpirationDate" || col.key === "ExpiryDate";
          });
          return expiryCol && expiryCol.tryKeys ? expiryCol.tryKeys.slice() : ["ExpirationDate", "ExpiryDate"];
        }

        function computeBylawTrainingStatus(item) {
          const columns = normalizedBylawTrainingColumns();
          return computeCertStatusFromExpiryKeys(item, bylawTrainingExpiryDateKeys(), certQualDateKeysFromColumns(columns));
        }

        function weaponsCertListTitle() {
          return String(LIST_WEAPONS_CERTIFICATIONS || "").trim();
        }

        function weaponsCertGuidRaw() {
          return String(LIST_WEAPONS_CERTIFICATIONS_GUID || "").trim();
        }

        function weaponsCertListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(weaponsCertGuidRaw());
        }

        function weaponsCertListApiPath() {
          if (weaponsCertListUsesGuid()) return "lists(guid'" + weaponsCertGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(weaponsCertListTitle()) + "')";
        }

        function setWeaponsCertState(kind, message) {
          if (!personWeaponsState) return;
          if (!message) {
            personWeaponsState.hidden = true;
            personWeaponsState.textContent = "";
            return;
          }
          personWeaponsState.hidden = false;
          personWeaponsState.className = "read-state " + kind;
          personWeaponsState.textContent = message;
        }

        function clearWeaponsCertTable() {
          if (personWeaponsThead) personWeaponsThead.innerHTML = "";
          if (personWeaponsBody) personWeaponsBody.innerHTML = "";
        }

        function clearPersonWeaponsCertSection() {
          clearWeaponsCertTable();
          setWeaponsCertState("", "");
          weaponsCertEditSession.item = null;
          hubSession.weaponsCertRows = null;
          setWeaponsCertAddPanelVisible(false);
          if (personWeaponsAddForm) personWeaponsAddForm.reset();
          if (personWeaponsEmpty) {
            personWeaponsEmpty.hidden = true;
            personWeaponsEmpty.textContent = "No Weapons Qualifications on file for this person.";
          }
          if (personWeaponsWrap) personWeaponsWrap.hidden = true;
        }

        function updateWeaponsCertToolbarLabel() {
          if (!personWeaponsAddBtn) return;
          if (personWeaponsAddPanel && !personWeaponsAddPanel.hidden) {
            personWeaponsAddBtn.textContent = weaponsCertEditSession.item ? "Cancel requalify" : "Cancel add";
          } else {
            personWeaponsAddBtn.textContent = "Add qualification";
          }
        }

        function setWeaponsCertFormMode(mode) {
          const isEdit = mode === "edit";
          if (personWeaponsFormTitle) {
            personWeaponsFormTitle.textContent = isEdit ? "Requalify Weapon" : "New Weapons Qualification";
          }
          if (personWeaponsSaveBtn) {
            personWeaponsSaveBtn.textContent = isEdit ? "Save requalification" : "Save qualification";
          }
          updateWeaponsCertToolbarLabel();
        }

        function setWeaponsCertAddPanelVisible(visible) {
          if (personWeaponsAddPanel) personWeaponsAddPanel.hidden = !visible;
          if (!visible) {
            weaponsCertEditSession.item = null;
            setWeaponsCertFormMode("add");
          } else {
            updateWeaponsCertToolbarLabel();
          }
        }

        function resolveWeaponsCertWriteKey(col, sampleRow) {
          if (col.saveKey) return String(col.saveKey).trim();
          if (sampleRow) {
            const tryKeys = col.tryKeys || [col.key];
            const hit = tryKeys.find(function (k) {
              return sampleRow && Object.prototype.hasOwnProperty.call(sampleRow, k);
            });
            if (hit) return hit;
          }
          return col.key;
        }

        function buildWeaponsCertFieldWrap(col, sampleRow, options) {
          options = options || {};
          const idPrefix = options.idPrefix || "wf_";
          const readOnlyKeys = Array.isArray(options.readOnlyKeys) ? options.readOnlyKeys : [];
          const readOnly = readOnlyKeys.indexOf(col.key) !== -1;
          const writeKey = resolveWeaponsCertWriteKey(col, sampleRow);
          const fwrap = document.createElement("div");
          fwrap.className = "add-field";
          const lab = document.createElement("label");
          lab.setAttribute("for", idPrefix + col.key);
          lab.textContent = col.label;
          lab.title =
            col.key === "ExpirationDate"
              ? "Auto-calculated from certification date (last day of month, one year out). REST write key: " + writeKey
              : "REST write key: " + writeKey;
          fwrap.appendChild(lab);

          const isDrop = WEAPONS_CERT_DROPDOWN_KEYS.indexOf(col.key) !== -1;
          let input;
          if (isDrop) {
            input = document.createElement("select");
            input.className = "add-field-select--autosize";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
            const opt0 = document.createElement("option");
            opt0.value = "";
            opt0.textContent = "(select)";
            input.appendChild(opt0);
          } else if (isWeaponsCertDateColumn(col)) {
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
          if (readOnly) input.disabled = true;
          fwrap.appendChild(input);
          return fwrap;
        }

        function buildWeaponsCertAddFormFields(options) {
          if (!personWeaponsAddFields) return;
          personWeaponsAddFields.innerHTML = "";
          const sampleRow = hubSession.weaponsCertSampleRow;
          weaponsCertFormColumns().forEach(function (col) {
            const f = buildWeaponsCertFieldWrap(col, sampleRow, options);
            if (f) personWeaponsAddFields.appendChild(f);
          });
        }

        function setWeaponsCertFormFieldValue(col, item, el) {
          if (!el || !item) return;
          const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
          if (isWeaponsCertDateColumn(col)) {
            el.value = isoDateForDateInput(raw);
            return;
          }
          if (el.tagName === "SELECT") {
            const writeKey = String(el.dataset.writeKey || "").trim();
            if (/Id$/.test(writeKey) && item[writeKey] != null && item[writeKey] !== "") {
              el.value = String(item[writeKey]);
              return;
            }
            const display = formatCellValue(raw);
            ensureSelectIncludesValue(el, display);
            el.value = display;
            return;
          }
          el.value = formatCellValue(raw);
        }

        function fillWeaponsCertFormFromItem(item) {
          if (!item) return;
          weaponsCertFormColumns().forEach(function (col) {
            const el = document.getElementById("wf_" + col.key);
            if (el) setWeaponsCertFormFieldValue(col, item, el);
          });
        }

        async function populateWeaponsCertAddDropdowns() {
          if (!personWeaponsAddForm || !personDetailSession.pw) return;
          const seg = weaponsCertListApiPath();
          const pw = personDetailSession.pw;
          const sampleRow = hubSession.weaponsCertSampleRow;
          const columns = normalizedWeaponsCertColumns();
          for (let i = 0; i < columns.length; i++) {
            const col = columns[i];
            if (WEAPONS_CERT_DROPDOWN_KEYS.indexOf(col.key) === -1) continue;
            const sel = document.getElementById("wf_" + col.key);
            if (!sel || sel.tagName !== "SELECT") continue;
            await fillDropdownSelect(sel, col, seg, pw, sampleRow, null);
          }
          applyAllSelectAutosizes(personWeaponsAddForm);
        }

        async function resolveWeaponsPersonPostKey(seg, pw) {
          if (hubSession.weaponsPersonPostKey) return hubSession.weaponsPersonPostKey;
          const filterField = String(hubSession.weaponsPersonFilterField || WEAPONS_CERT_PERSON_FIELD || "PersonnelId").trim();
          let postKey = filterField;
          try {
            const esc = filterField.replace(/'/g, "''");
            const data = await spFetch(
              `/_api/web/${seg}/fields/getbyinternalnameortitle('${esc}')?$select=InternalName,TypeAsString`,
              {},
              pw,
            );
            const internal = String(data.InternalName || filterField).trim();
            const type = String(data.TypeAsString || "");
            if (/lookup/i.test(type)) {
              postKey = internal.endsWith("Id") ? internal : internal + "Id";
            } else {
              postKey = internal;
            }
          } catch (_) {
            if (filterField === "PersonnelId" || filterField === "PersonnelID") postKey = "PersonnelIdId";
            else if (!/Id$/.test(filterField)) postKey = filterField + "Id";
          }
          hubSession.weaponsPersonPostKey = postKey;
          return postKey;
        }

        async function openWeaponsCertAddPanel() {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setWeaponsCertState("warn", "Open a Personnel Record before adding a qualification.");
            return;
          }
          weaponsCertEditSession.item = null;
          setWeaponsCertFormMode("add");
          buildWeaponsCertAddFormFields();
          setWeaponsCertAddPanelVisible(true);
          setWeaponsCertState("", "");
          await populateWeaponsCertAddDropdowns();
          wireWeaponsCertQualDateAutoExpiry(personWeaponsAddForm);
          const weaponEl = document.getElementById("wf_Weapon");
          if (weaponEl) weaponEl.focus();
        }

        async function openWeaponsCertEditPanel(item) {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setWeaponsCertState("warn", "Open a Personnel Record before requalifying.");
            return;
          }
          if (!item || item.Id == null) return;
          weaponsCertEditSession.item = item;
          setWeaponsCertFormMode("edit");
          buildWeaponsCertAddFormFields({ readOnlyKeys: ["Weapon"] });
          setWeaponsCertAddPanelVisible(true);
          setWeaponsCertState("", "");
          await populateWeaponsCertAddDropdowns();
          fillWeaponsCertFormFromItem(item);
          const qualEl = document.getElementById("wf_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyWeaponsCertExpirationFromQual();
          wireWeaponsCertQualDateAutoExpiry(personWeaponsAddForm);
          if (qualEl) qualEl.focus();
        }

        async function deleteWeaponsCertRow(id, pw, seg, personId) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid)) {
            setWeaponsCertState("err", "Invalid qualification Id for delete.");
            return;
          }
          if (!confirm("Delete Weapons Qualification Id " + sid + "? This cannot be undone.")) return;
          try {
            setWeaponsCertState("loading", "Deleting qualification...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            if (weaponsCertEditSession.item && parseInt(String(weaponsCertEditSession.item.Id), 10) === sid) {
              weaponsCertEditSession.item = null;
              setWeaponsCertAddPanelVisible(false);
              if (personWeaponsAddForm) personWeaponsAddForm.reset();
            }
            if (personId) await loadPersonWeaponsCertifications(personId, pw);
            setWeaponsCertState("ok", "Qualification deleted.");
            window.setTimeout(function () {
              setWeaponsCertState("", "");
            }, 2500);
          } catch (e) {
            setWeaponsCertState("err", "Delete failed: " + (e.message || String(e)).slice(0, 280));
            log("Weapons cert DELETE failed:\n" + (e.message || String(e)), "err");
          }
        }

        async function submitWeaponsCertSave() {
          const personItem = personDetailSession.item;
          const pw = personDetailSession.pw;
          if (!personItem || personItem.Id == null || !pw) return;

          const editItem = weaponsCertEditSession.item;
          const isEdit = editItem && editItem.Id != null;
          const seg = weaponsCertListApiPath();
          const sampleRow = hubSession.weaponsCertSampleRow;
          const weaponEl = document.getElementById("wf_Weapon");
          const weaponVal = weaponEl ? String(weaponEl.value || "").trim() : "";
          if (!isEdit && !weaponVal) {
            setWeaponsCertState("err", "Weapon is required.");
            return;
          }

          const qualEl = document.getElementById("wf_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) {
            setWeaponsCertState("err", "Certification date is required.");
            return;
          }

          const payload = {};
          weaponsCertFormColumns().forEach(function (col) {
            if (isEdit && col.key === "Weapon") return;
            const el = document.getElementById("wf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWeaponsCertWriteKey(col, sampleRow);
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });

          if (!isEdit) {
            const personId = parseInt(String(personItem.Id), 10);
            if (!personId || isNaN(personId)) {
              setWeaponsCertState("err", "Invalid Personnel Record Id.");
              return;
            }
            const personKey = await resolveWeaponsPersonPostKey(seg, pw);
            payload[personKey] = personId;
            if (WEAPONS_CERT_SET_TITLE) payload.Title = weaponVal;
          }

          try {
            setWeaponsCertState("loading", isEdit ? "Saving requalification..." : "Saving qualification...");
            if (personWeaponsSaveBtn) personWeaponsSaveBtn.disabled = true;
            if (isEdit) {
              await spFetch(`/_api/web/${seg}/items(${editItem.Id})`, { method: "MERGE", body: payload }, pw);
            } else {
              await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            }
            weaponsCertEditSession.item = null;
            if (personWeaponsAddForm) personWeaponsAddForm.reset();
            setWeaponsCertAddPanelVisible(false);
            await loadPersonWeaponsCertifications(personItem.Id, pw);
            setWeaponsCertState("ok", isEdit ? "Requalification saved." : "Qualification saved.");
            window.setTimeout(function () {
              setWeaponsCertState("", "");
            }, 2500);
          } catch (e) {
            setWeaponsCertState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
            log("Weapons cert save failed:\n" + (e.message || String(e)), "err");
          } finally {
            if (personWeaponsSaveBtn) personWeaponsSaveBtn.disabled = false;
          }
        }

        function renderWeaponsCertTable(rows) {
          clearWeaponsCertTable();
          const columns = normalizedWeaponsCertColumns();
          if (!personWeaponsThead || !personWeaponsBody || !columns.length) return;

          const showActions = !!(personDetailSession.pw && personWeaponsWrap && !personWeaponsWrap.hidden);
          const trHead = document.createElement("tr");
          columns.forEach(function (col) {
            const th = document.createElement("th");
            th.textContent = col.label;
            trHead.appendChild(th);
          });
          if (showActions) {
            const thAct = document.createElement("th");
            thAct.className = "roster-actions";
            thAct.textContent = " ";
            thAct.title = "Requalify / delete";
            trHead.appendChild(thAct);
          }
          personWeaponsThead.appendChild(trHead);

          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              if (col.computed && col.key === "Status") {
                const status = computeWeaponsCertStatus(item);
                td.textContent = displayCellText(status.text);
                if (status.tone && status.tone !== "unknown") {
                  td.className = "weapons-status weapons-status--" + status.tone;
                }
              } else if (isWeaponsCertDateColumn(col)) {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                td.textContent = displayCellText(formatWeaponsCertDisplayDate(raw));
              } else {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                const text = raw !== undefined && raw !== null ? formatCellValue(raw) : "";
                td.textContent = displayCellText(text);
              }
              tr.appendChild(td);
            });
            if (showActions && item.Id != null) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              const inner = document.createElement("div");
              inner.className = "roster-actions-inner";
              const requalBtn = document.createElement("button");
              requalBtn.type = "button";
              requalBtn.className = "btn-record";
              requalBtn.textContent = "Requalify";
              const weaponLabel = formatCellValue(
                valueFromItemByKeys(item, ["Weapon", "WeaponName", "Weapon_x0020_Name", "Title"]) || "",
              );
              requalBtn.title = weaponLabel ? "Requalify " + weaponLabel : "Requalify weapon";
              requalBtn.addEventListener("click", function () {
                void openWeaponsCertEditPanel(item);
              });
              inner.appendChild(requalBtn);
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              delBtn.title = weaponLabel ? "Delete " + weaponLabel + " qualification" : "Delete qualification";
              delBtn.addEventListener("click", function () {
                const personId = personDetailSession.item && personDetailSession.item.Id;
                void deleteWeaponsCertRow(item.Id, personDetailSession.pw, weaponsCertListApiPath(), personId);
              });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            } else if (showActions) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              tdAct.textContent = "-";
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          personWeaponsBody.appendChild(frag);

          if (personWeaponsEmpty) {
            personWeaponsEmpty.hidden = rows.length > 0;
          }
        }

        function weaponsPersonFilterFieldCandidates() {
          const primary = String(WEAPONS_CERT_PERSON_FIELD || "PersonnelId").trim();
          const out = [];
          if (primary) out.push(primary);
          const alts = Array.isArray(WEAPONS_CERT_PERSON_FIELD_ALT) ? WEAPONS_CERT_PERSON_FIELD_ALT : [];
          alts.forEach(function (name) {
            const n = String(name || "").trim();
            if (n && out.indexOf(n) === -1) out.push(n);
          });
          return out;
        }

        function weaponsPersonnelIdFromItem(item) {
          if (!item || typeof item !== "object") return null;
          const keys = weaponsPersonFilterFieldCandidates().concat([
            "PersonnelId",
            "PersonnelID",
            "PersonnelIdId",
            "Personnel_x0020_Id",
          ]);
          for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (!k || k.indexOf("/") !== -1) continue;
            if (Object.prototype.hasOwnProperty.call(item, k) && item[k] !== null && item[k] !== "") {
              return String(item[k]);
            }
          }
          if (item.Personnel && typeof item.Personnel === "object" && item.Personnel.Id != null) {
            return String(item.Personnel.Id);
          }
          return null;
        }

        function weaponsRowsMatchPersonnel(rows, personnelId) {
          const want = String(personnelId);
          return (Array.isArray(rows) ? rows : []).filter(function (item) {
            return weaponsPersonnelIdFromItem(item) === want;
          });
        }

        async function discoverWeaponsPersonFilterFields(seg, pw) {
          if (hubSession.weaponsPersonFilterFields && hubSession.weaponsPersonFilterFields.length) {
            return hubSession.weaponsPersonFilterFields.slice();
          }
          const out = [];
          try {
            const data = await spFetch(
              `/_api/web/${seg}/fields?$select=InternalName,Title,EntityPropertyName,StaticName,TypeAsString&$filter=Hidden eq false&$top=200`,
              {},
              pw,
            );
            const fields = (data && data.value) || [];
            const want = String(WEAPONS_CERT_PERSON_FIELD || "PersonnelId").trim().toLowerCase();
            let hit = fields.find(function (f) {
              const title = String(f.Title || "").trim().toLowerCase();
              const internal = String(f.InternalName || f.StaticName || f.EntityPropertyName || "").trim();
              const internalLower = internal.toLowerCase();
              return (
                internalLower === want ||
                title === want ||
                title === "personnel id" ||
                internalLower === "personnelid" ||
                internalLower === "personnelidid"
              );
            });
            if (!hit) {
              hit = fields.find(function (f) {
                const internal = String(f.InternalName || f.StaticName || "").toLowerCase();
                return internal.indexOf("personnel") !== -1 && internal.indexOf("id") !== -1;
              });
            }
            if (hit) {
              const internal = String(hit.InternalName || hit.StaticName || hit.EntityPropertyName || "").trim();
              const type = String(hit.TypeAsString || "");
              if (internal) out.push(internal);
              if (/lookup/i.test(type) && internal) {
                const lookupIdField = internal.endsWith("Id") ? internal + "Id" : internal + "Id";
                if (out.indexOf(lookupIdField) === -1) out.push(lookupIdField);
                const slashField = internal + "/Id";
                if (out.indexOf(slashField) === -1) out.push(slashField);
              }
            }
          } catch (_) {}
          hubSession.weaponsPersonFilterFields = out.slice();
          return out;
        }

        async function fetchWeaponsCertificationsForPersonnel(seg, pw, personnelId) {
          const id = parseInt(String(personnelId), 10);
          if (!id || isNaN(id)) return [];
          const orderByClause = String(WEAPONS_CERT_ITEMS_ORDERBY || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          const tried = [];
          const candidates = weaponsPersonFilterFieldCandidates().slice();

          const discovered = await discoverWeaponsPersonFilterFields(seg, pw);
          discovered.forEach(function (name) {
            if (name && candidates.indexOf(name) === -1) candidates.unshift(name);
          });
          if (discovered.length) {
            discovered.slice().reverse().forEach(function (name) {
              if (name && candidates.indexOf(name) !== -1) {
                candidates.splice(candidates.indexOf(name), 1);
                candidates.unshift(name);
              }
            });
          }

          for (let i = 0; i < candidates.length; i++) {
            const personField = candidates[i];
            if (!personField || tried.indexOf(personField) !== -1) continue;
            tried.push(personField);
            const filterRaw = personField + " eq " + id;
            const filter = encodeURIComponent(filterRaw);
            try {
              let data = null;
              try {
                data = await spFetch(`/_api/web/${seg}/items?$top=200&$filter=${filter}` + orderByQs, {}, pw);
              } catch (e0) {
                if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
                  data = await spFetch(`/_api/web/${seg}/items?$top=200&$filter=${filter}`, {}, pw);
                } else {
                  throw e0;
                }
              }
              hubSession.weaponsPersonFilterField = personField;
              return (data && data.value) || [];
            } catch (e1) {
              if (!/\b400\b/.test(String(e1.message || "")) && !/does not exist/i.test(String(e1.message || ""))) {
                throw e1;
              }
            }
          }

          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=500` + orderByQs, {}, pw);
          } catch (e2) {
            if (/\b400\b/.test(String(e2.message || "")) && orderByQs) {
              data = await spFetch(`/_api/web/${seg}/items?$top=500`, {}, pw);
            } else {
              throw e2;
            }
          }
          return weaponsRowsMatchPersonnel((data && data.value) || [], id);
        }

        async function loadPersonWeaponsCertifications(personId, pw) {
          if (!personId) return;
          if (personWeaponsWrap) personWeaponsWrap.hidden = false;
          setWeaponsCertState("loading", "Loading Weapons Qualifications...");
          clearWeaponsCertTable();
          if (personWeaponsEmpty) personWeaponsEmpty.hidden = true;

          const seg = weaponsCertListApiPath();
          if (!weaponsCertListTitle() && !weaponsCertListUsesGuid()) {
            renderWeaponsCertTable([]);
            setWeaponsCertState("", "");
            if (personWeaponsEmpty) {
              personWeaponsEmpty.hidden = false;
              personWeaponsEmpty.textContent = "Weapons certifications list is not configured.";
            }
            return;
          }

          try {
            const rows = await fetchWeaponsCertificationsForPersonnel(seg, pw, personId);
            const dateKeys = [
              "QualDate",
              "QualificationDate",
              "Qualification_x0020_Date",
              "CertificationDate",
              "Certification_x0020_Date",
              "CertDate",
            ];
            rows.sort(function (a, b) {
              const da = formatCellValue(valueFromItemByKeys(a, dateKeys) || "");
              const db = formatCellValue(valueFromItemByKeys(b, dateKeys) || "");
              if (da !== db) return String(db).localeCompare(String(da));
              const wa = formatCellValue(valueFromItemByKeys(a, ["Weapon", "WeaponName", "Title"]) || "");
              const wb = formatCellValue(valueFromItemByKeys(b, ["Weapon", "WeaponName", "Title"]) || "");
              return String(wa).localeCompare(String(wb));
            });
            hubSession.weaponsCertSampleRow = rows.length ? rows[0] : hubSession.weaponsCertSampleRow || null;
            hubSession.weaponsCertRows = rows.slice();
            renderWeaponsCertTable(rows);
            setWeaponsCertState("", "");
          } catch (e) {
            hubSession.weaponsCertSampleRow = hubSession.weaponsCertSampleRow || null;
            renderWeaponsCertTable([]);
            setWeaponsCertState("warn", "Could not load Weapons Qualifications: " + (e.message || String(e)).slice(0, 220));
            if (personWeaponsEmpty) {
              personWeaponsEmpty.hidden = false;
              personWeaponsEmpty.textContent =
                "No Weapons Qualifications loaded. If PersonnelId exists on the list, check that values match the Personnel Record Id.";
            }
          }
        }


        function bylawTrainingListTitle() {
          return String(LIST_BYLAW_TRAINING || "").trim();
        }

        function bylawTrainingGuidRaw() {
          return String(LIST_BYLAW_TRAINING_GUID || "").trim();
        }

        function bylawTrainingListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(bylawTrainingGuidRaw());
        }

        function bylawTrainingListApiPath() {
          if (bylawTrainingListUsesGuid()) return "lists(guid'" + bylawTrainingGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(bylawTrainingListTitle()) + "')";
        }

        function setBylawTrainingState(kind, message) {
          if (!personBylawState) return;
          if (!message) {
            personBylawState.hidden = true;
            personBylawState.textContent = "";
            return;
          }
          personBylawState.hidden = false;
          personBylawState.className = "read-state " + kind;
          personBylawState.textContent = message;
        }

        function clearBylawTrainingTable() {
          if (personBylawThead) personBylawThead.innerHTML = "";
          if (personBylawBody) personBylawBody.innerHTML = "";
        }

        function clearPersonBylawTrainingSection() {
          clearBylawTrainingTable();
          setBylawTrainingState("", "");
          bylawTrainingEditSession.item = null;
          hubSession.bylawTrainingRows = null;
          setBylawTrainingAddPanelVisible(false);
          if (personBylawAddForm) personBylawAddForm.reset();
          if (personBylawEmpty) {
            personBylawEmpty.hidden = true;
            personBylawEmpty.textContent = "No By-Law Training records on file for this person.";
          }
          if (personBylawWrap) personBylawWrap.hidden = true;
        }

        function updateBylawTrainingToolbarLabel() {
          if (!personBylawAddBtn) return;
          if (personBylawAddPanel && !personBylawAddPanel.hidden) {
            personBylawAddBtn.textContent = bylawTrainingEditSession.item ? "Cancel requalify" : "Cancel add";
          } else {
            personBylawAddBtn.textContent = "Add Training";
          }
        }

        function setBylawTrainingFormMode(mode) {
          const isEdit = mode === "edit";
          if (personBylawFormTitle) {
            personBylawFormTitle.textContent = isEdit ? "Requalify Training" : "New By-Law Training Record";
          }
          if (personBylawSaveBtn) {
            personBylawSaveBtn.textContent = isEdit ? "Save Requalification" : "Save Training";
          }
          updateBylawTrainingToolbarLabel();
        }

        function setBylawTrainingAddPanelVisible(visible) {
          if (personBylawAddPanel) personBylawAddPanel.hidden = !visible;
          if (!visible) {
            bylawTrainingEditSession.item = null;
            setBylawTrainingFormMode("add");
          } else {
            updateBylawTrainingToolbarLabel();
          }
        }

        function resolveBylawTrainingWriteKey(col, sampleRow) {
          if (col.saveKey) return String(col.saveKey).trim();
          if (sampleRow) {
            const tryKeys = col.tryKeys || [col.key];
            const hit = tryKeys.find(function (k) {
              return sampleRow && Object.prototype.hasOwnProperty.call(sampleRow, k);
            });
            if (hit) return hit;
          }
          return col.key;
        }

        function buildBylawTrainingFieldWrap(col, sampleRow, options) {
          options = options || {};
          const idPrefix = options.idPrefix || "bf_";
          const readOnlyKeys = Array.isArray(options.readOnlyKeys) ? options.readOnlyKeys : [];
          const readOnly = readOnlyKeys.indexOf(col.key) !== -1;
          const writeKey = resolveBylawTrainingWriteKey(col, sampleRow);
          const fwrap = document.createElement("div");
          fwrap.className = "add-field";
          const lab = document.createElement("label");
          lab.setAttribute("for", idPrefix + col.key);
          lab.textContent = col.label;
          lab.title =
            col.key === "ExpirationDate"
              ? "Auto-calculated from certification date (last day of month, one year out). REST write key: " + writeKey
              : "REST write key: " + writeKey;
          fwrap.appendChild(lab);

          const isDrop = BYLAW_TRAINING_DROPDOWN_KEYS.indexOf(col.key) !== -1;
          let input;
          if (isDrop) {
            input = document.createElement("select");
            input.className = "add-field-select--autosize";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
            const opt0 = document.createElement("option");
            opt0.value = "";
            opt0.textContent = "(select)";
            input.appendChild(opt0);
          } else if (isWeaponsCertDateColumn(col)) {
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
          if (readOnly) input.disabled = true;
          fwrap.appendChild(input);
          return fwrap;
        }

        function buildBylawTrainingAddFormFields(options) {
          if (!personBylawAddFields) return;
          personBylawAddFields.innerHTML = "";
          const sampleRow = hubSession.bylawTrainingSampleRow;
          bylawTrainingFormColumns().forEach(function (col) {
            const f = buildBylawTrainingFieldWrap(col, sampleRow, options);
            if (f) personBylawAddFields.appendChild(f);
          });
        }

        function setBylawTrainingFormFieldValue(col, item, el) {
          if (!el || !item) return;
          const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
          if (isWeaponsCertDateColumn(col)) {
            el.value = isoDateForDateInput(raw);
            return;
          }
          if (el.tagName === "SELECT") {
            const writeKey = String(el.dataset.writeKey || "").trim();
            if (/Id$/.test(writeKey) && item[writeKey] != null && item[writeKey] !== "") {
              el.value = String(item[writeKey]);
              return;
            }
            const lookupBase = writeKey.replace(/Id$/, "");
            if (lookupBase && item[lookupBase] && typeof item[lookupBase] === "object" && item[lookupBase].Id != null) {
              el.value = String(item[lookupBase].Id);
              return;
            }
            const display = formatSharePointLookupDisplay(raw);
            ensureSelectIncludesValue(el, display);
            el.value = display;
            return;
          }
          el.value = formatSharePointLookupDisplay(raw);
        }

        function fillBylawTrainingFormFromItem(item) {
          if (!item) return;
          bylawTrainingFormColumns().forEach(function (col) {
            const el = document.getElementById("bf_" + col.key);
            if (el) setBylawTrainingFormFieldValue(col, item, el);
          });
        }

        async function populateBylawTrainingAddDropdowns() {
          if (!personBylawAddForm || !personDetailSession.pw) return;
          const seg = bylawTrainingListApiPath();
          const pw = personDetailSession.pw;
          const sampleRow = hubSession.bylawTrainingSampleRow;
          const columns = normalizedBylawTrainingColumns();
          for (let i = 0; i < columns.length; i++) {
            const col = columns[i];
            if (BYLAW_TRAINING_DROPDOWN_KEYS.indexOf(col.key) === -1) continue;
            const sel = document.getElementById("bf_" + col.key);
            if (!sel || sel.tagName !== "SELECT") continue;
            if (col.key === "Certifier") {
              await fillBylawCertifierDropdown(sel, seg, pw, sampleRow);
            } else {
              await fillDropdownSelect(sel, col, seg, pw, sampleRow, null);
            }
          }
          applyAllSelectAutosizes(personBylawAddForm);
        }

        async function resolveBylawPersonPostKey(seg, pw) {
          if (hubSession.bylawPersonPostKey) return hubSession.bylawPersonPostKey;
          const filterField = String(hubSession.bylawPersonFilterField || BYLAW_TRAINING_PERSON_FIELD || "PersonnelId").trim();
          let postKey = filterField;
          try {
            const esc = filterField.replace(/'/g, "''");
            const data = await spFetch(
              `/_api/web/${seg}/fields/getbyinternalnameortitle('${esc}')?$select=InternalName,TypeAsString`,
              {},
              pw,
            );
            const internal = String(data.InternalName || filterField).trim();
            const type = String(data.TypeAsString || "");
            if (/lookup/i.test(type)) {
              postKey = internal.endsWith("Id") ? internal : internal + "Id";
            } else {
              postKey = internal;
            }
          } catch (_) {
            if (filterField === "PersonnelId" || filterField === "PersonnelID") postKey = "PersonnelIdId";
            else if (!/Id$/.test(filterField)) postKey = filterField + "Id";
          }
          hubSession.bylawPersonPostKey = postKey;
          return postKey;
        }

        async function openBylawTrainingAddPanel() {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setBylawTrainingState("warn", "Open a Personnel Record before adding training.");
            return;
          }
          bylawTrainingEditSession.item = null;
          setBylawTrainingFormMode("add");
          buildBylawTrainingAddFormFields();
          setBylawTrainingAddPanelVisible(true);
          setBylawTrainingState("", "");
          await populateBylawTrainingAddDropdowns();
          wireBylawTrainingQualDateAutoExpiry(personBylawAddForm);
          const itemEl = document.getElementById("bf_Item");
          if (itemEl) itemEl.focus();
        }

        async function openBylawTrainingEditPanel(item) {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setBylawTrainingState("warn", "Open a Personnel Record before requalifying.");
            return;
          }
          if (!item || item.Id == null) return;
          bylawTrainingEditSession.item = item;
          setBylawTrainingFormMode("edit");
          buildBylawTrainingAddFormFields({ readOnlyKeys: ["Item"] });
          setBylawTrainingAddPanelVisible(true);
          setBylawTrainingState("", "");
          await populateBylawTrainingAddDropdowns();
          fillBylawTrainingFormFromItem(item);
          const qualEl = document.getElementById("bf_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyBylawTrainingExpirationFromQual();
          wireBylawTrainingQualDateAutoExpiry(personBylawAddForm);
          if (qualEl) qualEl.focus();
        }

        async function deleteBylawTrainingRow(id, pw, seg, personId) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid)) {
            setBylawTrainingState("err", "Invalid training record Id for delete.");
            return;
          }
          if (!confirm("Delete By-Law Training record Id " + sid + "? This cannot be undone.")) return;
          try {
            setBylawTrainingState("loading", "Deleting training record...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            if (bylawTrainingEditSession.item && parseInt(String(bylawTrainingEditSession.item.Id), 10) === sid) {
              bylawTrainingEditSession.item = null;
              setBylawTrainingAddPanelVisible(false);
              if (personBylawAddForm) personBylawAddForm.reset();
            }
            if (personId) await loadPersonBylawTraining(personId, pw);
            setBylawTrainingState("ok", "Training record deleted.");
            window.setTimeout(function () {
              setBylawTrainingState("", "");
            }, 2500);
          } catch (e) {
            setBylawTrainingState("err", "Delete failed: " + (e.message || String(e)).slice(0, 280));
            log("By-Law Training DELETE failed:\n" + (e.message || String(e)), "err");
          }
        }

        async function submitBylawTrainingSave() {
          const personItem = personDetailSession.item;
          const pw = personDetailSession.pw;
          if (!personItem || personItem.Id == null || !pw) return;

          const editItem = bylawTrainingEditSession.item;
          const isEdit = editItem && editItem.Id != null;
          const seg = bylawTrainingListApiPath();
          const sampleRow = hubSession.bylawTrainingSampleRow;
          const itemEl = document.getElementById("bf_Item");
          const itemVal = itemEl ? String(itemEl.value || "").trim() : "";
          if (!isEdit && !itemVal) {
            setBylawTrainingState("err", "Item is required.");
            return;
          }

          const qualEl = document.getElementById("bf_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) {
            setBylawTrainingState("err", "Certification date is required.");
            return;
          }

          const payload = {};
          bylawTrainingFormColumns().forEach(function (col) {
            if (isEdit && col.key === "Item") return;
            const el = document.getElementById("bf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveBylawTrainingWriteKey(col, sampleRow);
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });

          if (!isEdit) {
            const personId = parseInt(String(personItem.Id), 10);
            if (!personId || isNaN(personId)) {
              setBylawTrainingState("err", "Invalid Personnel Record Id.");
              return;
            }
            const personKey = await resolveBylawPersonPostKey(seg, pw);
            payload[personKey] = personId;
            if (BYLAW_TRAINING_SET_TITLE) payload.Title = itemVal;
          }

          try {
            setBylawTrainingState("loading", isEdit ? "Saving requalification..." : "Saving training record...");
            if (personBylawSaveBtn) personBylawSaveBtn.disabled = true;
            if (isEdit) {
              await spFetch(`/_api/web/${seg}/items(${editItem.Id})`, { method: "MERGE", body: payload }, pw);
            } else {
              await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            }
            bylawTrainingEditSession.item = null;
            if (personBylawAddForm) personBylawAddForm.reset();
            setBylawTrainingAddPanelVisible(false);
            await loadPersonBylawTraining(personItem.Id, pw);
            setBylawTrainingState("ok", isEdit ? "Requalification saved." : "Training record saved.");
            window.setTimeout(function () {
              setBylawTrainingState("", "");
            }, 2500);
          } catch (e) {
            setBylawTrainingState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
            log("By-Law Training save failed:\n" + (e.message || String(e)), "err");
          } finally {
            if (personBylawSaveBtn) personBylawSaveBtn.disabled = false;
          }
        }

        function renderBylawTrainingTable(rows) {
          clearBylawTrainingTable();
          const columns = normalizedBylawTrainingColumns();
          if (!personBylawThead || !personBylawBody || !columns.length) return;

          const showActions = !!(personDetailSession.pw && personBylawWrap && !personBylawWrap.hidden);
          const trHead = document.createElement("tr");
          columns.forEach(function (col) {
            const th = document.createElement("th");
            th.textContent = col.label;
            trHead.appendChild(th);
          });
          if (showActions) {
            const thAct = document.createElement("th");
            thAct.className = "roster-actions";
            thAct.textContent = " ";
            thAct.title = "Requalify / delete";
            trHead.appendChild(thAct);
          }
          personBylawThead.appendChild(trHead);

          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              if (col.computed && col.key === "Status") {
                const status = computeBylawTrainingStatus(item);
                td.textContent = displayCellText(status.text);
                if (status.tone && status.tone !== "unknown") {
                  td.className = "cert-status cert-status--" + status.tone;
                }
              } else if (isWeaponsCertDateColumn(col)) {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                td.textContent = displayCellText(formatWeaponsCertDisplayDate(raw));
              } else if (col.key === "Certifier") {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                td.textContent = displayCellText(formatSharePointLookupDisplay(raw));
              } else {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                const text = raw !== undefined && raw !== null ? formatCellValue(raw) : "";
                td.textContent = displayCellText(text);
              }
              tr.appendChild(td);
            });
            if (showActions && item.Id != null) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              const inner = document.createElement("div");
              inner.className = "roster-actions-inner";
              const requalBtn = document.createElement("button");
              requalBtn.type = "button";
              requalBtn.className = "btn-record";
              requalBtn.textContent = "Requalify";
              const trainingLabel = formatCellValue(valueFromItemByKeys(item, BYLAW_TRAINING_ITEM_SORT_KEYS) || "");
              requalBtn.title = trainingLabel ? "Requalify " + trainingLabel : "Requalify training";
              requalBtn.addEventListener("click", function () {
                void openBylawTrainingEditPanel(item);
              });
              inner.appendChild(requalBtn);
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              delBtn.title = trainingLabel ? "Delete " + trainingLabel + " qualification" : "Delete training record";
              delBtn.addEventListener("click", function () {
                const personId = personDetailSession.item && personDetailSession.item.Id;
                void deleteBylawTrainingRow(item.Id, personDetailSession.pw, bylawTrainingListApiPath(), personId);
              });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            } else if (showActions) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              tdAct.textContent = "-";
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          personBylawBody.appendChild(frag);

          if (personBylawEmpty) {
            personBylawEmpty.hidden = rows.length > 0;
          }
        }

        function bylawPersonFilterFieldCandidates() {
          const primary = String(BYLAW_TRAINING_PERSON_FIELD || "PersonnelId").trim();
          const out = [];
          if (primary) out.push(primary);
          const alts = Array.isArray(BYLAW_TRAINING_PERSON_FIELD_ALT) ? BYLAW_TRAINING_PERSON_FIELD_ALT : [];
          alts.forEach(function (name) {
            const n = String(name || "").trim();
            if (n && out.indexOf(n) === -1) out.push(n);
          });
          return out;
        }

        function bylawPersonnelIdFromItem(item) {
          if (!item || typeof item !== "object") return null;
          const keys = bylawPersonFilterFieldCandidates().concat([
            "PersonnelId",
            "PersonnelID",
            "PersonnelIdId",
            "Personnel_x0020_Id",
          ]);
          for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (!k || k.indexOf("/") !== -1) continue;
            if (Object.prototype.hasOwnProperty.call(item, k) && item[k] !== null && item[k] !== "") {
              return String(item[k]);
            }
          }
          if (item.Personnel && typeof item.Personnel === "object" && item.Personnel.Id != null) {
            return String(item.Personnel.Id);
          }
          return null;
        }

        function bylawRowsMatchPersonnel(rows, personnelId) {
          const want = String(personnelId);
          return (Array.isArray(rows) ? rows : []).filter(function (item) {
            return bylawPersonnelIdFromItem(item) === want;
          });
        }

        async function discoverBylawPersonFilterFields(seg, pw) {
          if (hubSession.bylawPersonFilterFields && hubSession.bylawPersonFilterFields.length) {
            return hubSession.bylawPersonFilterFields.slice();
          }
          const out = [];
          try {
            const data = await spFetch(
              `/_api/web/${seg}/fields?$select=InternalName,Title,EntityPropertyName,StaticName,TypeAsString&$filter=Hidden eq false&$top=200`,
              {},
              pw,
            );
            const fields = (data && data.value) || [];
            const want = String(BYLAW_TRAINING_PERSON_FIELD || "PersonnelId").trim().toLowerCase();
            let hit = fields.find(function (f) {
              const title = String(f.Title || "").trim().toLowerCase();
              const internal = String(f.InternalName || f.StaticName || f.EntityPropertyName || "").trim();
              const internalLower = internal.toLowerCase();
              return (
                internalLower === want ||
                title === want ||
                title === "personnel id" ||
                internalLower === "personnelid" ||
                internalLower === "personnelidid"
              );
            });
            if (!hit) {
              hit = fields.find(function (f) {
                const internal = String(f.InternalName || f.StaticName || "").toLowerCase();
                return internal.indexOf("personnel") !== -1 && internal.indexOf("id") !== -1;
              });
            }
            if (hit) {
              const internal = String(hit.InternalName || hit.StaticName || hit.EntityPropertyName || "").trim();
              const type = String(hit.TypeAsString || "");
              if (internal) out.push(internal);
              if (/lookup/i.test(type) && internal) {
                const lookupIdField = internal.endsWith("Id") ? internal + "Id" : internal + "Id";
                if (out.indexOf(lookupIdField) === -1) out.push(lookupIdField);
                const slashField = internal + "/Id";
                if (out.indexOf(slashField) === -1) out.push(slashField);
              }
            }
          } catch (_) {}
          hubSession.bylawPersonFilterFields = out.slice();
          return out;
        }

        async function fetchBylawTrainingForPersonnel(seg, pw, personnelId) {
          const id = parseInt(String(personnelId), 10);
          if (!id || isNaN(id)) return [];
          const orderByClause = String(BYLAW_TRAINING_ITEMS_ORDERBY || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          const tried = [];
          const candidates = bylawPersonFilterFieldCandidates().slice();

          const discovered = await discoverBylawPersonFilterFields(seg, pw);
          discovered.forEach(function (name) {
            if (name && candidates.indexOf(name) === -1) candidates.unshift(name);
          });
          if (discovered.length) {
            discovered.slice().reverse().forEach(function (name) {
              if (name && candidates.indexOf(name) !== -1) {
                candidates.splice(candidates.indexOf(name), 1);
                candidates.unshift(name);
              }
            });
          }

          for (let i = 0; i < candidates.length; i++) {
            const personField = candidates[i];
            if (!personField || tried.indexOf(personField) !== -1) continue;
            tried.push(personField);
            const filterRaw = personField + " eq " + id;
            const filter = encodeURIComponent(filterRaw);
            try {
              let data = null;
              try {
                data = await spFetch(`/_api/web/${seg}/items?$top=200&$filter=${filter}` + orderByQs, {}, pw);
              } catch (e0) {
                if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
                  data = await spFetch(`/_api/web/${seg}/items?$top=200&$filter=${filter}`, {}, pw);
                } else {
                  throw e0;
                }
              }
              hubSession.bylawPersonFilterField = personField;
              return (data && data.value) || [];
            } catch (e1) {
              if (!/\b400\b/.test(String(e1.message || "")) && !/does not exist/i.test(String(e1.message || ""))) {
                throw e1;
              }
            }
          }

          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=500` + orderByQs, {}, pw);
          } catch (e2) {
            if (/\b400\b/.test(String(e2.message || "")) && orderByQs) {
              data = await spFetch(`/_api/web/${seg}/items?$top=500`, {}, pw);
            } else {
              throw e2;
            }
          }
          return bylawRowsMatchPersonnel((data && data.value) || [], id);
        }

        async function loadPersonBylawTraining(personId, pw) {
          if (!personId) return;
          if (personBylawWrap) personBylawWrap.hidden = false;
          setBylawTrainingState("loading", "Loading By-Law Training records...");
          clearBylawTrainingTable();
          if (personBylawEmpty) personBylawEmpty.hidden = true;

          const seg = bylawTrainingListApiPath();
          if (!bylawTrainingListTitle() && !bylawTrainingListUsesGuid()) {
            renderBylawTrainingTable([]);
            setBylawTrainingState("", "");
            if (personBylawEmpty) {
              personBylawEmpty.hidden = false;
              personBylawEmpty.textContent = "By-Law Training list is not configured.";
            }
            return;
          }

          try {
            const rows = await fetchBylawTrainingForPersonnel(seg, pw, personId);
            const dateKeys = [
              "QualDate",
              "QualificationDate",
              "Qualification_x0020_Date",
              "CertificationDate",
              "Certification_x0020_Date",
              "CertDate",
            ];
            rows.sort(function (a, b) {
              const da = formatCellValue(valueFromItemByKeys(a, dateKeys) || "");
              const db = formatCellValue(valueFromItemByKeys(b, dateKeys) || "");
              if (da !== db) return String(db).localeCompare(String(da));
              const wa = formatCellValue(valueFromItemByKeys(a, BYLAW_TRAINING_ITEM_SORT_KEYS) || "");
              const wb = formatCellValue(valueFromItemByKeys(b, BYLAW_TRAINING_ITEM_SORT_KEYS) || "");
              return String(wa).localeCompare(String(wb));
            });
            hubSession.bylawTrainingSampleRow = rows.length ? rows[0] : hubSession.bylawTrainingSampleRow || null;
            hubSession.bylawTrainingRows = rows.slice();
            renderBylawTrainingTable(rows);
            setBylawTrainingState("", "");
          } catch (e) {
            hubSession.bylawTrainingSampleRow = hubSession.bylawTrainingSampleRow || null;
            renderBylawTrainingTable([]);
            setBylawTrainingState("warn", "Could not load By-Law Training records: " + (e.message || String(e)).slice(0, 220));
            if (personBylawEmpty) {
              personBylawEmpty.hidden = false;
              personBylawEmpty.textContent =
                "No By-Law Training records loaded. If PersonnelId exists on the list, check that values match the Personnel Record Id.";
            }
          }
        }


        function setHubListViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = !visible;
          });
          if (personDetailSection) personDetailSection.hidden = visible;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
        }

        function setInstructorsViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = !visible;
          if (reportsSection) reportsSection.hidden = true;
        }

        function setReportsViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = !visible;
        }

        function setReportsState(kind, message) {
          if (!reportsReadState) return;
          if (!message) {
            reportsReadState.hidden = true;
            reportsReadState.textContent = "";
            return;
          }
          reportsReadState.hidden = false;
          reportsReadState.className = "read-state " + kind;
          reportsReadState.textContent = message;
        }

        function reportTypeById(reportId) {
          const id = String(reportId || "").trim();
          return (Array.isArray(HUB_REPORT_TYPES) ? HUB_REPORT_TYPES : []).find(function (r) {
            return r && String(r.id) === id;
          });
        }

        function showReportsHub() {
          reportsSession.activeReportId = null;
          if (reportsHubPanel) reportsHubPanel.hidden = false;
          if (reportsDetailPanel) reportsDetailPanel.hidden = true;
          if (reportsPrintBtn) reportsPrintBtn.hidden = true;
          renderReportsHubGrid();
        }

        function showReportsDetail(reportDef) {
          if (!reportDef) return;
          reportsSession.activeReportId = reportDef.id;
          if (reportsHubPanel) reportsHubPanel.hidden = true;
          if (reportsDetailPanel) reportsDetailPanel.hidden = false;
          if (reportsDetailTitle) reportsDetailTitle.textContent = reportDef.title || "Report";
          if (reportsDetailSubtitle) reportsDetailSubtitle.textContent = reportDef.subtitle || "";
          if (reportsPrintBtn) reportsPrintBtn.hidden = !reportDef.printable;
        }

        function renderReportsHubGrid() {
          if (!reportsHubGrid) return;
          reportsHubGrid.innerHTML = "";
          const frag = document.createDocumentFragment();
          (Array.isArray(HUB_REPORT_TYPES) ? HUB_REPORT_TYPES : []).forEach(function (def) {
            if (!def || !def.id) return;
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "reports-hub-card";
            btn.setAttribute("role", "listitem");
            btn.title = "Open " + (def.title || def.id);
            const badge = document.createElement("span");
            badge.className = "reports-hub-card-badge";
            badge.textContent = def.badge || "RPT";
            const title = document.createElement("span");
            title.className = "reports-hub-card-title";
            title.textContent = def.title || def.id;
            const desc = document.createElement("span");
            desc.className = "reports-hub-card-desc";
            desc.textContent = def.subtitle || "";
            const status = document.createElement("span");
            status.className =
              "reports-hub-card-status " + (def.implemented ? "reports-hub-card-status--ready" : "reports-hub-card-status--soon");
            status.textContent = def.implemented ? "Ready" : "Coming soon";
            btn.appendChild(badge);
            btn.appendChild(title);
            btn.appendChild(desc);
            btn.appendChild(status);
            btn.addEventListener("click", function () {
              void openHubReport(def.id);
            });
            frag.appendChild(btn);
          });
          reportsHubGrid.appendChild(frag);
        }

        function trainingStatusToneRank(tone) {
          const ranks = { unknown: 0, ok: 1, warn: 2, urgent: 3, expired: 4 };
          return ranks[String(tone || "unknown")] != null ? ranks[String(tone || "unknown")] : 0;
        }

        function worstTrainingStatusTone(tones) {
          let worst = "unknown";
          (Array.isArray(tones) ? tones : []).forEach(function (tone) {
            if (trainingStatusToneRank(tone) > trainingStatusToneRank(worst)) worst = tone;
          });
          return worst;
        }

        function summarizeTrainingRows(rows, computeFn) {
          const list = Array.isArray(rows) ? rows : [];
          if (!list.length) return { text: "No records", tone: "unknown", count: 0 };
          const statuses = list.map(function (row) {
            return computeFn(row);
          });
          const tone = worstTrainingStatusTone(
            statuses.map(function (s) {
              return s.tone;
            }),
          );
          let text = "Qualified";
          if (tone === "expired") text = "Expired";
          else if (tone === "urgent") text = "Due <=30d";
          else if (tone === "warn") text = "Due 31-60d";
          else if (tone === "unknown") text = "No records";
          return { text: text, tone: tone, count: list.length };
        }

        function applyReportStatusCell(td, status) {
          const s = status || { text: "-", tone: "unknown" };
          td.textContent = displayCellText(s.text);
          if (s.tone && s.tone !== "unknown") {
            td.className = "cert-status cert-status--" + s.tone;
          }
        }

        function personOfficeLabel(item) {
          const office = itemFieldText(item, "OfficeSymbol");
          const squadron = itemFieldText(item, "Squadron");
          const parts = [office, squadron].filter(Boolean);
          return parts.length ? parts.join(" / ") : "-";
        }

        async function fetchAllWeaponsCertRows(pw) {
          const seg = weaponsCertListApiPath();
          if (!seg || !pw) return [];
          try {
            const data = await spFetch(`/_api/web/${seg}/items?$top=5000`, {}, pw);
            return (data && data.value) || [];
          } catch (e) {
            log("Reports weapons fetch failed:\n" + (e.message || String(e)), "warn");
            return [];
          }
        }

        async function fetchAllBylawTrainingRows(pw) {
          const seg = bylawTrainingListApiPath();
          if (!seg || !pw) return [];
          try {
            const data = await spFetch(`/_api/web/${seg}/items?$top=5000`, {}, pw);
            return (data && data.value) || [];
          } catch (e) {
            log("Reports by-law fetch failed:\n" + (e.message || String(e)), "warn");
            return [];
          }
        }

        async function ensureReportsTrainingCache(pw) {
          if (!pw) return { weaponsRows: [], bylawRows: [] };
          if (!Array.isArray(reportsSession.weaponsRows)) {
            reportsSession.weaponsRows = await fetchAllWeaponsCertRows(pw);
          }
          if (!Array.isArray(reportsSession.bylawRows)) {
            reportsSession.bylawRows = await fetchAllBylawTrainingRows(pw);
          }
          return {
            weaponsRows: reportsSession.weaponsRows || [],
            bylawRows: reportsSession.bylawRows || [],
          };
        }

        function buildStatusOfTrainingRows(personRows, weaponsRows, bylawRows) {
          return (Array.isArray(personRows) ? personRows : []).map(function (person) {
            if (!person || person.Id == null) return null;
            const pid = String(person.Id);
            const personWeapons = weaponsRows.filter(function (row) {
              return weaponsPersonnelIdFromItem(row) === pid;
            });
            const personBylaw = bylawRows.filter(function (row) {
              return bylawPersonnelIdFromItem(row) === pid;
            });
            const weaponsStatus = summarizeTrainingRows(personWeapons, computeWeaponsCertStatus);
            const bylawStatus = summarizeTrainingRows(personBylaw, computeBylawTrainingStatus);
            const overallTone = worstTrainingStatusTone([weaponsStatus.tone, bylawStatus.tone]);
            let overallText = "Qualified";
            if (overallTone === "expired") overallText = "Expired";
            else if (overallTone === "urgent") overallText = "Action needed";
            else if (overallTone === "warn") overallText = "Due soon";
            else if (overallTone === "unknown") overallText = "No training data";
            return {
              person: person,
              weaponsStatus: weaponsStatus,
              bylawStatus: bylawStatus,
              overall: { text: overallText, tone: overallTone },
            };
          }).filter(Boolean);
        }

        function renderStatusOfTrainingSummary(rows) {
          const counts = { total: rows.length, ok: 0, warn: 0, urgent: 0, expired: 0, unknown: 0 };
          rows.forEach(function (entry) {
            const tone = entry.overall && entry.overall.tone ? entry.overall.tone : "unknown";
            if (counts[tone] != null) counts[tone] += 1;
          });
          const grid = document.createElement("div");
          grid.className = "reports-summary-grid";
          const stats = [
            { label: "Personnel", value: counts.total, mod: "" },
            { label: "Qualified", value: counts.ok, mod: "" },
            { label: "Due 31-60d", value: counts.warn, mod: "warn" },
            { label: "Due <=30d", value: counts.urgent, mod: "urgent" },
            { label: "Expired", value: counts.expired, mod: "expired" },
            { label: "No data", value: counts.unknown, mod: "" },
          ];
          stats.forEach(function (stat) {
            const box = document.createElement("div");
            box.className = "reports-summary-stat" + (stat.mod ? " reports-summary-stat--" + stat.mod : "");
            const val = document.createElement("span");
            val.className = "reports-summary-stat-value";
            val.textContent = String(stat.value);
            const lab = document.createElement("span");
            lab.className = "reports-summary-stat-label";
            lab.textContent = stat.label;
            box.appendChild(val);
            box.appendChild(lab);
            grid.appendChild(box);
          });
          return grid;
        }

        function renderStatusOfTrainingTable(rows) {
          const wrap = document.createElement("div");
          wrap.className = "roster-wrap";
          const table = document.createElement("table");
          table.className = "roster reports-status-table";
          table.setAttribute("aria-label", "Status of Training");
          const thead = document.createElement("thead");
          const trHead = document.createElement("tr");
          ["Personnel", "Office / Squadron", "Weapons", "By-Law", "Overall"].forEach(function (label) {
            const th = document.createElement("th");
            th.textContent = label;
            trHead.appendChild(th);
          });
          thead.appendChild(trHead);
          table.appendChild(thead);
          const tbody = document.createElement("tbody");
          const sorted = rows.slice().sort(function (a, b) {
            return formatPersonDisplayName(a.person).localeCompare(formatPersonDisplayName(b.person), undefined, {
              sensitivity: "base",
            });
          });
          const frag = document.createDocumentFragment();
          sorted.forEach(function (entry) {
            const tr = document.createElement("tr");
            const tdName = document.createElement("td");
            const nameBtn = document.createElement("button");
            nameBtn.type = "button";
            nameBtn.className = "reports-name-link";
            nameBtn.textContent = formatPersonDisplayName(entry.person);
            nameBtn.title = "Open Personnel Record";
            nameBtn.addEventListener("click", function () {
              navigateToPersonDetail(entry.person.Id);
            });
            tdName.appendChild(nameBtn);
            tr.appendChild(tdName);
            const tdOffice = document.createElement("td");
            tdOffice.textContent = displayCellText(personOfficeLabel(entry.person));
            tr.appendChild(tdOffice);
            const tdWeapons = document.createElement("td");
            applyReportStatusCell(tdWeapons, entry.weaponsStatus);
            tr.appendChild(tdWeapons);
            const tdBylaw = document.createElement("td");
            applyReportStatusCell(tdBylaw, entry.bylawStatus);
            tr.appendChild(tdBylaw);
            const tdOverall = document.createElement("td");
            applyReportStatusCell(tdOverall, entry.overall);
            tr.appendChild(tdOverall);
            frag.appendChild(tr);
          });
          tbody.appendChild(frag);
          table.appendChild(tbody);
          wrap.appendChild(table);
          return wrap;
        }

        async function loadStatusOfTrainingReport(reportDef) {
          const pw = hubSession.pw;
          if (!pw) {
            setReportsState("err", "Hub not ready. Refresh the Personnel Roster first.");
            return;
          }
          if (!reportsDetailBody) return;
          showReportsDetail(reportDef);
          reportsDetailBody.innerHTML = "";
          setReportsState("loading", "Building Status of Training report...");
          try {
            const personRows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
            if (!personRows.length) {
              setReportsState("warn", "Personnel Roster is empty. Refresh the list first.");
              const empty = document.createElement("div");
              empty.className = "reports-placeholder-panel";
              const emptyMsg = document.createElement("p");
              emptyMsg.textContent = "No Personnel rows loaded.";
              empty.appendChild(emptyMsg);
              reportsDetailBody.appendChild(empty);
              return;
            }
            const cache = await ensureReportsTrainingCache(pw);
            const rows = buildStatusOfTrainingRows(personRows, cache.weaponsRows, cache.bylawRows);
            reportsDetailBody.appendChild(renderStatusOfTrainingSummary(rows));
            reportsDetailBody.appendChild(renderStatusOfTrainingTable(rows));
            setReportsState("ok", "Status of Training updated.");
            window.setTimeout(function () {
              setReportsState("", "");
            }, 2200);
          } catch (e) {
            setReportsState("err", "Could not build report: " + (e.message || String(e)).slice(0, 220));
            log("Status of Training report failed:\n" + (e.message || String(e)), "err");
          }
        }

        function renderPlaceholderReport(reportDef) {
          showReportsDetail(reportDef);
          if (!reportsDetailBody) return;
          reportsDetailBody.innerHTML = "";
          const panel = document.createElement("div");
          panel.className = "reports-placeholder-panel";
          const p = document.createElement("p");
          p.textContent =
            reportDef.placeholder ||
            "This report is queued for a future hub update. The layout shell is ready so data can be wired in next.";
          panel.appendChild(p);
          reportsDetailBody.appendChild(panel);
          setReportsState("", "");
        }

        async function openHubReport(reportId) {
          const def = reportTypeById(reportId);
          if (!def) return;
          setReportsState("", "");
          if (def.implemented && def.id === "status-of-training") {
            await loadStatusOfTrainingReport(def);
            return;
          }
          renderPlaceholderReport(def);
        }

        function invalidateReportsTrainingCache() {
          reportsSession.weaponsRows = null;
          reportsSession.bylawRows = null;
        }

        function printActiveReport() {
          const def = reportTypeById(reportsSession.activeReportId);
          if (!def || !def.printable) return;
          window.print();
        }

        function setInstructorsState(kind, message) {
          if (!instructorsReadState) return;
          if (!message) {
            instructorsReadState.hidden = true;
            instructorsReadState.textContent = "";
            return;
          }
          instructorsReadState.hidden = false;
          instructorsReadState.className = "read-state " + kind;
          instructorsReadState.textContent = message;
        }

        function setInstructorsAddPanelVisible(show) {
          if (instructorsAddPanel) instructorsAddPanel.hidden = !show;
          if (!show && instructorsAddForm) instructorsAddForm.reset();
        }

        function certifierDisplayLabel(item) {
          if (!item) return "";
          const personKeys = ["Personnel", "PersonnelId", "PersonnelID"];
          if (hubSession.certifiersPersonDisplayKey) personKeys.unshift(hubSession.certifiersPersonDisplayKey);
          for (let i = 0; i < personKeys.length; i++) {
            const raw = valueFromItemByKeys(item, [personKeys[i]]);
            const fromLookup = formatSharePointLookupDisplay(raw);
            if (fromLookup) return fromLookup;
          }
          const label = rowLabelFromListItem(item, CERTIFIERS_LIST_DISPLAY_FIELD);
          if (label) return label;
          return item.Id != null ? "Instructor #" + item.Id : "Instructor";
        }

        function certifierPersonnelIdFromItem(item) {
          if (!item) return null;
          const keys = [];
          if (hubSession.certifiersPersonPostKey) keys.push(hubSession.certifiersPersonPostKey);
          if (hubSession.certifiersPersonDisplayKey) keys.push(hubSession.certifiersPersonDisplayKey);
          keys.push("PersonnelId", "PersonnelID", "PersonnelIdId", "Personnel");
          const seen = new Set();
          for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (!k || seen.has(k)) continue;
            seen.add(k);
            if (item[k] == null || item[k] === "") continue;
            const v = item[k];
            if (typeof v === "object" && v.Id != null) {
              const n = parseInt(String(v.Id), 10);
              if (n && !isNaN(n)) return n;
            }
            const n = parseInt(String(v), 10);
            if (n && !isNaN(n)) return n;
          }
          return null;
        }

        function clearInstructorsTable() {
          if (instructorsThead) instructorsThead.innerHTML = "";
          if (instructorsTableBody) instructorsTableBody.innerHTML = "";
        }

        function renderInstructorsTable(rows) {
          clearInstructorsTable();
          const pw = hubSession.pw;
          const showActions = !!(pw && instructorsSection && !instructorsSection.hidden);
          if (!instructorsThead || !instructorsTableBody) return;

          const trHead = document.createElement("tr");
          const thName = document.createElement("th");
          thName.textContent = "Instructor";
          trHead.appendChild(thName);
          if (showActions) {
            const thAct = document.createElement("th");
            thAct.className = "roster-actions";
            thAct.textContent = " ";
            thAct.title = "Delete instructor";
            trHead.appendChild(thAct);
          }
          instructorsThead.appendChild(trHead);

          const sorted = (Array.isArray(rows) ? rows.slice() : []).sort(function (a, b) {
            return certifierDisplayLabel(a).localeCompare(certifierDisplayLabel(b), undefined, { sensitivity: "base" });
          });

          const frag = document.createDocumentFragment();
          sorted.forEach(function (item) {
            const tr = document.createElement("tr");
            const tdName = document.createElement("td");
            tdName.textContent = displayCellText(certifierDisplayLabel(item));
            tr.appendChild(tdName);
            if (showActions && item.Id != null) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              const inner = document.createElement("div");
              inner.className = "roster-actions-inner";
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              const label = certifierDisplayLabel(item);
              delBtn.title = label ? "Remove " + label : "Remove instructor";
              delBtn.addEventListener("click", function () {
                void deleteCertifierRow(item.Id, hubSession.pw);
              });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            } else if (showActions) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              tdAct.textContent = "-";
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          instructorsTableBody.appendChild(frag);

          if (instructorsEmpty) {
            instructorsEmpty.hidden = sorted.length > 0;
            instructorsEmpty.textContent = sorted.length ? "" : "No instructors on file.";
          }
        }

        async function discoverCertifiersPersonnelField(pw) {
          if (hubSession.certifiersPersonPostKey !== undefined) return hubSession.certifiersPersonPostKey;
          const seg = certifiersListApiSegment();
          if (!seg || !pw) {
            hubSession.certifiersPersonPostKey = null;
            hubSession.certifiersPersonDisplayKey = null;
            return null;
          }
          const hint = String(CERTIFIERS_PERSON_FIELD || "PersonnelId").trim();
          try {
            const esc = hint.replace(/'/g, "''");
            const direct = await spFetch(
              `/_api/web/${seg}/fields/getbyinternalnameortitle('${esc}')?$select=Title,InternalName,TypeAsString`,
              {},
              pw,
            ).catch(function () {
              return null;
            });
            if (direct && /lookup/i.test(String(direct.TypeAsString || ""))) {
              const internal = String(direct.InternalName || hint).trim();
              hubSession.certifiersPersonPostKey = internal.endsWith("Id") ? internal : internal + "Id";
              hubSession.certifiersPersonDisplayKey = internal.endsWith("Id") ? internal.slice(0, -2) : internal;
              return hubSession.certifiersPersonPostKey;
            }
            const data = await spFetch(
              `/_api/web/${seg}/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false&$top=200`,
              {},
              pw,
            );
            const fields = (data && data.value) || [];
            let hit = fields.find(function (f) {
              return /lookup/i.test(String(f.TypeAsString || "")) && /personnel/i.test(String(f.Title || "") + String(f.InternalName || ""));
            });
            if (!hit) {
              hit = fields.find(function (f) {
                return /lookup/i.test(String(f.TypeAsString || "")) && /personnelid/i.test(String(f.InternalName || ""));
              });
            }
            if (hit) {
              const internal = String(hit.InternalName || "").trim();
              hubSession.certifiersPersonPostKey = internal.endsWith("Id") ? internal : internal + "Id";
              hubSession.certifiersPersonDisplayKey = internal.endsWith("Id") ? internal.slice(0, -2) : internal;
              return hubSession.certifiersPersonPostKey;
            }
          } catch (e) {
            log("Certifiers personnel field discovery failed: " + (e.message || String(e)), "warn");
          }
          hubSession.certifiersPersonPostKey = null;
          hubSession.certifiersPersonDisplayKey = null;
          return null;
        }

        async function fetchCertifierRows(pw) {
          const seg = certifiersListApiSegment();
          if (!seg) return [];
          const data = await spFetch(`/_api/web/${seg}/items?$top=500`, {}, pw);
          return (data && data.value) || [];
        }

        async function loadInstructorsList() {
          const pw = hubSession.pw;
          const seg = certifiersListApiSegment();
          if (!pw || !seg) {
            renderInstructorsTable([]);
            setInstructorsState("warn", "Certifiers list is not available.");
            return;
          }
          try {
            setInstructorsState("loading", "Loading instructors...");
            await discoverCertifiersPersonnelField(pw);
            const rows = await fetchCertifierRows(pw);
            instructorsSession.rows = rows;
            hubSession.certifiersSampleRow = rows.length ? rows[0] : null;
            renderInstructorsTable(rows);
            setInstructorsState("", "");
          } catch (e) {
            instructorsSession.rows = [];
            renderInstructorsTable([]);
            setInstructorsState("err", "Could not load instructors: " + (e.message || String(e)).slice(0, 220));
            log("Instructors load failed:\n" + (e.message || String(e)), "err");
          }
        }

        async function deleteCertifierRow(id, pw) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid)) {
            setInstructorsState("err", "Invalid instructor Id for delete.");
            return;
          }
          const seg = certifiersListApiSegment();
          if (!pw || !seg) {
            setInstructorsState("err", "Certifiers list is not available.");
            return;
          }
          if (!confirm("Remove instructor Id " + sid + " from the Certifiers list? This cannot be undone.")) return;
          try {
            setInstructorsState("loading", "Deleting instructor...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            log("DELETE certifier Id " + sid + " OK", "ok");
            await loadInstructorsList();
            setInstructorsState("ok", "Instructor removed.");
            window.setTimeout(function () {
              setInstructorsState("", "");
            }, 2500);
          } catch (e) {
            setInstructorsState("err", "Delete failed: " + (e.message || String(e)).slice(0, 280));
            log("Certifier DELETE failed:\n" + (e.message || String(e)), "err");
          }
        }

        function populateInstructorsPersonSelect() {
          if (!instructorsPersonSelect) return;
          const prior = String(instructorsPersonSelect.value || "").trim();
          while (instructorsPersonSelect.options.length > 1) instructorsPersonSelect.remove(1);
          const rows = Array.isArray(hubSession.rows) ? hubSession.rows.slice() : [];
          const existing = Array.isArray(instructorsSession.rows) ? instructorsSession.rows : [];
          const takenIds = new Set();
          const takenNames = new Set();
          existing.forEach(function (row) {
            const pid = certifierPersonnelIdFromItem(row);
            if (pid) takenIds.add(String(pid));
            const label = certifierDisplayLabel(row).toLowerCase();
            if (label) takenNames.add(label);
          });
          rows.sort(function (a, b) {
            return formatPersonDisplayName(a).localeCompare(formatPersonDisplayName(b), undefined, { sensitivity: "base" });
          });
          rows.forEach(function (person) {
            if (!person || person.Id == null) return;
            const displayName = formatPersonDisplayName(person);
            if (!displayName) return;
            if (takenIds.has(String(person.Id)) || takenNames.has(displayName.toLowerCase())) return;
            const o = document.createElement("option");
            o.value = String(person.Id);
            o.textContent = displayName;
            instructorsPersonSelect.appendChild(o);
          });
          if (prior) instructorsPersonSelect.value = prior;
          applySelectAutosize(instructorsPersonSelect);
        }

        async function submitInstructorAdd(ev) {
          if (ev && ev.preventDefault) ev.preventDefault();
          const pw = hubSession.pw;
          const seg = certifiersListApiSegment();
          if (!pw || !seg) {
            setInstructorsState("err", "Hub not ready. Refresh the roster first.");
            return;
          }
          const personId = instructorsPersonSelect ? parseInt(String(instructorsPersonSelect.value || "").trim(), 10) : 0;
          if (!personId || isNaN(personId)) {
            setInstructorsState("err", "Select a person from the Personnel Roster.");
            return;
          }
          const person = (hubSession.rows || []).find(function (r) {
            return r && String(r.Id) === String(personId);
          });
          if (!person) {
            setInstructorsState("err", "Person not found in roster. Refresh the list and try again.");
            return;
          }
          const displayName = formatPersonDisplayName(person);
          const existing = Array.isArray(instructorsSession.rows) ? instructorsSession.rows : [];
          for (let i = 0; i < existing.length; i++) {
            const row = existing[i];
            const existingPid = certifierPersonnelIdFromItem(row);
            if (existingPid && existingPid === personId) {
              setInstructorsState("warn", displayName + " is already listed as an instructor.");
              return;
            }
            if (certifierDisplayLabel(row).toLowerCase() === displayName.toLowerCase()) {
              setInstructorsState("warn", displayName + " is already listed as an instructor.");
              return;
            }
          }

          const personPostKey = await discoverCertifiersPersonnelField(pw);
          const certField = String(CERTIFIERS_LIST_DISPLAY_FIELD || "Certifier").trim();
          const payload = { Title: displayName };
          if (personPostKey) payload[personPostKey] = personId;
          if (!personPostKey || certField) payload[certField] = displayName;

          try {
            setInstructorsState("loading", "Adding instructor...");
            if (instructorsSaveBtn) instructorsSaveBtn.disabled = true;
            await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            setInstructorsAddPanelVisible(false);
            await loadInstructorsList();
            setInstructorsState("ok", "Instructor added.");
            window.setTimeout(function () {
              setInstructorsState("", "");
            }, 2500);
          } catch (e) {
            setInstructorsState("err", "Add failed: " + (e.message || String(e)).slice(0, 280));
            log("Certifier add failed:\n" + (e.message || String(e)), "err");
          } finally {
            if (instructorsSaveBtn) instructorsSaveBtn.disabled = false;
          }
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

        function setPersonDetailEditMode(editing) {
          personDetailSession.editing = !!editing;
          if (personDetailEditBtn) personDetailEditBtn.hidden = !!editing;
          if (personDetailSaveBtn) personDetailSaveBtn.hidden = !editing;
          if (personDetailCancelBtn) personDetailCancelBtn.hidden = !editing;
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
          return rank || name || "Personnel Record";
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
          val.textContent = displayCellText(text);
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
          if (el.tagName === "SELECT" && el.dataset.writeKey && /Id$/.test(el.dataset.writeKey)) {
            const idKey = el.dataset.writeKey;
            if (item[idKey] != null && item[idKey] !== "") {
              el.value = String(item[idKey]);
              ensureSelectIncludesValue(el, el.value);
              return;
            }
          }
          const tryKeys = col.tryKeys || [col.key];
          const raw = valueFromItemByKeys(item, tryKeys);
          if (raw === undefined || raw === null || raw === "") return;
          const text = formatCellValue(raw);
          if (el.type === "date") {
            const m = text.match(/^(\d{4}-\d{2}-\d{2})/);
            if (m) el.value = m[1];
            return;
          }
          if (el.tagName === "SELECT") {
            for (let i = 0; i < el.options.length; i++) {
              if (el.options[i].value === text || el.options[i].textContent === text) {
                el.value = el.options[i].value;
                return;
              }
            }
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

        async function populatePersonEditDropdowns(seg, pw, sampleRow, root, item) {
          const cols = normalizedRosterColumns();
          for (let i = 0; i < cols.length; i++) {
            const col = cols[i];
            if (col.key === "Id") continue;
            if (DROPDOWN_COLUMN_KEYS.indexOf(col.key) === -1) continue;
            const sel = root.querySelector("#pf_" + col.key);
            await fillDropdownSelect(sel, col, seg, pw, sampleRow, item);
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
            await populatePersonEditDropdowns(
              personDetailSession.seg,
              personDetailSession.pw,
              personDetailSession.sampleRow,
              form,
              item,
            );
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                balanceContactNotesToIdentityColumn(form);
              });
            });
          } else {
            personDetailContent.appendChild(fields);
          }
        }

        function mergePayloadIntoItem(item, body) {
          const updated = Object.assign({}, item);
          const cols = normalizedRosterColumns();
          Object.keys(body || {}).forEach(function (key) {
            if (key === "Title") {
              updated.Title = body[key];
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
            updated[targetKey] = val;
          });
          return updated;
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
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });

          if (SET_TITLE_ON_CREATE) {
            payload.Title = [lastName, firstName].filter(Boolean).join(", ") || dod || "Personnel Record";
          }

          try {
            setPersonDetailState("loading", "Saving changes...");
            if (personDetailSaveBtn) personDetailSaveBtn.disabled = true;
            await spFetch(`/_api/web/${s.seg}/items(${s.item.Id})`, { method: "MERGE", body: payload }, s.pw);
            const updated = mergePayloadIntoItem(s.item, payload);
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
            setPersonDetailState("loading", "Loading Personnel Record...");
            if (personDetailTitle) personDetailTitle.textContent = "Personnel Record";
            try {
              item = await spFetch(`/_api/web/${seg}/items(${itemId})`, {}, pw);
            } catch (e) {
              setPersonDetailState("err", "Cannot load Personnel Record: " + (e.message || String(e)).slice(0, 280));
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
          await loadPersonWeaponsCertifications(item.Id, personDetailSession.pw);
          await loadPersonBylawTraining(item.Id, personDetailSession.pw);
        }

        async function showPersonDetailById(itemId, meta, pw, seg, cachedRows) {
          const sampleRow =
            (Array.isArray(cachedRows) && cachedRows.length ? cachedRows[0] : null) || hubSession.sampleRow;
          await openPersonDetailRecord(itemId, meta, pw, seg, cachedRows, sampleRow);
        }

        if (personDetailBackLink) {
          personDetailBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (instructorsBackLink) {
          instructorsBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (hubNavInstructors) {
          hubNavInstructors.addEventListener("click", function () {
            void navigateToInstructors();
          });
        }

        if (instructorsAddBtn) {
          instructorsAddBtn.addEventListener("click", function () {
            if (!hubSession.rows || !hubSession.rows.length) {
              setInstructorsState("warn", "Personnel Roster is empty. Refresh the list first.");
              return;
            }
            populateInstructorsPersonSelect();
            setInstructorsAddPanelVisible(true);
            setInstructorsState("", "");
            if (instructorsPersonSelect) instructorsPersonSelect.focus();
          });
        }

        if (instructorsAddCancelBtn) {
          instructorsAddCancelBtn.addEventListener("click", function () {
            setInstructorsAddPanelVisible(false);
          });
        }

        if (instructorsAddForm) {
          instructorsAddForm.addEventListener("submit", function (ev) {
            void submitInstructorAdd(ev);
          });
        }

        if (reportsBackLink) {
          reportsBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (reportsDetailBackLink) {
          reportsDetailBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            showReportsHub();
            setReportsState("", "");
          });
        }

        if (hubNavReports) {
          hubNavReports.addEventListener("click", function () {
            void navigateToReports();
          });
        }

        if (reportsPrintBtn) {
          reportsPrintBtn.addEventListener("click", function () {
            printActiveReport();
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

        if (personWeaponsAddBtn) {
          personWeaponsAddBtn.addEventListener("click", function () {
            if (personWeaponsAddPanel && !personWeaponsAddPanel.hidden) {
              setWeaponsCertAddPanelVisible(false);
              if (personWeaponsAddForm) personWeaponsAddForm.reset();
              return;
            }
            void openWeaponsCertAddPanel();
          });
        }

        if (personWeaponsAddForm) {
          personWeaponsAddForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitWeaponsCertSave();
          });
        }

        if (personWeaponsAddCancelBtn) {
          personWeaponsAddCancelBtn.addEventListener("click", function () {
            setWeaponsCertAddPanelVisible(false);
            if (personWeaponsAddForm) personWeaponsAddForm.reset();
          });
        }

        if (personBylawAddBtn) {
          personBylawAddBtn.addEventListener("click", function () {
            if (personBylawAddPanel && !personBylawAddPanel.hidden) {
              setBylawTrainingAddPanelVisible(false);
              if (personBylawAddForm) personBylawAddForm.reset();
              return;
            }
            void openBylawTrainingAddPanel();
          });
        }

        if (personBylawAddForm) {
          personBylawAddForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitBylawTrainingSave();
          });
        }

        if (personBylawAddCancelBtn) {
          personBylawAddCancelBtn.addEventListener("click", function () {
            setBylawTrainingAddPanelVisible(false);
            if (personBylawAddForm) personBylawAddForm.reset();
          });
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
            th.title = (col.tryKeys || [col.key]).join(" * ");
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
              td.textContent = displayCellText(text);
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
              tdAct.textContent = "-";
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

          const listTitle = meta && meta.Title ? String(meta.Title) : personnelListTitle();
          const ic = meta && meta.ItemCount != null ? meta.ItemCount : "?";
          if (rows.length > 0) {
            setReadState(
              "ok",
              "Loaded " +
                rows.length +
                " row(s), " +
                columns.length +
                " column(s) (" +
                (plan.mode === "explicit" ? "explicit ROSTER_COLUMNS" : "auto field union") +
                "). List \"" +
                listTitle +
                "\" - ItemCount " +
                ic +
                ".",
            );
          }
        }

        function log(msg, cls) {
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
          const el =
            document.getElementById("__REQUESTDIGEST") ||
            (function () {
              try {
                return window.parent.document.getElementById("__REQUESTDIGEST");
              } catch (_) {
                return null;
              }
            })() ||
            (function () {
              try {
                return window.top.document.getElementById("__REQUESTDIGEST");
              } catch (_) {
                return null;
              }
            })();
          return el && el.value ? el.value : "";
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
          const base =
            baseUrlOverride != null && String(baseUrlOverride).trim()
              ? String(baseUrlOverride).trim().replace(/\/$/, "")
              : resolvedPersonnelRestBase();
          const headers = Object.assign({}, opts.headers || {});
          headers["Accept"] = "application/json;odata=nometadata";
          const method = (opts.method || "GET").toUpperCase();
          if (method !== "GET" && method !== "DELETE") {
            headers["Content-Type"] = "application/json;odata=nometadata";
          }

          const isPostWithBody = method === "POST" && opts.body != null;
          const needsWriteDigest =
            !opts.skipAntiForgery && (method === "MERGE" || method === "DELETE" || isPostWithBody);

          let digest = "";
          if (needsWriteDigest) {
            try {
              digest = await fetchFormDigestForWeb(base);
            } catch (_) {
              digest = "";
            }
            if (!digest) digest = getDigest();
          } else {
            digest = getDigest();
          }
          if (digest) headers["X-RequestDigest"] = digest;

          let realMethod = method;
          if (method === "MERGE" || method === "DELETE") {
            headers["X-HTTP-Method"] = method;
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
          log("--- Listing lists on this web (first 200, not hidden) - find exact Title for LIST_PERSONNEL ---");
          const data = await spFetch(
            "/_api/web/lists?$select=Title,Id,ItemCount,Hidden&$filter=Hidden eq false&$orderby=Title&$top=200",
            {},
            pw,
          );
          const rows = (data && data.value) || [];
          rows.forEach((r) => {
            log("  * " + (r.Title || "") + "  (Id=" + r.Id + ", ItemCount=" + (r.ItemCount != null ? r.ItemCount : "?") + ")");
          });
        }

        function listItemDisplayFieldCandidates(displayField) {
          const out = [];
          function add(k) {
            const s = String(k || "").trim();
            if (s && out.indexOf(s) === -1) out.push(s);
          }
          add(displayField);
          add(CERTIFIERS_LIST_DISPLAY_FIELD);
          add("Certifier");
          add("Title");
          add("Name");
          return out;
        }

        function rowLabelFromListItem(row, displayField) {
          if (!row || typeof row !== "object") return "";
          const candidates = listItemDisplayFieldCandidates(displayField);
          for (let i = 0; i < candidates.length; i++) {
            const key = candidates[i];
            if (row[key] != null && row[key] !== "") return String(row[key]).trim();
          }
          const keys = Object.keys(row);
          for (let i = 0; i < candidates.length; i++) {
            const want = candidates[i].toLowerCase();
            for (let j = 0; j < keys.length; j++) {
              const k = keys[j];
              if (k.toLowerCase() !== want) continue;
              if (row[k] != null && row[k] !== "") return String(row[k]).trim();
            }
          }
          return "";
        }

        function listItemChoicesFromRows(rows, displayField) {
          const seen = new Set();
          const out = [];
          (Array.isArray(rows) ? rows : []).forEach(function (row) {
            if (!row) return;
            const label = rowLabelFromListItem(row, displayField);
            const id = row.Id;
            if (!label || id == null || seen.has(String(id))) return;
            seen.add(String(id));
            out.push({ id: id, text: label });
          });
          out.sort(function (a, b) {
            return String(a.text).localeCompare(String(b.text), undefined, { sensitivity: "base" });
          });
          return out;
        }

        function certifiersListApiSegment() {
          const guid = String(LIST_CERTIFIERS_GUID || "").replace(/[{}' ]/g, "");
          if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(guid)) {
            return "lists(guid'" + guid + "')";
          }
          const title = String(LIST_CERTIFIERS || "").trim();
          if (!title) return "";
          return "lists/getbytitle('" + escListTitle(title) + "')";
        }

        async function fetchListItemChoices(pw, listTitle, listGuid, displayField) {
          const field = String(displayField || "Title").trim();
          let guid = String(listGuid || "").replace(/[{}' ]/g, "");
          let listSeg = "";
          if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(guid)) {
            listSeg = "lists(guid'" + guid + "')";
          } else {
            const title = String(listTitle || "").trim();
            if (!title) return [];
            listSeg = "lists/getbytitle('" + escListTitle(title) + "')";
          }

          async function loadRows(withSelect) {
            const base = `/_api/web/${listSeg}/items?$top=500`;
            const path = withSelect ? base + "&$select=Id," + encodeURIComponent(field) : base;
            try {
              const data = await spFetch(path, {}, pw);
              return (data && data.value) || [];
            } catch (e) {
              if (withSelect && /\b400\b/.test(String(e.message || ""))) return loadRows(false);
              return [];
            }
          }

          let rows = await loadRows(true);
          if (!rows.length) rows = await loadRows(false);
          return listItemChoicesFromRows(rows, field);
        }

        async function fetchCertifiersListEntries(pw) {
          const fieldsToTry = listItemDisplayFieldCandidates(CERTIFIERS_LIST_DISPLAY_FIELD);
          for (let i = 0; i < fieldsToTry.length; i++) {
            const entries = await fetchListItemChoices(pw, LIST_CERTIFIERS, LIST_CERTIFIERS_GUID, fieldsToTry[i]);
            if (entries.length) return entries;
          }
          return [];
        }

        function selectHasRealChoices(sel) {
          if (!sel || sel.tagName !== "SELECT") return false;
          for (let i = 0; i < sel.options.length; i++) {
            const o = sel.options[i];
            if (!o || o.disabled || !String(o.value || "").trim()) continue;
            const t = String(o.textContent || "");
            if (t.indexOf("(no choices") === 0 || t.indexOf("(error loading") === 0) continue;
            return true;
          }
          return false;
        }

        function appendCertifierSelectOptions(sel, entries) {
          if (!sel || !entries || !entries.length) return;
          const priorValue = String(sel.value || "").trim();
          while (sel.options.length > 1) sel.remove(1);
          entries.forEach(function (entry) {
            const o = document.createElement("option");
            o.value = String(entry.id);
            o.textContent = entry.text;
            sel.appendChild(o);
          });
          if (priorValue) sel.value = priorValue;
          applySelectAutosize(sel);
        }

        async function fetchLookupListChoices(pw, fieldMeta) {
          let listId = fieldMeta && fieldMeta.LookupList;
          if (!listId) return [];
          listId = String(listId).replace(/[{}' ]/g, "");
          const lookupField = (fieldMeta.LookupField || "Title").trim();
          return fetchListItemChoices(pw, "", listId, lookupField);
        }

        async function resolveBylawCertifierWriteKey(seg, pw) {
          try {
            const r = await fetchChoiceOptions(seg, pw, "Certifier");
            const internal = String((r.field && r.field.InternalName) || "Certifier").trim();
            const type = String((r.field && r.field.TypeAsString) || "");
            if (/lookup/i.test(type)) {
              return internal.endsWith("Id") ? internal : internal + "Id";
            }
            return internal;
          } catch (_) {
            return "CertifierId";
          }
        }

        async function fillBylawCertifierDropdown(sel, seg, pw, sampleRow) {
          if (!sel || sel.tagName !== "SELECT") return;
          const certCol = normalizedBylawTrainingColumns().find(function (c) {
            return c.key === "Certifier";
          });
          if (!certCol) return;

          const writeKey = await resolveBylawCertifierWriteKey(seg, pw);
          const lookupBase = writeKey.replace(/Id$/, "");
          sel.dataset.writeKey = writeKey;
          if (lookupBase) sel.dataset.lookupDisplayKey = lookupBase;

          const entries = await fetchCertifiersListEntries(pw);
          if (entries.length) {
            appendCertifierSelectOptions(sel, entries);
            return;
          }

          await fillDropdownSelect(sel, certCol, seg, pw, sampleRow, null);
          if (selectHasRealChoices(sel)) return;

          const listSeg = certifiersListApiSegment();
          const errText = listSeg
            ? "(no certifiers in list " + LIST_CERTIFIERS + " - check Certifier column and list permissions)"
            : "(Certifiers list not configured)";
          while (sel.options.length > 1) sel.remove(1);
          const o = document.createElement("option");
          o.value = "";
          o.textContent = errText;
          o.disabled = true;
          sel.appendChild(o);
          applySelectAutosize(sel);
          log("Certifier dropdown: no rows from " + LIST_CERTIFIERS + " (segment " + listSeg + ")", "warn");
        }

        async function fetchChoiceOptions(seg, pw, fieldInternalOrTitle) {
          const esc = String(fieldInternalOrTitle || "").replace(/'/g, "''");
          if (!esc) return { choices: [], error: "empty field name" };
          const data = await spFetch(
            `/_api/web/${seg}/fields/getbyinternalnameortitle('${esc}')?$select=Choices,Title,InternalName,TypeAsString,LookupList,LookupField`,
            {},
            pw,
          );
          let choices = [];
          if (Array.isArray(data.Choices)) choices = data.Choices;
          else if (data.Choices && typeof data.Choices === "object" && Array.isArray(data.Choices.results))
            choices = data.Choices.results;
          choices = choices.map((c) => String(c).trim()).filter(Boolean);
          let lookupEntries = [];
          if (!choices.length && data.TypeAsString && /lookup/i.test(String(data.TypeAsString))) {
            lookupEntries = await fetchLookupListChoices(pw, data);
            choices = lookupEntries.map(function (e) {
              return e.text;
            });
          }
          return { choices: choices, lookupEntries: lookupEntries, field: data };
        }

        function choiceFieldCandidateNames(col, sampleRow) {
          const seen = new Set();
          const out = [];
          function add(k) {
            const s = String(k || "").trim();
            if (!s || seen.has(s)) return;
            seen.add(s);
            out.push(s);
          }
          add(resolveWriteKey(col, sampleRow));
          add(col.key);
          if (col.saveKey) add(col.saveKey);
          (col.altKeys || []).forEach(add);
          (col.tryKeys || []).forEach(add);
          if (sampleRow && typeof sampleRow === "object") {
            Object.keys(sampleRow).forEach(function (k) {
              if (isODataOrMetaKey(k)) return;
              const ck = String(col.key).toLowerCase();
              const kl = String(k).toLowerCase();
              if (k === col.key || kl === ck || kl.indexOf(ck) === 0) add(k);
            });
          }
          return out;
        }

        async function fetchChoiceOptionsForColumn(col, seg, pw, sampleRow) {
          const candidates = choiceFieldCandidateNames(col, sampleRow);
          for (let i = 0; i < candidates.length; i++) {
            try {
              const r = await fetchChoiceOptions(seg, pw, candidates[i]);
              if ((r.choices && r.choices.length) || (r.lookupEntries && r.lookupEntries.length)) {
                return {
                  choices: r.choices || [],
                  lookupEntries: r.lookupEntries || [],
                  fieldKey: candidates[i],
                  field: r.field,
                };
              }
            } catch (_) {}
          }
          try {
            const data = await spFetch(
              `/_api/web/${seg}/fields?$select=Title,InternalName,Choices,TypeAsString,LookupList,LookupField&$filter=Hidden eq false&$top=200`,
              {},
              pw,
            );
            const fields = (data && data.value) || [];
            const label = String(col.label || col.key).trim().toLowerCase();
            const keyLower = String(col.key).trim().toLowerCase();
            for (let j = 0; j < fields.length; j++) {
              const f = fields[j];
              const title = String(f.Title || "").trim().toLowerCase();
              const internal = String(f.InternalName || "").trim().toLowerCase();
              if (title !== label && internal !== keyLower && title !== keyLower) continue;
              if (f.TypeAsString && /lookup/i.test(String(f.TypeAsString)) && f.LookupList) {
                const lookupEntries = await fetchLookupListChoices(pw, f);
                if (lookupEntries.length) {
                  return {
                    choices: lookupEntries.map(function (e) {
                      return e.text;
                    }),
                    lookupEntries: lookupEntries,
                    fieldKey: f.InternalName || f.Title,
                    field: f,
                  };
                }
              }
              let choices = [];
              if (Array.isArray(f.Choices)) choices = f.Choices;
              else if (f.Choices && typeof f.Choices === "object" && Array.isArray(f.Choices.results))
                choices = f.Choices.results;
              choices = choices.map(function (c) {
                return String(c).trim();
              }).filter(Boolean);
              if (choices.length) {
                return { choices: choices, fieldKey: f.InternalName || f.Title, field: f };
              }
            }
          } catch (_) {}
          return { choices: [], fieldKey: candidates[0] || col.key };
        }

        function formFieldPayloadValue(el, writeKey) {
          const v = String(el.value || "").trim();
          if (v === "") return null;
          if (el.type === "date") return v + "T00:00:00Z";
          if (el.tagName === "SELECT" && /Id$/.test(writeKey)) {
            const n = parseInt(v, 10);
            return isNaN(n) ? v : n;
          }
          return v;
        }

        function ensureSelectIncludesValue(sel, value) {
          if (!sel || sel.tagName !== "SELECT") return;
          const v = String(value != null ? value : "").trim();
          if (!v) return;
          for (let i = 0; i < sel.options.length; i++) {
            if (sel.options[i].value === v) return;
          }
          const o = document.createElement("option");
          o.value = v;
          o.textContent = v;
          sel.appendChild(o);
        }

        async function fillDropdownSelect(sel, col, seg, pw, sampleRow, item) {
          if (!sel || sel.tagName !== "SELECT") return;
          const priorValue = String(sel.value || "").trim();
          while (sel.options.length > 1) sel.remove(1);
          let loaded = false;
          try {
            const r = await fetchChoiceOptionsForColumn(col, seg, pw, sampleRow);
            const fieldMeta = r.field || {};
            const internal = fieldMeta.InternalName || r.fieldKey || col.key;
            if (r.lookupEntries && r.lookupEntries.length) {
              sel.dataset.writeKey = internal + "Id";
              sel.dataset.lookupDisplayKey = internal;
              r.lookupEntries.forEach(function (entry) {
                const o = document.createElement("option");
                o.value = String(entry.id);
                o.textContent = entry.text;
                sel.appendChild(o);
              });
              loaded = true;
            } else if (r.fieldKey) {
              sel.dataset.writeKey = r.fieldKey;
              delete sel.dataset.lookupDisplayKey;
            }
            if (!loaded && r.choices && r.choices.length) {
              r.choices.forEach(function (c) {
                const o = document.createElement("option");
                o.value = c;
                o.textContent = c;
                sel.appendChild(o);
              });
              loaded = true;
            } else if (!loaded) {
              const o = document.createElement("option");
              o.value = "";
              o.textContent = "(no choices - tried: " + choiceFieldCandidateNames(col, sampleRow).join(", ") + ")";
              o.disabled = true;
              sel.appendChild(o);
            }
          } catch (e) {
            log("Choices for " + col.key + ": " + (e.message || String(e)), "err");
            const o = document.createElement("option");
            o.value = "";
            o.textContent = "(error loading choices)";
            o.disabled = true;
            sel.appendChild(o);
          }
          if (priorValue) sel.value = priorValue;
          else if (item) setFieldValueFromItem(sel, col, item);
          ensureSelectIncludesValue(sel, sel.value);
          applySelectAutosize(sel);
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
          const notesTa = document.getElementById("nf_Notes");
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

          const form = document.createElement("form");
          form.id = "newPersonForm";
          form.className = "add-form-stack";
          form.addEventListener("submit", (ev) => {
            ev.preventDefault();
            void submitNewPerson(meta, pw, seg, sampleRow);
          });

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
          submitBtn.type = "submit";
          submitBtn.className = "btn";
          submitBtn.textContent = "Submit";
          bar.appendChild(submitBtn);
          const clearBtn = document.createElement("button");
          clearBtn.type = "button";
          clearBtn.className = "btn-secondary";
          clearBtn.textContent = "Clear form";
          clearBtn.addEventListener("click", () => {
            form.reset();
            primeRecordDateDefault();
          });
          bar.appendChild(clearBtn);
          form.appendChild(bar);
          panel.appendChild(form);

          for (let i = 0; i < cols.length; i++) {
            const col = cols[i];
            if (col.key === "Id") continue;
            if (DROPDOWN_COLUMN_KEYS.indexOf(col.key) === -1) continue;
            const sel = document.getElementById("nf_" + col.key);
            await fillDropdownSelect(sel, col, seg, pw, sampleRow, sampleRow);
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
          /** With `application/json;odata=nometadata`, do not send `__metadata` - SharePoint treats it as a list column and returns 400. */
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
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });

          if (SET_TITLE_ON_CREATE) {
            const title = [lastName, firstName].filter(Boolean).join(", ") || dod || "New Personnel";
            payload.Title = title;
          }

          try {
            setReadState("loading", "Submitting...");
            const created = await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            const newId = parseCreatedItemId(created);
            log("POST created item Id: " + (newId || "?"), "ok");
            const form = document.getElementById("newPersonForm");
            if (form) form.reset();
            primeRecordDateDefault();
            await runProbe();
            setReadState("ok", newId ? "Submitted (Id " + newId + "). Table refreshed." : "Submitted. Table refreshed.");
          } catch (e) {
            setReadState("err", "Submit failed: " + (e.message || String(e)).slice(0, 420));
            log("POST failed:\n" + (e.message || String(e)), "err");
          }
        }

        async function runProbe() {
          out.innerHTML = "";
          clearRosterTable();
          const addPanel = document.getElementById("addPersonPanel");
          if (addPanel) addPanel.innerHTML = "";
          setReadState("loading", "Loading Personnel list from SharePoint...");

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
          log("Page web from context (informational only):\n" + (pageWeb || "(no _spPageContextInfo - OK if explicit URLs are set)"));
          log(
            "REST base used for ALL API calls below:\n" +
              pw +
              (pageWeb && pw === pageWeb
                ? "  (same as page - from context or your URL matched the page site)"
                : "  (explicit URL / derived from PERSONNEL_SHAREPOINT_URL - not necessarily this page's site)"),
          );
          log("List API segment:\n/_api/web/" + seg);
          log("LIST_PERSONNEL (trimmed):\n" + JSON.stringify(personnelListTitle()));

          let meta = null;
          try {
            log("--- Request: GET list metadata ---", "ok");
            meta = await spFetch(`/_api/web/${seg}?$select=Id,Title,ItemCount,ListItemEntityTypeFullName`, {}, pw);
            log(meta);
          } catch (e) {
            log("LIST METADATA FAILED:\n" + (e.message || String(e)), "err");
            setReadState("err", "Cannot read this list (metadata request failed). Open Diagnostics for details and list names.");
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
              log("orderby failed with 400; retry without orderby...", "err");
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
            log("List is empty (ItemCount may still be > 0 if filtered view - this is default view items).");
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

          hubSession = {
            rows: rows,
            meta: meta,
            pw: pw,
            seg: seg2,
            sampleRow: sampleRow,
          };
          invalidateReportsTrainingCache();

          await renderAddPersonForm(meta, pw, seg2, sampleRow);
          renderRosterTable(rows, meta, pw, seg2);
          setHubListViewVisible(true);
        }

        btn.addEventListener("click", () => {
          runProbe().catch((e) => {
            setReadState("err", "Unexpected error: " + (e.message || String(e)).slice(0, 280));
            log("UNCAUGHT:\n" + (e.message || String(e)), "err");
          });
        });

        function initHubAccessGate() {
          const gate = document.getElementById("hubAccessGate");
          const input = document.getElementById("hubAccessInput");
          const submit = document.getElementById("hubAccessSubmit");
          const err = document.getElementById("hubAccessError");
          const password = String(HUB_ACCESS_PASSWORD || "").trim();

          function isHubAccessGranted() {
            try {
              return sessionStorage.getItem(HUB_ACCESS_STORAGE_KEY) === "1";
            } catch (_) {
              return false;
            }
          }

          function grantHubAccess() {
            try {
              sessionStorage.setItem(HUB_ACCESS_STORAGE_KEY, "1");
            } catch (_) {}
          }

          function startHubApp() {
            runProbe().catch(function (e) {
              setReadState("err", "Unexpected error: " + (e.message || String(e)).slice(0, 280));
              log("UNCAUGHT:\n" + (e.message || String(e)), "err");
            });
          }

          if (!password) {
            if (gate) gate.hidden = true;
            startHubApp();
            return;
          }

          function tryUnlock() {
            if (!input) return;
            if (String(input.value || "") === password) {
              grantHubAccess();
              if (err) err.hidden = true;
              if (gate) gate.hidden = true;
              startHubApp();
              return;
            }
            if (err) {
              err.hidden = false;
              err.textContent = "Incorrect password.";
            }
            input.value = "";
            input.focus();
          }

          if (isHubAccessGranted()) {
            if (gate) gate.hidden = true;
            startHubApp();
            return;
          }

          if (gate) gate.hidden = false;
          if (input) {
            input.value = "";
            window.setTimeout(function () {
              input.focus();
            }, 0);
          }
          if (submit) submit.addEventListener("click", tryUnlock);
          if (input) {
            input.addEventListener("keydown", function (ev) {
              if (ev.key === "Enter") {
                ev.preventDefault();
                tryUnlock();
              }
            });
          }
        }

        if (window.__trainingHubSkipAccessGate) {
          const gateEarly = document.getElementById("hubAccessGate");
          if (gateEarly) gateEarly.hidden = true;
          runProbe().catch(function (e) {
            setReadState("err", "Unexpected error: " + (e.message || String(e)).slice(0, 280));
            log("UNCAUGHT:\n" + (e.message || String(e)), "err");
          });
        } else {
          initHubAccessGate();
        }
      })();
