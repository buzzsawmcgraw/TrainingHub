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
        /**
         * Official AF letterhead is maintained as a Word/PDF template (not in a SharePoint list).
         * Upload the squadron letterhead PDF to Site Assets and set the path below.
         * Printed memos leave top margin for letterhead; use the template link to print letterhead PDF separately.
         */
        const MEMO_LETTERHEAD_PDF_URL = "";
        const MEMO_LETTERHEAD_PDF_SITE_RELATIVE_PATH = "S3T%20Files/AF-Letterhead.pdf";
        const DEFAULT_MEMO_PRINT_TOP_MARGIN_IN = 1.35;
        const MEMO_LETTERHEAD_STORAGE_KEY = "trainingHub.memoLetterhead.v3";
        const DEFAULT_MEMO_LETTERHEAD_FOOTER = "";
        /**
         * Optional MQL signature store (single row). Columns: MqlSignatureName, MqlSignatureTitle.
         * Create on the site as **HubLetterhead** so all users share the same commander signature block.
         */
        const LIST_HUB_LETTERHEAD = "HubLetterhead";
        const LIST_HUB_LETTERHEAD_GUID = "";
        const CERTIFIERS_PERSON_FIELD = "PersonnelId";
        const CERTIFIERS_LIST_DISPLAY_FIELD = "Certifier";
        /** Set SharePoint Title on new rows from the Item value when the list still requires Title. */
        const BYLAW_TRAINING_SET_TITLE = true;
        const BYLAW_TRAINING_ITEM_SORT_KEYS = ["Item", "Item0", "Title", "Training", "ByLaw", "TrainingName", "Training_x0020_Name"];

        /**
         * Appointments list (SharePoint title **Appointments**). Rows are filtered by PeronnelId = personnel list item Id.
         * List columns: PeronnelId, AppointmentDateTime, Location, ProviderOffice (Location), ReasonType,
         * InstructorInitials, MissedAppointment, CreatedAt. PeronnelId is filter-only (not shown on record page).
         */
        const LIST_APPOINTMENTS = "Appointments";
        const LIST_APPOINTMENTS_GUID = "";
        /** Past appointments are moved here automatically (same columns as Appointments plus ArchivedAt). */
        const LIST_APPOINTMENTS_ARCHIVE = "AppointmentsArchive";
        const LIST_APPOINTMENTS_ARCHIVE_GUID = "";
        const APPOINTMENTS_PERSON_FIELD = "PeronnelId";
        /** Archive list uses correct spelling (Appointments list has the PeronnelId typo). */
        const APPOINTMENTS_ARCHIVE_PERSON_FIELD = "PersonnelId";
        const APPOINTMENTS_ARCHIVE_PERSON_FIELD_ALT = [
          "PersonnelID",
          "PersonnelIdId",
          "Personnel_x0020_Id",
          "Personnel_x0020_ID",
        ];
        /** Additional REST filter field names to try if the primary name fails (lookup / casing variants). */
        const APPOINTMENTS_PERSON_FIELD_ALT = [
          "PersonnelId",
          "PersonnelID",
          "PersonnelIdId",
          "PeronnelIdId",
          "Personnel_x0020_Id",
          "Personnel_x0020_ID",
          "Peronnel_x0020_Id",
          "Personnel/Id",
        ];
        const APPOINTMENTS_ITEMS_ORDERBY = "AppointmentDateTime asc";
        const APPOINTMENTS_ARCHIVE_ITEMS_ORDERBY = "AppointmentDateTime desc";
        /** Archived rows are deleted after this many days (from ArchivedAt). */
        const APPOINTMENTS_ARCHIVE_RETENTION_DAYS = 30;
        /** Highlight archived rows when this many days or fewer remain before deletion. */
        const APPOINTMENTS_ARCHIVE_WARN_YELLOW_DAYS = 10;
        const APPOINTMENTS_ARCHIVE_WARN_RED_DAYS = 5;
        const APPOINTMENTS_COLUMNS = [
          {
            key: "AppointmentDateTime",
            label: "Appointment date/time",
            altKeys: [
              "Appointment_x0020_Date_x0020_Time",
              "Appointment_x0020_DateTime",
              "AppointmentTimeDate",
              "AppointmentTime_x0020_Date",
              "AppointmentDate",
              "Appointment_x0020_Date",
              "Date",
            ],
          },
          { key: "Location", label: "Facility", altKeys: ["Location0"] },
          {
            key: "ProviderOffice",
            label: "Location",
            altKeys: [
              "ProviderOffice_x0020__x0028_Location_x0029_",
              "Provider_x0020_Office",
              "ProviderOffice0",
            ],
          },
          {
            key: "ReasonType",
            label: "Description",
            altKeys: ["Reason_x0020_Type", "ReasonType0"],
          },
          {
            key: "InstructorInitials",
            label: "Instructor initials",
            altKeys: ["Instructor_x0020_Initials", "InstructorInitials0"],
          },
          {
            key: "MissedAppointment",
            label: "Missed appointment",
            altKeys: ["Missed_x0020_Appointment", "MissedAppointment0"],
          },
          {
            key: "CreatedAt",
            label: "Created at",
            altKeys: ["Created", "Created_x0020_At"],
          },
        ];
        const LIST_PHASE_ONE_TRACKING = "PhaseOneTracking";
        const LIST_PHASE_ONE_TRACKING_GUID = "";
        const LIST_PHASE_ONE_ARCHIVE = "PhaseOneArchive";
        const LIST_PHASE_ONE_ARCHIVE_GUID = "";
        const PHASE_ONE_TRACKING_PERSON_FIELD = "PersonnelId";
        const PHASE_ONE_TRACKING_ITEMS_ORDERBY = "PhaseOneStartDate desc";
        const PHASE_ONE_ARCHIVE_ITEMS_ORDERBY = "ArchivedAt desc";
        /** Calendar days of house-hunting leave before Phase 1 starts (weekends/federal holidays adjusted - see phaseOneStartDateAfterHouseHunting). */
        const PHASE_ONE_HOUSE_HUNTING_DAYS = 10;
        const PHASE_ONE_DUTY_STATUS_HOUSE_HUNTING = "House Hunting";
        const PHASE_ONE_DUTY_STATUS_ACTIVE = "Phase 1";
        const PHASE_ONE_TRACKING_COLUMNS = [
          { key: "PersonnelName", label: "Personnel" },
          { key: "DutyStatus", label: "Duty status" },
          { key: "DutySection", label: "Duty section" },
          { key: "DateArrived", label: "Date arrived" },
          { key: "PhaseOneStartDate", label: "Phase 1 start" },
          { key: "ProjectedOfficeFlight", label: "Projected Office/Flight" },
          { key: "DateReleased", label: "Date released" },
          { key: "Notes", label: "Notes" },
        ];
        const PHASE_ONE_ARCHIVE_COLUMNS = [
          { key: "PersonnelName", label: "Personnel" },
          { key: "DutyStatus", label: "Duty status" },
          { key: "DutySection", label: "Duty section" },
          { key: "DateArrived", label: "Date arrived" },
          { key: "PhaseOneStartDate", label: "Phase 1 start" },
          { key: "ProjectedOfficeFlight", label: "Projected Office/Flight" },
          { key: "DateReleased", label: "Date released" },
          { key: "ArchivedAt", label: "Archived" },
          { key: "Notes", label: "Notes" },
        ];
        const PHASE_ONE_EDIT_FIELDS = [
          { key: "DutyStatus", label: "Duty status", inputType: "text" },
          { key: "DutySection", label: "Duty section", inputType: "text" },
          { key: "PhaseOneStartDate", label: "Phase 1 start date", inputType: "date" },
          { key: "ProjectedOfficeFlight", label: "Projected Office/Flight", inputType: "select" },
          { key: "DateReleased", label: "Date released", inputType: "date" },
          { key: "Notes", label: "Notes", inputType: "textarea" },
        ];

        const APPOINTMENTS_SQUADRON_LABEL = "88 SFS";
        const AF_REPORT_UNIT_LABEL = "88th Security Forces Squadron";
        const MQL_REPORT_PRINT_TITLE = "Weapons Qual / Authorization Report";
        const MQL_SIGNED_SORT_LINE = "by Duty Status / Last Name";
        const MQL_EXPORT_SORT_LINE = "Alphabetical by Last Name";
        const MQL_SIGNATURE_STORAGE_KEY = "trainingHub.mqlSignature.v1";
        const DEFAULT_MQL_SIGNATURE_NAME = "Name, Rank, USAF";
        const DEFAULT_MQL_SIGNATURE_TITLE = "Commander";
        /** Always shown on standard MQL report/PDF (qual / exp per column). */
        const MQL_REQUIRED_WEAPON_LABELS = ["M4", "M18"];
        /** By-law columns omitted from all MQL reports (case-insensitive partial match). */
        const MQL_EXCLUDED_CATALOG_LABELS = ["airfield driving", "combatives", "cqb"];
        const MQL_EXPORT_PRINT_TITLE = "MQL Data Export";
        const MQL_PRINT_STYLE_ID = "hubMqlLandscapePrintStyle";
        /** Fixed landscape columns for standard MQL report and data export. */
        const MQL_STANDARD_COLUMN_DEFS = [
          { label: "Duty Status", kind: "person", field: "status" },
          { label: "Duty Section", kind: "person", field: "section" },
          { label: "Full Name", kind: "person", field: "fullName" },
          { label: "DoDID", kind: "person", field: "dodid" },
          { label: "M18 Qual", kind: "cert", item: "M18", part: "qual", preferWeapon: true },
          { label: "M18 Exp", kind: "cert", item: "M18", part: "exp", preferWeapon: true },
          { label: "M4 Qual", kind: "cert", item: "M4", part: "qual", preferWeapon: true },
          { label: "M4 Exp", kind: "cert", item: "M4", part: "exp", preferWeapon: true },
          { label: "Baton Qual", kind: "cert", item: "Baton", part: "qual", preferWeapon: true, match: ["Baton"] },
          { label: "Baton Exp", kind: "cert", item: "Baton", part: "exp", preferWeapon: true, match: ["Baton"] },
          { label: "Taser Qual", kind: "cert", item: "Taser", part: "qual", preferWeapon: true, match: ["Taser", "TASER"] },
          { label: "Taser Exp", kind: "cert", item: "Taser", part: "exp", preferWeapon: true, match: ["Taser", "TASER"] },
          {
            label: "DD2760 Exp",
            kind: "cert",
            item: "DD2760",
            part: "exp",
            preferWeapon: false,
            match: ["DD Form 2760", "DD 2760", "DD2760", "Form 2760", "2760"],
          },
          {
            label: "AUOF Exp",
            kind: "cert",
            item: "AUOF",
            part: "exp",
            preferWeapon: false,
            match: ["AUOF", "Annual Use of Force", "Use of Force"],
          },
        ];
        /**
         * Future hub footer quick links (render when populated). Example:
         * { label: "SharePoint", url: "https://..." }
         */
        const HUB_EXTERNAL_LINKS = [];
        const APPOINTMENTS_MEMO_DEFAULT_FROM = "88 SFS/S3T";
        const MEMO_SIGNATURE_TITLE_PRESETS = [
          "Training NCOIC",
          "Supervisor",
          "Commander",
          "First Sergeant",
          "Superintendent",
        ];
        const APPOINTMENTS_FORM_FIELDS = [
          {
            key: "AppointmentDateTime",
            label: "Appointment date/time",
            inputType: "datetime-local",
            required: true,
            altKeys: [
              "Appointment_x0020_Date_x0020_Time",
              "Appointment_x0020_DateTime",
              "AppointmentTimeDate",
            ],
          },
          { key: "Location", label: "Facility", required: false, altKeys: ["Location0"] },
          {
            key: "ProviderOffice",
            label: "Location",
            required: false,
            altKeys: [
              "ProviderOffice_x0020__x0028_Location_x0029_",
              "Provider_x0020_Office",
              "ProviderOffice0",
            ],
          },
          {
            key: "ReasonType",
            label: "Description",
            required: false,
            altKeys: ["Reason_x0020_Type", "ReasonType0"],
          },
          {
            key: "InstructorInitials",
            label: "Instructor initials",
            required: false,
            altKeys: ["Instructor_x0020_Initials", "InstructorInitials0"],
          },
          {
            key: "MissedAppointment",
            label: "Missed appointment",
            inputType: "checkbox",
            required: false,
            altKeys: ["Missed_x0020_Appointment", "MissedAppointment0"],
          },
        ];

        /** ETHOS Members list (SharePoint title **EthosMembers**). */
        const LIST_ETHOS_MEMBERS = "EthosMembers";
        const LIST_ETHOS_MEMBERS_GUID = "";
        const ETHOS_ITEMS_ORDERBY = "LastName asc,FirstName asc";
        const ETHOS_ROSTER_COLUMNS = [
          { key: "Rank", label: "Rank", altKeys: ["Rank0", "PayGrade"] },
          { key: "LastName", label: "LastName" },
          { key: "FirstName", label: "FirstName" },
          { key: "MiddleInitial", label: "Middle Initial" },
          { key: "Status", label: "Status" },
          { key: "OfficeSymbol", label: "OfficeSymbol" },
          { key: "DoDID", label: "DoDID" },
          { key: "Squadron", label: "Squadron" },
        ];
        const ETHOS_ADD_FORM_FIELD_GROUPS = [
          {
            title: "Identity",
            keys: ["Rank", "LastName", "FirstName", "MiddleInitial", "DoDID"],
            split: "left",
            layout: "stack",
          },
          {
            title: "Record",
            keys: ["Status", "OfficeSymbol", "Squadron"],
            split: "right",
            layout: "stack",
          },
        ];
        const ETHOS_DROPDOWN_COLUMN_KEYS = ["Status", "OfficeSymbol", "Squadron", "Rank"];
        const ETHOS_SET_TITLE_ON_CREATE = true;
        const ETHOS_ROSTER_SCROLL_AFTER_ROWS = 10;

        const ETHOS_WEAPONS_PERSON_FIELD = "EthosMemberId";
        const ETHOS_WEAPONS_PERSON_FIELD_ALT = [
          "EthosMemberID",
          "EthosMemberIdId",
          "EthosMember_x0020_Id",
          "EthosMember_x0020_ID",
          "EthosMember/Id",
        ];
        const LIST_ETHOS_WEAPONS_CERTIFICATIONS = "EthosWeaponsCertifications";
        const LIST_ETHOS_WEAPONS_CERTIFICATIONS_GUID = "";
        const ETHOS_WEAPONS_CERT_ITEMS_ORDERBY = "QualDate desc";
        const ETHOS_WEAPONS_CERT_COLUMNS = WEAPONS_CERT_COLUMNS;
        const ETHOS_WEAPONS_CERT_DROPDOWN_KEYS = ["Weapon"];
        const ETHOS_WEAPONS_CERT_SET_TITLE = true;

        const ETHOS_BYLAW_PERSON_FIELD = "EthosMemberId";
        const ETHOS_BYLAW_PERSON_FIELD_ALT = [
          "EthosMemberID",
          "EthosMemberIdId",
          "EthosMember_x0020_Id",
          "EthosMember_x0020_ID",
          "EthosMember/Id",
        ];
        const LIST_ETHOS_BYLAW_TRAINING = "EthosByLawTraining";
        const LIST_ETHOS_BYLAW_TRAINING_GUID = "";
        const ETHOS_BYLAW_TRAINING_ITEMS_ORDERBY = "QualDate desc";
        const ETHOS_BYLAW_TRAINING_COLUMNS = BYLAW_TRAINING_COLUMNS;
        const ETHOS_BYLAW_TRAINING_DROPDOWN_KEYS = BYLAW_TRAINING_DROPDOWN_KEYS;
        const ETHOS_BYLAW_TRAINING_SET_TITLE = true;
        const ETHOS_BYLAW_TRAINING_ITEM_SORT_KEYS = BYLAW_TRAINING_ITEM_SORT_KEYS;
        const ETHOS_SQUADRON_LABEL = "88 SFS";

        /** Monthly Status of Training print: months shown in the dropdown (including current month). */
        const SOT_REPORT_MONTHS_BACK = 24;
        /**
         * Personnel Status values included in training posture counts (OfficeSymbol groups).
         * Leave empty [] to include every status. Example: ["Active", "AGR"].
         */
        const SOT_PERSONNEL_STATUS_FOR_COUNTS = [];
        /** Printed monthly report header (matches legacy SOT PDF). */
        const SOT_SQUADRON_LABEL = "88 SFS";

        /** Hub Reports menu (Reports nav). Set implemented:true when the report renderer is wired. */
        const HUB_REPORT_TYPES = [
          {
            id: "status-of-training",
            title: "Status of Training",
            subtitle:
              "Active snapshot with Personnel names, or a monthly office summary for printing (no names).",
            badge: "SOT",
            implemented: true,
            printable: true,
          },
          {
            id: "mql-report",
            title: "MQL Report",
            subtitle: "Signed commander report - M4, M18, and By-Law qualifications by section.",
            badge: "MQL",
            implemented: true,
            printable: true,
          },
          {
            id: "mql-pdf",
            title: "MQL PDF",
            subtitle: "Export layout with separate qual/expiration columns for Excel and email.",
            badge: "PDF",
            implemented: true,
            printable: true,
          },
          {
            id: "heavy-weapons",
            title: "Heavy Weapons MQL",
            subtitle: "MQL layout for all weapons except M4 and M18, plus By-Law training.",
            badge: "HVY",
            implemented: true,
            printable: true,
          },
          {
            id: "cleo-report",
            title: "CLEO Report",
            subtitle: "CLEO-designated Personnel and the date each became CLEO.",
            badge: "CLEO",
            implemented: true,
            printable: true,
          },
          {
            id: "post-rc-report",
            title: "Post RC Report",
            subtitle: "Personnel with Post RC on file - qualification and expiration dates only.",
            badge: "PRC",
            implemented: true,
            printable: true,
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

        /**
         * Personnel list CLEO / Post RC fields (not shown on main roster table).
         * Create on **Personnel**: IsCLEO (Yes/No), CLEODate, PostRCQualDate, PostRCExpirationDate.
         */
        const PERSONNEL_CLEO_COLUMNS = [
          {
            key: "IsCLEO",
            label: "CLEO",
            inputType: "checkbox",
            altKeys: ["Is_x0020_CLEO", "CLEO", "Cleo", "IsCLEO0"],
          },
          {
            key: "CLEODate",
            label: "Date became CLEO",
            inputType: "date",
            altKeys: [
              "CLEO_x0020_Date",
              "CLEODate0",
              "DateBecameCLEO",
              "Date_x0020_Became_x0020_CLEO",
              "CLEOEffectiveDate",
            ],
          },
          {
            key: "PostRCQualDate",
            label: "Post RC qual date",
            inputType: "date",
            altKeys: [
              "PostRC_x0020_Qual_x0020_Date",
              "PostRCQual_x0020_Date",
              "PostRCQualDate0",
              "PostRCDate",
            ],
          },
          {
            key: "PostRCExpirationDate",
            label: "Post RC expiration",
            inputType: "date",
            readOnly: true,
            altKeys: [
              "PostRC_x0020_Expiration_x0020_Date",
              "PostRCExpiration_x0020_Date",
              "PostRCExpDate",
              "PostRCExpirationDate0",
            ],
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
        const personWeaponsBulkAddBtn = document.getElementById("personWeaponsBulkAddBtn");
        const personWeaponsBulkPanel = document.getElementById("personWeaponsBulkPanel");
        const personWeaponsBulkForm = document.getElementById("personWeaponsBulkForm");
        const personWeaponsBulkSharedFields = document.getElementById("personWeaponsBulkSharedFields");
        const personWeaponsBulkItems = document.getElementById("personWeaponsBulkItems");
        const personWeaponsBulkSelectAll = document.getElementById("personWeaponsBulkSelectAll");
        const personWeaponsBulkClearAll = document.getElementById("personWeaponsBulkClearAll");
        const personWeaponsBulkSaveBtn = document.getElementById("personWeaponsBulkSaveBtn");
        const personWeaponsBulkCancelBtn = document.getElementById("personWeaponsBulkCancelBtn");
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
        const personBylawBulkAddBtn = document.getElementById("personBylawBulkAddBtn");
        const personBylawBulkPanel = document.getElementById("personBylawBulkPanel");
        const personBylawBulkForm = document.getElementById("personBylawBulkForm");
        const personBylawBulkSharedFields = document.getElementById("personBylawBulkSharedFields");
        const personBylawBulkItems = document.getElementById("personBylawBulkItems");
        const personBylawBulkSelectAll = document.getElementById("personBylawBulkSelectAll");
        const personBylawBulkClearAll = document.getElementById("personBylawBulkClearAll");
        const personBylawBulkSaveBtn = document.getElementById("personBylawBulkSaveBtn");
        const personBylawBulkCancelBtn = document.getElementById("personBylawBulkCancelBtn");
        const personBylawAddPanel = document.getElementById("personBylawAddPanel");
        const personBylawAddForm = document.getElementById("personBylawAddForm");
        const personBylawAddFields = document.getElementById("personBylawAddFields");
        const personBylawAddCancelBtn = document.getElementById("personBylawAddCancelBtn");
        const personBylawSaveBtn = document.getElementById("personBylawSaveBtn");
        const personBylawFormTitle = document.getElementById("personBylawFormTitle");
        const personAppointmentsWrap = document.getElementById("personAppointmentsWrap");
        const personAppointmentsThead = document.getElementById("personAppointmentsThead");
        const personAppointmentsBody = document.getElementById("personAppointmentsBody");
        const personAppointmentsEmpty = document.getElementById("personAppointmentsEmpty");
        const personAppointmentsState = document.getElementById("personAppointmentsState");
        const hubNavInstructors = document.getElementById("hubNavInstructors");
        const hubNavPhase1Tracker = document.getElementById("hubNavPhase1Tracker");
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
        const reportsExportBtn = document.getElementById("reportsExportBtn");
        const reportsDetailTitle = document.getElementById("reportsDetailTitle");
        const reportsDetailSubtitle = document.getElementById("reportsDetailSubtitle");
        const reportsDetailBody = document.getElementById("reportsDetailBody");
        const hubExternalLinksBar = document.getElementById("hubExternalLinksBar");
        const hubNavScheduling = document.getElementById("hubNavScheduling");
        const hubNavEthos = document.getElementById("hubNavEthos");
        const ethosSection = document.getElementById("ethosSection");
        const ethosReadState = document.getElementById("ethosReadState");
        const ethosProbeRun = document.getElementById("ethosProbeRun");
        const ethosBackLink = document.getElementById("ethosBackLink");
        const ethosRosterTableBody = document.getElementById("ethosRosterTableBody");
        const ethosMqlReportBtn = document.getElementById("ethosMqlReportBtn");
        const ethosMqlPrintBtn = document.getElementById("ethosMqlPrintBtn");
        const ethosMqlCloseBtn = document.getElementById("ethosMqlCloseBtn");
        const ethosMqlReportPanel = document.getElementById("ethosMqlReportPanel");
        const ethosMqlReportBody = document.getElementById("ethosMqlReportBody");
        const ethosPrintSurface = document.getElementById("ethosPrintSurface");
        const ethosMemberDetailSection = document.getElementById("ethosMemberDetailSection");
        const ethosMemberDetailTitle = document.getElementById("ethosMemberDetailTitle");
        const ethosMemberDetailContent = document.getElementById("ethosMemberDetailContent");
        const ethosMemberDetailReadState = document.getElementById("ethosMemberDetailReadState");
        const ethosMemberDetailBackLink = document.getElementById("ethosMemberDetailBackLink");
        const ethosMemberDetailEditBtn = document.getElementById("ethosMemberDetailEditBtn");
        const ethosMemberDetailSaveBtn = document.getElementById("ethosMemberDetailSaveBtn");
        const ethosMemberDetailCancelBtn = document.getElementById("ethosMemberDetailCancelBtn");
        const ethosWeaponsWrap = document.getElementById("ethosWeaponsWrap");
        const ethosWeaponsThead = document.getElementById("ethosWeaponsThead");
        const ethosWeaponsBody = document.getElementById("ethosWeaponsBody");
        const ethosWeaponsEmpty = document.getElementById("ethosWeaponsEmpty");
        const ethosWeaponsState = document.getElementById("ethosWeaponsState");
        const ethosWeaponsAddBtn = document.getElementById("ethosWeaponsAddBtn");
        const ethosWeaponsBulkAddBtn = document.getElementById("ethosWeaponsBulkAddBtn");
        const ethosWeaponsBulkPanel = document.getElementById("ethosWeaponsBulkPanel");
        const ethosWeaponsBulkForm = document.getElementById("ethosWeaponsBulkForm");
        const ethosWeaponsBulkSharedFields = document.getElementById("ethosWeaponsBulkSharedFields");
        const ethosWeaponsBulkItems = document.getElementById("ethosWeaponsBulkItems");
        const ethosWeaponsBulkSelectAll = document.getElementById("ethosWeaponsBulkSelectAll");
        const ethosWeaponsBulkClearAll = document.getElementById("ethosWeaponsBulkClearAll");
        const ethosWeaponsBulkSaveBtn = document.getElementById("ethosWeaponsBulkSaveBtn");
        const ethosWeaponsBulkCancelBtn = document.getElementById("ethosWeaponsBulkCancelBtn");
        const ethosWeaponsAddPanel = document.getElementById("ethosWeaponsAddPanel");
        const ethosWeaponsAddForm = document.getElementById("ethosWeaponsAddForm");
        const ethosWeaponsAddFields = document.getElementById("ethosWeaponsAddFields");
        const ethosWeaponsAddCancelBtn = document.getElementById("ethosWeaponsAddCancelBtn");
        const ethosWeaponsSaveBtn = document.getElementById("ethosWeaponsSaveBtn");
        const ethosWeaponsFormTitle = document.getElementById("ethosWeaponsFormTitle");
        const ethosBylawWrap = document.getElementById("ethosBylawWrap");
        const ethosBylawThead = document.getElementById("ethosBylawThead");
        const ethosBylawBody = document.getElementById("ethosBylawBody");
        const ethosBylawEmpty = document.getElementById("ethosBylawEmpty");
        const ethosBylawState = document.getElementById("ethosBylawState");
        const ethosBylawAddBtn = document.getElementById("ethosBylawAddBtn");
        const ethosBylawBulkAddBtn = document.getElementById("ethosBylawBulkAddBtn");
        const ethosBylawBulkPanel = document.getElementById("ethosBylawBulkPanel");
        const ethosBylawBulkForm = document.getElementById("ethosBylawBulkForm");
        const ethosBylawBulkSharedFields = document.getElementById("ethosBylawBulkSharedFields");
        const ethosBylawBulkItems = document.getElementById("ethosBylawBulkItems");
        const ethosBylawBulkSelectAll = document.getElementById("ethosBylawBulkSelectAll");
        const ethosBylawBulkClearAll = document.getElementById("ethosBylawBulkClearAll");
        const ethosBylawBulkSaveBtn = document.getElementById("ethosBylawBulkSaveBtn");
        const ethosBylawBulkCancelBtn = document.getElementById("ethosBylawBulkCancelBtn");
        const ethosBylawAddPanel = document.getElementById("ethosBylawAddPanel");
        const ethosBylawAddForm = document.getElementById("ethosBylawAddForm");
        const ethosBylawAddFields = document.getElementById("ethosBylawAddFields");
        const ethosBylawAddCancelBtn = document.getElementById("ethosBylawAddCancelBtn");
        const ethosBylawSaveBtn = document.getElementById("ethosBylawSaveBtn");
        const ethosBylawFormTitle = document.getElementById("ethosBylawFormTitle");
        const schedulingSection = document.getElementById("schedulingSection");
        const schedulingReadState = document.getElementById("schedulingReadState");
        const schedulingBackLink = document.getElementById("schedulingBackLink");
        const schedulingAddBtn = document.getElementById("schedulingAddBtn");
        const schedulingPrintReportBtn = document.getElementById("schedulingPrintReportBtn");
        const schedulingAddPanel = document.getElementById("schedulingAddPanel");
        const schedulingAddForm = document.getElementById("schedulingAddForm");
        const schedulingAddFields = document.getElementById("schedulingAddFields");
        const schedulingPersonSelect = document.getElementById("schedulingPersonSelect");
        const schedulingAddCancelBtn = document.getElementById("schedulingAddCancelBtn");
        const schedulingSaveBtn = document.getElementById("schedulingSaveBtn");
        const schedulingThead = document.getElementById("schedulingThead");
        const schedulingTableBody = document.getElementById("schedulingTableBody");
        const schedulingEmpty = document.getElementById("schedulingEmpty");
        const schedulingPrintSurface = document.getElementById("schedulingPrintSurface");
        const schedulingPrintMissedReportBtn = document.getElementById("schedulingPrintMissedReportBtn");
        const schedulingNewMemoBtn = document.getElementById("schedulingNewMemoBtn");
        const schedulingTabAll = document.getElementById("schedulingTabAll");
        const schedulingTabMissed = document.getElementById("schedulingTabMissed");
        const schedulingTabArchive = document.getElementById("schedulingTabArchive");
        const phase1Section = document.getElementById("phase1Section");
        const phase1ReadState = document.getElementById("phase1ReadState");
        const phase1BackLink = document.getElementById("phase1BackLink");
        const phase1AddBtn = document.getElementById("phase1AddBtn");
        const phase1RefreshBtn = document.getElementById("phase1RefreshBtn");
        const phase1AddPanel = document.getElementById("phase1AddPanel");
        const phase1PersonSelect = document.getElementById("phase1PersonSelect");
        const phase1HouseHuntingYes = document.getElementById("phase1HouseHuntingYes");
        const phase1HouseHuntingNo = document.getElementById("phase1HouseHuntingNo");
        const phase1AddCancelBtn = document.getElementById("phase1AddCancelBtn");
        const phase1EditPanel = document.getElementById("phase1EditPanel");
        const phase1EditTitle = document.getElementById("phase1EditTitle");
        const phase1EditForm = document.getElementById("phase1EditForm");
        const phase1EditFields = document.getElementById("phase1EditFields");
        const phase1EditSaveBtn = document.getElementById("phase1EditSaveBtn");
        const phase1EditCancelBtn = document.getElementById("phase1EditCancelBtn");
        const phase1ArchiveBtn = document.getElementById("phase1ArchiveBtn");
        const phase1TabActive = document.getElementById("phase1TabActive");
        const phase1TabArchive = document.getElementById("phase1TabArchive");
        const phase1ListTitle = document.getElementById("phase1ListTitle");
        const phase1Thead = document.getElementById("phase1Thead");
        const phase1TableBody = document.getElementById("phase1TableBody");
        const phase1Empty = document.getElementById("phase1Empty");
        const schedulingListTitle = document.getElementById("schedulingListTitle");
        const schedulingMemoOverlay = document.getElementById("schedulingMemoOverlay");
        const schedulingMemoForm = document.getElementById("schedulingMemoForm");
        const schedulingMemoFor = document.getElementById("schedulingMemoFor");
        const schedulingMemoFrom = document.getElementById("schedulingMemoFrom");
        const schedulingMemoSubject = document.getElementById("schedulingMemoSubject");
        const schedulingMemoBody = document.getElementById("schedulingMemoBody");
        const schedulingMemoSig1Person = document.getElementById("schedulingMemoSig1Person");
        const schedulingMemoSig1Title = document.getElementById("schedulingMemoSig1Title");
        const schedulingMemoSig1Squadron = document.getElementById("schedulingMemoSig1Squadron");
        const schedulingMemoSig2Person = document.getElementById("schedulingMemoSig2Person");
        const schedulingMemoSig2Title = document.getElementById("schedulingMemoSig2Title");
        const schedulingMemoSig2Squadron = document.getElementById("schedulingMemoSig2Squadron");
        const schedulingMemoSig3Person = document.getElementById("schedulingMemoSig3Person");
        const schedulingMemoSig3Title = document.getElementById("schedulingMemoSig3Title");
        const schedulingMemoSig3Squadron = document.getElementById("schedulingMemoSig3Squadron");
        const schedulingMemoSignatureBlocks = [
          {
            personSelect: schedulingMemoSig1Person,
            titleSelect: schedulingMemoSig1Title,
            squadronSelect: schedulingMemoSig1Squadron,
          },
          {
            personSelect: schedulingMemoSig2Person,
            titleSelect: schedulingMemoSig2Title,
            squadronSelect: schedulingMemoSig2Squadron,
          },
          {
            personSelect: schedulingMemoSig3Person,
            titleSelect: schedulingMemoSig3Title,
            squadronSelect: schedulingMemoSig3Squadron,
          },
        ];
        const schedulingMemoPrintBtn = document.getElementById("schedulingMemoPrintBtn");
        const schedulingMemoCancelBtn = document.getElementById("schedulingMemoCancelBtn");
        const schedulingMemoLetterheadPdfLink = document.getElementById("schedulingMemoLetterheadPdfLink");
        const schedulingMemoPrintTopMargin = document.getElementById("schedulingMemoPrintTopMargin");
        const schedulingMemoLetterheadFooter = document.getElementById("schedulingMemoLetterheadFooter");
        const schedulingMemoLetterheadSaveBtn = document.getElementById("schedulingMemoLetterheadSaveBtn");

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
          appointmentsPersonFilterField: null,
          appointmentsPersonFilterFields: null,
          appointmentsPersonPostKey: null,
          appointmentsSampleRow: null,
          appointmentsArchivePersonPostKey: null,
          appointmentsArchiveSampleRow: null,
          memoLetterhead: null,
          memoLetterheadItemId: null,
          memoLetterheadLoaded: false,
          memoLetterheadUsesSharePoint: false,
          mqlSignature: null,
          mqlSignatureLoaded: false,
          mqlSignatureItemId: null,
          mqlSignatureUsesSharePoint: false,
        };

        let phase1Session = {
          listView: "active",
          trackingRows: null,
          archiveRows: null,
          editItem: null,
          personPostKey: null,
        };

        let schedulingSession = {
          rows: null,
          archiveRows: null,
          listView: "all",
        };

        let instructorsSession = {
          rows: null,
        };

        let reportsSession = {
          activeReportId: null,
          weaponsRows: null,
          bylawRows: null,
          mqlRows: null,
          mqlCatalog: null,
          mqlMode: "standard",
          bylawItemChoices: null,
          sotView: "active",
          selectedMonth: null,
          sotRows: null,
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

        let ethosSession = {
          rows: null,
          meta: null,
          pw: null,
          seg: null,
          sampleRow: null,
          weaponsCertRows: null,
          weaponsCertSampleRow: null,
          weaponsPersonFilterField: null,
          weaponsPersonFilterFields: null,
          weaponsPersonPostKey: null,
          bylawTrainingRows: null,
          bylawTrainingSampleRow: null,
          bylawPersonFilterField: null,
          bylawPersonFilterFields: null,
          mqlBylawByMember: null,
          mqlRows: null,
          mqlCatalog: null,
          mqlBylawItemChoices: null,
        };

        let ethosDetailSession = {
          item: null,
          editing: false,
          meta: null,
          pw: null,
          seg: null,
          sampleRow: null,
        };

        let ethosWeaponsCertEditSession = { item: null };
        let ethosBylawTrainingEditSession = { item: null };

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
          if (fieldKey === "MiddleInitial") {
            return ["MiddleInitial", "Middle_x0020_Initial", "MiddleInitial0", "MI"];
          }
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

        function hideEthosSections() {
          if (ethosSection) ethosSection.hidden = true;
          if (ethosMemberDetailSection) ethosMemberDetailSection.hidden = true;
          if (ethosMqlReportPanel) ethosMqlReportPanel.hidden = true;
          if (ethosMqlPrintBtn) ethosMqlPrintBtn.hidden = true;
          if (ethosMqlCloseBtn) ethosMqlCloseBtn.hidden = true;
        }

        function setEthosViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = true;
          if (phase1Section) phase1Section.hidden = true;
          if (ethosSection) ethosSection.hidden = !visible;
          if (ethosMemberDetailSection) ethosMemberDetailSection.hidden = true;
        }

        function hidePhase1Section() {
          if (phase1Section) phase1Section.hidden = true;
          if (phase1AddPanel) phase1AddPanel.hidden = true;
          if (phase1EditPanel) phase1EditPanel.hidden = true;
          phase1Session.editItem = null;
        }

        async function navigateToRoster() {
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = true;
          hidePhase1Section();
          hideEthosSections();
          clearEthosMemberDetailSection();
          setInstructorsAddPanelVisible(false);
          setSchedulingAddPanelVisible(false);
          setSchedulingMemoOverlayVisible(false);
          showReportsHub();
          personDetailSession = { item: null, editing: false, meta: null, pw: null, seg: null, sampleRow: null };
          ethosDetailSession = { item: null, editing: false, meta: null, pw: null, seg: null, sampleRow: null };
          clearPersonWeaponsCertSection();
          clearPersonBylawTrainingSection();
          clearPersonAppointmentsSection();
          setPersonDetailEditMode(false);
          setHubListViewVisible(true);
          await ensureRosterViewRendered();
        }

        function hubAlternateNavActive() {
          return (
            (ethosSection && !ethosSection.hidden) ||
            (ethosMemberDetailSection && !ethosMemberDetailSection.hidden) ||
            (reportsSection && !reportsSection.hidden) ||
            (schedulingSection && !schedulingSection.hidden) ||
            (phase1Section && !phase1Section.hidden) ||
            (instructorsSection && !instructorsSection.hidden) ||
            (personDetailSection && !personDetailSection.hidden)
          );
        }

        async function navigateToEthosRoster() {
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = true;
          hidePhase1Section();
          if (ethosMemberDetailSection) ethosMemberDetailSection.hidden = true;
          if (ethosMqlReportPanel) ethosMqlReportPanel.hidden = true;
          if (ethosMqlPrintBtn) ethosMqlPrintBtn.hidden = true;
          if (ethosMqlCloseBtn) ethosMqlCloseBtn.hidden = true;
          ethosDetailSession = { item: null, editing: false, meta: null, pw: null, seg: null, sampleRow: null };
          clearEthosMemberWeaponsSection();
          clearEthosMemberBylawSection();
          setEthosMemberDetailEditMode(false);
          setEthosViewVisible(true);
          const pw = resolvedPersonnelRestBase() || hubSession.pw || ethosSession.pw;
          if (!pw) {
            clearEthosRosterTable();
            const addPanel = document.getElementById("ethosAddMemberPanel");
            if (addPanel) addPanel.innerHTML = "";
            setEthosReadState(
              "warn",
              "Load the Personnel Roster first (wait for list to load or click Refresh list).",
            );
            return;
          }
          try {
            if (ethosSession.rows == null || !ethosSession.pw || !ethosSession.seg) {
              await runEthosProbe();
            } else {
              await ensureEthosRosterRendered();
            }
          } catch (e) {
            setEthosReadState("err", "Could not open ETHOS roster: " + (e.message || String(e)).slice(0, 280));
          }
        }

        async function navigateToEthos() {
          await navigateToEthosRoster();
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

        async function navigateToScheduling() {
          setSchedulingViewVisible(true);
          if (!hubSession.pw) {
            setSchedulingState(
              "warn",
              "Load the Personnel Roster first (wait for list to load or click Refresh list).",
            );
            return;
          }
          if (!appointmentsListTitle() && !appointmentsListUsesGuid()) {
            setSchedulingState("warn", "Appointments list is not configured.");
            return;
          }
          await loadSchedulingAppointmentsList();
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

        function setPhase1ViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = true;
          hideEthosSections();
          if (phase1Section) phase1Section.hidden = !visible;
          if (!visible) hidePhase1Section();
        }

        async function navigateToPhase1() {
          setPhase1ViewVisible(true);
          if (!hubSession.pw) {
            setPhase1State(
              "warn",
              "Load the Personnel Roster first (wait for list to load or click Refresh list).",
            );
            return;
          }
          if (!phaseOneTrackingListTitle() && !phaseOneTrackingListUsesGuid()) {
            setPhase1State("warn", "PhaseOneTracking list is not configured.");
            return;
          }
          await loadPhase1TrackerData();
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

        function sameCalendarMonth(a, b) {
          if (!a || !b || isNaN(a.getTime()) || isNaN(b.getTime())) return false;
          return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
        }

        function mqlExpiringNextCalendarMonth(expiry, today) {
          if (!expiry || !today) return false;
          const next = new Date(today.getFullYear(), today.getMonth() + 1, 1);
          return expiry.getFullYear() === next.getFullYear() && expiry.getMonth() === next.getMonth();
        }

        function computeMqlCertCellStatus(item, expiryKeys, qualKeys) {
          if (!item) return { tone: "expired" };
          const today = new Date();
          const qual = parseWeaponsCertCalendarDate(valueFromItemByKeys(item, qualKeys));
          const exp = resolveCertExpirationDate(item, expiryKeys, qualKeys);
          if (qual && sameCalendarMonth(qual, today)) return { tone: "recert" };
          const expRaw = parseWeaponsCertCalendarDate(valueFromItemByKeys(item, expiryKeys));
          if (!qual && !expRaw) return { tone: "expired" };
          if (!exp) return qual ? { tone: "ok" } : { tone: "expired" };
          if (calendarDaysBetween(today, exp) < 0) return { tone: "expired" };
          if (sameCalendarMonth(exp, today)) return { tone: "urgent" };
          if (mqlExpiringNextCalendarMonth(exp, today)) return { tone: "warn" };
          return { tone: "ok" };
        }

        function computeMqlWeaponsCertCellStatus(item) {
          return computeMqlCertCellStatus(item, weaponsCertExpiryDateKeys(), weaponsQualDateKeys());
        }

        function computeMqlBylawCertCellStatus(item) {
          return computeMqlCertCellStatus(item, bylawTrainingExpiryDateKeys(), bylawQualDateKeys());
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

        function normalizedPersonnelCleoColumns() {
          const arr = Array.isArray(PERSONNEL_CLEO_COLUMNS) ? PERSONNEL_CLEO_COLUMNS : [];
          return arr
            .map(function (entry) {
              if (!entry || !entry.key) return null;
              const key = String(entry.key).trim();
              const label = String(entry.label || key).trim() || key;
              const alt = Array.isArray(entry.altKeys)
                ? entry.altKeys.map(function (x) {
                    return String(x).trim();
                  }).filter(Boolean)
                : [];
              const tryKeys = [key];
              alt.forEach(function (a) {
                if (tryKeys.indexOf(a) === -1) tryKeys.push(a);
              });
              return {
                key: key,
                label: label,
                tryKeys: tryKeys,
                inputType: entry.inputType || "text",
                readOnly: !!entry.readOnly,
              };
            })
            .filter(Boolean);
        }

        function personnelPostRCQualDateKeys() {
          const col = normalizedPersonnelCleoColumns().find(function (c) {
            return c.key === "PostRCQualDate";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["PostRCQualDate"];
        }

        function personnelPostRCExpiryDateKeys() {
          const col = normalizedPersonnelCleoColumns().find(function (c) {
            return c.key === "PostRCExpirationDate";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["PostRCExpirationDate"];
        }

        function personnelCleoFlagKeys() {
          const col = normalizedPersonnelCleoColumns().find(function (c) {
            return c.key === "IsCLEO";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["IsCLEO"];
        }

        function parsePersonnelCleoFlag(raw) {
          if (raw === true || raw === 1) return true;
          if (raw === false || raw === 0 || raw === null || raw === "") return false;
          const s = String(formatCellValue(raw)).trim().toLowerCase();
          return s === "yes" || s === "true" || s === "1";
        }

        function personIsCleo(item) {
          return parsePersonnelCleoFlag(valueFromItemByKeys(item, personnelCleoFlagKeys()));
        }

        function personnelCleoDateKeys() {
          const col = normalizedPersonnelCleoColumns().find(function (c) {
            return c.key === "CLEODate";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["CLEODate"];
        }

        function formatPersonCleoDateDisplay(item) {
          const raw = valueFromItemByKeys(item, personnelCleoDateKeys());
          const d = parseWeaponsCertCalendarDate(raw);
          if (!d) return formatCellValue(raw) || "-";
          return d.toLocaleDateString("en-US", { month: "short", day: "2-digit", year: "numeric" });
        }

        function personHasPostRCRecord(item) {
          const raw = valueFromItemByKeys(item, personnelPostRCQualDateKeys());
          return !!parseWeaponsCertCalendarDate(raw);
        }

        function formatPersonPostRCQualDisplay(item) {
          const raw = valueFromItemByKeys(item, personnelPostRCQualDateKeys());
          const d = parseWeaponsCertCalendarDate(raw);
          if (!d) return formatCellValue(raw) || "-";
          return d.toLocaleDateString("en-US", { month: "short", day: "2-digit", year: "numeric" });
        }

        function formatPersonPostRCExpiryDisplay(item) {
          const expiryKeys = personnelPostRCExpiryDateKeys();
          const qualKeys = personnelPostRCQualDateKeys();
          let raw = valueFromItemByKeys(item, expiryKeys);
          if (raw === null || raw === undefined || raw === "") {
            const qual = parseWeaponsCertCalendarDate(valueFromItemByKeys(item, qualKeys));
            if (qual) {
              const exp = expirationDateFromQualDate(qual);
              return exp.toLocaleDateString("en-US", { month: "short", day: "2-digit", year: "numeric" });
            }
          }
          const d = parseWeaponsCertCalendarDate(raw);
          if (!d) return formatCellValue(raw) || "-";
          return d.toLocaleDateString("en-US", { month: "short", day: "2-digit", year: "numeric" });
        }

        function computePersonPostRCStatus(item) {
          return computeCertStatusFromExpiryKeys(
            item,
            personnelPostRCExpiryDateKeys(),
            personnelPostRCQualDateKeys(),
          );
        }

        function personnelCleoPayloadValue(el, writeKey, sampleRow) {
          if (el && el.type === "checkbox") {
            const checked = !!el.checked;
            if (sampleRow && writeKey && Object.prototype.hasOwnProperty.call(sampleRow, writeKey)) {
              const example = sampleRow[writeKey];
              if (typeof example === "boolean") return checked;
              if (typeof example === "number") return checked ? 1 : 0;
              const s = String(example).trim().toLowerCase();
              if (s === "yes" || s === "no") return checked ? "Yes" : "No";
            }
            return checked;
          }
          return formFieldPayloadValue(el, writeKey);
        }

        function applyPersonPostRCExpirationFromQual(idPrefix) {
          const qualEl = document.getElementById(idPrefix + "PostRCQualDate");
          const expEl = document.getElementById(idPrefix + "PostRCExpirationDate");
          if (!qualEl || !expEl) return;
          const qual = parseWeaponsCertCalendarDate(qualEl.value);
          if (!qual) {
            expEl.value = "";
            return;
          }
          expEl.value = isoDateFromCalendarDate(expirationDateFromQualDate(qual));
        }

        function wirePersonPostRCQualAutoExpiry(form, idPrefix) {
          if (!form || form.dataset.postRcExpiryWired === "1") return;
          form.dataset.postRcExpiryWired = "1";
          const qualFieldId = idPrefix + "PostRCQualDate";
          form.addEventListener("input", function (ev) {
            if (ev.target && ev.target.id === qualFieldId) applyPersonPostRCExpirationFromQual(idPrefix);
          });
          form.addEventListener("change", function (ev) {
            if (ev.target && ev.target.id === qualFieldId) applyPersonPostRCExpirationFromQual(idPrefix);
          });
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
          setWeaponsBulkPanelVisible(false);
          if (personWeaponsAddForm) personWeaponsAddForm.reset();
          if (personWeaponsBulkForm) personWeaponsBulkForm.reset();
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
          if (visible && personWeaponsBulkPanel) personWeaponsBulkPanel.hidden = true;
          if (!visible) {
            weaponsCertEditSession.item = null;
            setWeaponsCertFormMode("add");
          } else {
            updateWeaponsCertToolbarLabel();
          }
        }

        function setWeaponsBulkPanelVisible(visible) {
          if (personWeaponsBulkPanel) personWeaponsBulkPanel.hidden = !visible;
          if (visible) {
            setWeaponsCertAddPanelVisible(false);
            weaponsCertEditSession.item = null;
            setWeaponsCertFormMode("add");
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

        function buildWeaponsBulkSharedFields(container, idPrefix, sampleRow, columns) {
          if (!container) return;
          container.innerHTML = "";
          ["QualDate", "ExpirationDate"].forEach(function (key) {
            const col = columns.find(function (c) {
              return c.key === key;
            });
            if (!col) return;
            const f = buildWeaponsCertFieldWrap(col, sampleRow, {
              idPrefix: idPrefix,
              readOnlyKeys: key === "ExpirationDate" ? ["ExpirationDate"] : [],
            });
            if (f) container.appendChild(f);
          });
        }

        async function populateWeaponsBulkWeaponChecklist(container, seg, pw, sampleRow, columns) {
          if (!container) return [];
          container.innerHTML = "";
          const weaponCol = columns.find(function (c) {
            return c.key === "Weapon";
          });
          if (!weaponCol) {
            container.textContent = "(Weapon field not configured.)";
            return [];
          }
          const weaponWriteKey = resolveWeaponsCertWriteKey(weaponCol, sampleRow);
          let choices = [];
          try {
            const r = await fetchChoiceOptionsForColumn(weaponCol, seg, pw, sampleRow);
            choices = (r.choices || []).map(function (c) {
              return String(c).trim();
            }).filter(Boolean);
          } catch (e) {
            container.textContent = "(Could not load weapons: " + (e.message || String(e)).slice(0, 120) + ")";
            return [];
          }
          if (!choices.length) {
            container.textContent = "(No weapons found in the list.)";
            return [];
          }
          choices.forEach(function (choice) {
            const label = document.createElement("label");
            label.className = "bylaw-bulk-item-check";
            const cb = document.createElement("input");
            cb.type = "checkbox";
            cb.value = choice;
            cb.dataset.writeKey = weaponWriteKey;
            label.appendChild(cb);
            label.appendChild(document.createTextNode(choice));
            container.appendChild(label);
          });
          return choices;
        }

        function collectWeaponsBulkSharedPayload(idPrefix, columns, sampleRow) {
          const payload = {};
          columns.forEach(function (col) {
            if (col.key === "Weapon") return;
            const el = document.getElementById(idPrefix + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWeaponsCertWriteKey(col, sampleRow);
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          return payload;
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

        async function openWeaponsBulkAddPanel() {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setWeaponsCertState("warn", "Open a Personnel Record before adding qualifications.");
            return;
          }
          const pw = personDetailSession.pw;
          if (!pw) {
            setWeaponsCertState("warn", "Personnel record is still loading. Try again in a moment.");
            return;
          }
          weaponsCertEditSession.item = null;
          setWeaponsBulkPanelVisible(true);
          setWeaponsCertState("", "");
          const seg = weaponsCertListApiPath();
          const sampleRow = hubSession.weaponsCertSampleRow;
          const columns = weaponsCertFormColumns();
          buildWeaponsBulkSharedFields(personWeaponsBulkSharedFields, "wb_", sampleRow, columns);
          await populateWeaponsBulkWeaponChecklist(personWeaponsBulkItems, seg, pw, sampleRow, columns);
          if (personWeaponsBulkForm) {
            delete personWeaponsBulkForm.dataset.weaponsBulkAutoExpiryWired;
            wireCertQualDateAutoExpiry(personWeaponsBulkForm, "wb_", "weaponsBulkAutoExpiryWired");
          }
          const qualEl = document.getElementById("wb_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyCertExpirationFromQual("wb_");
          if (qualEl) qualEl.focus();
        }

        async function submitWeaponsBulkSave() {
          const personItem = personDetailSession.item;
          const pw = personDetailSession.pw;
          if (!personItem || personItem.Id == null || !pw) return;

          const selected = getSelectedBylawBulkItems(personWeaponsBulkItems);
          if (!selected.length) {
            setWeaponsCertState("err", "Select at least one weapon.");
            return;
          }
          const qualEl = document.getElementById("wb_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) {
            setWeaponsCertState("err", "Certification date is required.");
            return;
          }

          const seg = weaponsCertListApiPath();
          const sampleRow = hubSession.weaponsCertSampleRow;
          const columns = weaponsCertFormColumns();
          const sharedPayload = collectWeaponsBulkSharedPayload("wb_", columns, sampleRow);
          const personId = parseInt(String(personItem.Id), 10);
          if (!personId || isNaN(personId)) {
            setWeaponsCertState("err", "Invalid Personnel Record Id.");
            return;
          }
          const personKey = await resolveWeaponsPersonPostKey(seg, pw);
          sharedPayload[personKey] = personId;

          try {
            if (personWeaponsBulkSaveBtn) personWeaponsBulkSaveBtn.disabled = true;
            let saved = 0;
            for (let i = 0; i < selected.length; i++) {
              setWeaponsCertState("loading", "Saving " + (i + 1) + " of " + selected.length + "...");
              const payload = Object.assign({}, sharedPayload);
              payload[selected[i].writeKey] = selected[i].value;
              if (WEAPONS_CERT_SET_TITLE) payload.Title = selected[i].value;
              await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
              saved++;
            }
            setWeaponsBulkPanelVisible(false);
            if (personWeaponsBulkForm) personWeaponsBulkForm.reset();
            await loadPersonWeaponsCertifications(personItem.Id, pw);
            setWeaponsCertState("ok", "Saved " + saved + " qualification(s).");
            window.setTimeout(function () {
              setWeaponsCertState("", "");
            }, 2500);
          } catch (e) {
            setWeaponsCertState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
            log("Weapons bulk save failed:\n" + (e.message || String(e)), "err");
          } finally {
            if (personWeaponsBulkSaveBtn) personWeaponsBulkSaveBtn.disabled = false;
          }
        }

        async function openWeaponsCertEditPanel(item) {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setWeaponsCertState("warn", "Open a Personnel Record before requalifying.");
            return;
          }
          if (!item || item.Id == null) return;
          setWeaponsBulkPanelVisible(false);
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
          setBylawBulkPanelVisible(false);
          if (personBylawAddForm) personBylawAddForm.reset();
          if (personBylawBulkForm) personBylawBulkForm.reset();
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
          if (visible && personBylawBulkPanel) personBylawBulkPanel.hidden = true;
          if (!visible) {
            bylawTrainingEditSession.item = null;
            setBylawTrainingFormMode("add");
          } else {
            updateBylawTrainingToolbarLabel();
          }
        }

        function setBylawBulkPanelVisible(visible) {
          if (personBylawBulkPanel) personBylawBulkPanel.hidden = !visible;
          if (visible) {
            setBylawTrainingAddPanelVisible(false);
            bylawTrainingEditSession.item = null;
            setBylawTrainingFormMode("add");
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

        function buildBylawBulkSharedFields(container, idPrefix, sampleRow, columns) {
          if (!container) return;
          container.innerHTML = "";
          ["QualDate", "Certifier", "ExpirationDate"].forEach(function (key) {
            const col = columns.find(function (c) {
              return c.key === key;
            });
            if (!col) return;
            const f = buildBylawTrainingFieldWrap(col, sampleRow, {
              idPrefix: idPrefix,
              readOnlyKeys: key === "ExpirationDate" ? ["ExpirationDate"] : [],
            });
            if (f) container.appendChild(f);
          });
        }

        async function populateBylawBulkItemChecklist(container, seg, pw, sampleRow, columns) {
          if (!container) return [];
          container.innerHTML = "";
          const itemCol = columns.find(function (c) {
            return c.key === "Item";
          });
          if (!itemCol) {
            container.textContent = "(Training Item field not configured.)";
            return [];
          }
          const itemWriteKey = resolveBylawTrainingWriteKey(itemCol, sampleRow);
          let choices = [];
          try {
            const r = await fetchChoiceOptionsForColumn(itemCol, seg, pw, sampleRow);
            choices = (r.choices || []).map(function (c) {
              return String(c).trim();
            }).filter(Boolean);
          } catch (e) {
            container.textContent = "(Could not load training items: " + (e.message || String(e)).slice(0, 120) + ")";
            return [];
          }
          if (!choices.length) {
            container.textContent = "(No training items found in the list.)";
            return [];
          }
          choices.forEach(function (choice) {
            const label = document.createElement("label");
            label.className = "bylaw-bulk-item-check";
            const cb = document.createElement("input");
            cb.type = "checkbox";
            cb.value = choice;
            cb.dataset.writeKey = itemWriteKey;
            label.appendChild(cb);
            label.appendChild(document.createTextNode(choice));
            container.appendChild(label);
          });
          return choices;
        }

        function getSelectedBylawBulkItems(container) {
          if (!container) return [];
          return Array.prototype.slice
            .call(container.querySelectorAll('input[type="checkbox"]:checked'))
            .map(function (cb) {
              return {
                value: String(cb.value || "").trim(),
                writeKey: String(cb.dataset.writeKey || "Item").trim() || "Item",
              };
            })
            .filter(function (row) {
              return !!row.value;
            });
        }

        function collectBylawBulkSharedPayload(idPrefix, columns, sampleRow) {
          const payload = {};
          columns.forEach(function (col) {
            if (col.key === "Item") return;
            const el = document.getElementById(idPrefix + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveBylawTrainingWriteKey(col, sampleRow);
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          return payload;
        }

        function setBylawBulkItemChecks(container, checked) {
          if (!container) return;
          Array.prototype.slice.call(container.querySelectorAll('input[type="checkbox"]')).forEach(function (cb) {
            cb.checked = !!checked;
          });
        }

        async function populateBylawBulkCertifierDropdown(idPrefix, seg, pw, sampleRow) {
          const sel = document.getElementById(idPrefix + "Certifier");
          if (!sel || sel.tagName !== "SELECT") return;
          await fillBylawCertifierDropdown(sel, seg, pw, sampleRow);
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

        async function openBylawBulkAddPanel() {
          if (!personDetailSession.item || personDetailSession.item.Id == null) {
            setBylawTrainingState("warn", "Open a Personnel Record before adding training.");
            return;
          }
          const pw = personDetailSession.pw;
          if (!pw) {
            setBylawTrainingState("warn", "Personnel record is still loading. Try again in a moment.");
            return;
          }
          bylawTrainingEditSession.item = null;
          setBylawBulkPanelVisible(true);
          setBylawTrainingState("", "");
          const seg = bylawTrainingListApiPath();
          const sampleRow = hubSession.bylawTrainingSampleRow;
          const columns = bylawTrainingFormColumns();
          buildBylawBulkSharedFields(personBylawBulkSharedFields, "bb_", sampleRow, columns);
          await populateBylawBulkCertifierDropdown("bb_", seg, pw, sampleRow);
          await populateBylawBulkItemChecklist(personBylawBulkItems, seg, pw, sampleRow, columns);
          if (personBylawBulkForm) {
            delete personBylawBulkForm.dataset.bylawBulkAutoExpiryWired;
            wireCertQualDateAutoExpiry(personBylawBulkForm, "bb_", "bylawBulkAutoExpiryWired");
          }
          const qualEl = document.getElementById("bb_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyCertExpirationFromQual("bb_");
          if (personBylawBulkForm) applyAllSelectAutosizes(personBylawBulkForm);
          if (qualEl) qualEl.focus();
        }

        async function submitBylawBulkSave() {
          const personItem = personDetailSession.item;
          const pw = personDetailSession.pw;
          if (!personItem || personItem.Id == null || !pw) return;

          const selected = getSelectedBylawBulkItems(personBylawBulkItems);
          if (!selected.length) {
            setBylawTrainingState("err", "Select at least one training item.");
            return;
          }
          const qualEl = document.getElementById("bb_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) {
            setBylawTrainingState("err", "Certification date is required.");
            return;
          }

          const seg = bylawTrainingListApiPath();
          const sampleRow = hubSession.bylawTrainingSampleRow;
          const columns = bylawTrainingFormColumns();
          const sharedPayload = collectBylawBulkSharedPayload("bb_", columns, sampleRow);
          const personId = parseInt(String(personItem.Id), 10);
          if (!personId || isNaN(personId)) {
            setBylawTrainingState("err", "Invalid Personnel Record Id.");
            return;
          }
          const personKey = await resolveBylawPersonPostKey(seg, pw);
          sharedPayload[personKey] = personId;

          try {
            if (personBylawBulkSaveBtn) personBylawBulkSaveBtn.disabled = true;
            let saved = 0;
            for (let i = 0; i < selected.length; i++) {
              setBylawTrainingState("loading", "Saving " + (i + 1) + " of " + selected.length + "...");
              const payload = Object.assign({}, sharedPayload);
              payload[selected[i].writeKey] = selected[i].value;
              if (BYLAW_TRAINING_SET_TITLE) payload.Title = selected[i].value;
              await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
              saved++;
            }
            setBylawBulkPanelVisible(false);
            if (personBylawBulkForm) personBylawBulkForm.reset();
            await loadPersonBylawTraining(personItem.Id, pw);
            setBylawTrainingState("ok", "Saved " + saved + " training record(s).");
            window.setTimeout(function () {
              setBylawTrainingState("", "");
            }, 2500);
          } catch (e) {
            setBylawTrainingState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
            log("By-Law bulk save failed:\n" + (e.message || String(e)), "err");
          } finally {
            if (personBylawBulkSaveBtn) personBylawBulkSaveBtn.disabled = false;
          }
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

        function normalizedAppointmentsColumns() {
          const arr = Array.isArray(APPOINTMENTS_COLUMNS) ? APPOINTMENTS_COLUMNS : [];
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
              return { key: key, label: label, tryKeys: tryKeys };
            })
            .filter(Boolean);
        }

        function appointmentsListTitle() {
          return String(LIST_APPOINTMENTS || "").trim();
        }

        function appointmentsGuidRaw() {
          return String(LIST_APPOINTMENTS_GUID || "").trim();
        }

        function appointmentsListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(appointmentsGuidRaw());
        }

        function appointmentsListApiPath() {
          if (appointmentsListUsesGuid()) return "lists(guid'" + appointmentsGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(appointmentsListTitle()) + "')";
        }
        function appointmentsArchiveListTitle() {
          return String(LIST_APPOINTMENTS_ARCHIVE || "").trim();
        }
        function appointmentsArchiveGuidRaw() {
          return String(LIST_APPOINTMENTS_ARCHIVE_GUID || "").trim();
        }
        function appointmentsArchiveListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(appointmentsArchiveGuidRaw());
        }
        function appointmentsArchiveListApiPath() {
          if (appointmentsArchiveListUsesGuid()) return "lists(guid'" + appointmentsArchiveGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(appointmentsArchiveListTitle()) + "')";
        }
        function appointmentDateTimeKeys() {
          const dateCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "AppointmentDateTime";
          });
          return dateCol && dateCol.tryKeys ? dateCol.tryKeys.slice() : ["AppointmentDateTime"];
        }
        function appointmentCreatedAtKeys() {
          const col = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "CreatedAt";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["CreatedAt", "Created", "Created_x0020_At"];
        }
        function appointmentDateTimeFromItem(item) {
          return parseAppointmentDateTime(valueFromItemByKeys(item, appointmentDateTimeKeys()));
        }
        function appointmentIsPast(item) {
          const d = appointmentDateTimeFromItem(item);
          if (!d) return false;
          return d.getTime() < Date.now();
        }
        function appointmentIsUpcoming(item) {
          return !appointmentIsPast(item);
        }
        function appointmentArchiveTitleFromItem(item) {
          const when = formatAppointmentDateTimeDisplay(valueFromItemByKeys(item, appointmentDateTimeKeys()));
          const pid = appointmentPersonnelIdFromItem(item);
          const person = pid ? personnelById(pid) : null;
          const name = person ? formatPersonDisplayName(person) : pid ? "Personnel Id " + pid : "Appointment";
          return when ? name + " - " + when : name;
        }
        function appointmentArchiveDatePayloadValue(raw) {
          const d = parseAppointmentDateTime(raw);
          if (d) return d.toISOString();
          if (raw === null || raw === undefined || raw === "") return null;
          return raw;
        }
        async function resolveAppointmentsArchivePersonPostKey(archSeg, pw) {
          if (hubSession.appointmentsArchivePersonPostKey) return hubSession.appointmentsArchivePersonPostKey;
          const sampleRow = hubSession.appointmentsArchiveSampleRow;
          const tryKeys = appointmentsArchivePersonFieldCandidates();
          if (sampleRow) {
            const hit = tryKeys.find(function (k) {
              return Object.prototype.hasOwnProperty.call(sampleRow, k);
            });
            if (hit) {
              hubSession.appointmentsArchivePersonPostKey = hit;
              return hit;
            }
          }
          hubSession.appointmentsArchivePersonPostKey = String(
            APPOINTMENTS_ARCHIVE_PERSON_FIELD || "PersonnelId",
          ).trim();
          return hubSession.appointmentsArchivePersonPostKey;
        }
        async function appointmentArchivePayloadFromItem(item, archSeg, pw) {
          const sampleRow = hubSession.appointmentsArchiveSampleRow || item || null;
          const payload = {
            Title: appointmentArchiveTitleFromItem(item),
            ArchivedAt: spDateTimeNowIso(),
          };
          normalizedAppointmentsColumns().forEach(function (col) {
            const tryKeys = col.tryKeys || [col.key];
            const raw = valueFromItemByKeys(item, tryKeys);
            if (raw === null || raw === undefined || raw === "") return;
            const writeKey = resolveAppointmentsWriteKey({ key: col.key, tryKeys: tryKeys }, sampleRow);
            if (col.key === "AppointmentDateTime") {
              const v = appointmentArchiveDatePayloadValue(raw);
              if (v !== null && v !== "") payload[writeKey] = v;
            } else if (col.key === "CreatedAt") {
              const v = appointmentArchiveDatePayloadValue(raw);
              if (v !== null && v !== "") payload[writeKey] = v;
            } else if (col.key === "MissedAppointment") {
              payload[writeKey] = missedAppointmentPayloadValue(appointmentIsMissed(item), sampleRow);
            } else {
              payload[writeKey] = formatCellValue(raw);
            }
          });
          const personId = appointmentPersonnelIdFromItem(item);
          if (personId) {
            const personKey = await resolveAppointmentsArchivePersonPostKey(archSeg, pw);
            payload[personKey] = personId;
          }
          return payload;
        }
        async function archivePastAppointmentItems(seg, pw, rows) {
          const list = Array.isArray(rows) ? rows.slice() : [];
          if (!list.length) return { archived: 0, remaining: list, pastHidden: 0, failures: [] };
          const pastAll = list.filter(appointmentIsPast);
          if (!appointmentsArchiveListTitle() && !appointmentsArchiveListUsesGuid()) {
            return {
              archived: 0,
              remaining: list,
              pastHidden: pastAll.length,
              failures: [],
            };
          }
          const archSeg = appointmentsArchiveListApiPath();
          if (!pastAll.length) return { archived: 0, remaining: list, pastHidden: 0, failures: [] };
          const pastIds = new Set();
          let archived = 0;
          const failures = [];
          for (let i = 0; i < pastAll.length; i++) {
            const item = pastAll[i];
            if (!item || item.Id == null) continue;
            try {
              const payload = await appointmentArchivePayloadFromItem(item, archSeg, pw);
              const created = await spFetch(`/_api/web/${archSeg}/items`, { method: "POST", body: payload }, pw);
              const newId = created && (created.Id != null ? created.Id : created.ID);
              if (newId == null) {
                throw new Error("Archive list did not return a new item Id.");
              }
              await spFetch(`/_api/web/${seg}/items(${item.Id})`, { method: "DELETE" }, pw);
              pastIds.add(String(item.Id));
              archived++;
            } catch (e) {
              const errMsg = (e.message || String(e)).slice(0, 280);
              failures.push({ id: item.Id, message: errMsg });
              log("Appointment archive failed (Id " + item.Id + "):\n" + errMsg, "err");
            }
          }
          const remaining = list.filter(function (r) {
            return r && !pastIds.has(String(r.Id));
          });
          return { archived: archived, remaining: remaining, pastHidden: 0, failures: failures };
        }

        function setAppointmentsState(kind, message) {
          if (!personAppointmentsState) return;
          if (!message) {
            personAppointmentsState.hidden = true;
            personAppointmentsState.textContent = "";
            return;
          }
          personAppointmentsState.hidden = false;
          personAppointmentsState.className = "read-state " + kind;
          personAppointmentsState.textContent = message;
        }

        function clearAppointmentsTable() {
          if (personAppointmentsThead) personAppointmentsThead.innerHTML = "";
          if (personAppointmentsBody) personAppointmentsBody.innerHTML = "";
        }

        function clearPersonAppointmentsSection() {
          clearAppointmentsTable();
          setAppointmentsState("", "");
          if (personAppointmentsEmpty) {
            personAppointmentsEmpty.hidden = true;
            personAppointmentsEmpty.textContent = "No appointments on file for this person.";
          }
          if (personAppointmentsWrap) personAppointmentsWrap.hidden = true;
        }

        function appointmentMissedFieldSpec() {
          const missedCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "MissedAppointment";
          });
          return {
            key: "MissedAppointment",
            tryKeys: missedCol ? missedCol.tryKeys.slice() : ["MissedAppointment", "Missed_x0020_Appointment"],
          };
        }

        function appointmentIsMissed(item) {
          const spec = appointmentMissedFieldSpec();
          const raw = valueFromItemByKeys(item, spec.tryKeys);
          if (raw === true || raw === 1) return true;
          if (raw === false || raw === 0 || raw === null || raw === "") return false;
          const s = String(formatCellValue(raw)).trim().toLowerCase();
          return s === "yes" || s === "true" || s === "1";
        }

        function missedAppointmentPayloadValue(missed, sampleRow) {
          const field = appointmentMissedFieldSpec();
          const writeKey = resolveAppointmentsWriteKey(field, sampleRow);
          if (sampleRow && writeKey && Object.prototype.hasOwnProperty.call(sampleRow, writeKey)) {
            const example = sampleRow[writeKey];
            if (typeof example === "boolean") return !!missed;
            if (typeof example === "number") return missed ? 1 : 0;
            const s = String(example).trim().toLowerCase();
            if (s === "yes" || s === "no") return missed ? "Yes" : "No";
          }
          return !!missed;
        }

        async function updateAppointmentMissedStatus(item, missed, pw, onDone) {
          if (!item || item.Id == null || !pw) return;
          const seg = appointmentsListApiPath();
          const sampleRow = hubSession.appointmentsSampleRow || item;
          const field = appointmentMissedFieldSpec();
          const writeKey = resolveAppointmentsWriteKey(field, sampleRow);
          const payload = {};
          payload[writeKey] = missedAppointmentPayloadValue(!!missed, sampleRow);
          try {
            await spFetch(`/_api/web/${seg}/items(${item.Id})`, { method: "MERGE", body: payload }, pw);
            item[writeKey] = payload[writeKey];
            if (typeof onDone === "function") await onDone();
          } catch (e) {
            const msg = "Could not update missed status: " + (e.message || String(e)).slice(0, 220);
            if (schedulingSection && !schedulingSection.hidden) setSchedulingState("err", msg);
            else setAppointmentsState("err", msg);
            log("Appointment missed update failed:\n" + (e.message || String(e)), "err");
          }
        }

        function buildAppointmentActionsCell(item, refreshFn, options) {
          const opts = options || {};
          const td = document.createElement("td");
          td.className = "roster-actions scheduling-appt-actions";
          if (!opts.memoOnly) {
            const missed = appointmentIsMissed(item);
            const markBtn = document.createElement("button");
            markBtn.type = "button";
            markBtn.className = "btn-secondary";
            markBtn.textContent = missed ? "Clear missed" : "Mark missed";
            markBtn.addEventListener("click", function () {
              const pw = hubSession.pw || personDetailSession.pw;
              if (!pw) return;
              void updateAppointmentMissedStatus(item, !missed, pw, refreshFn);
            });
            td.appendChild(markBtn);
          }
          const memoBtn = document.createElement("button");
          memoBtn.type = "button";
          memoBtn.className = "btn-secondary";
          memoBtn.textContent = "Memo";
          memoBtn.addEventListener("click", function () {
            openSchedulingMemoForAppointment(item);
          });
          td.appendChild(memoBtn);
          return td;
        }

        function renderAppointmentsTable(rows) {
          clearAppointmentsTable();
          const columns = normalizedAppointmentsColumns();
          if (!personAppointmentsThead || !personAppointmentsBody || !columns.length) return;
          const showActions = !!(personDetailSession.pw && personAppointmentsWrap && !personAppointmentsWrap.hidden);

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
            thAct.title = "Mark missed / memorandum";
            trHead.appendChild(thAct);
          }
          personAppointmentsThead.appendChild(trHead);

          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            if (appointmentIsMissed(item)) tr.className = "appointments-row--missed";
            columns.forEach(function (col) {
              const td = document.createElement("td");
              const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
              let text = "";
              if (raw !== undefined && raw !== null) {
                if (col.key === "AppointmentDateTime") {
                  text = formatAppointmentDateTimeDisplay(raw);
                } else if (col.key === "MissedAppointment") {
                  text = appointmentIsMissed(item) ? "Yes" : "No";
                } else {
                  text = formatCellValue(raw);
                }
              } else if (col.key === "MissedAppointment") {
                text = "No";
              }
              td.textContent = displayCellText(text);
              tr.appendChild(td);
            });
            if (showActions) {
              tr.appendChild(
                buildAppointmentActionsCell(item, async function () {
                  if (personDetailSession.item && personDetailSession.item.Id != null) {
                    await loadPersonAppointments(personDetailSession.item.Id, personDetailSession.pw);
                  }
                }),
              );
            }
            frag.appendChild(tr);
          });
          personAppointmentsBody.appendChild(frag);

          if (personAppointmentsEmpty) {
            personAppointmentsEmpty.hidden = rows.length > 0;
          }
        }

        function appointmentsPersonFilterFieldCandidates() {
          const primary = String(APPOINTMENTS_PERSON_FIELD || "PersonnelId").trim();
          const out = [];
          if (primary) out.push(primary);
          const alts = Array.isArray(APPOINTMENTS_PERSON_FIELD_ALT) ? APPOINTMENTS_PERSON_FIELD_ALT : [];
          alts.forEach(function (name) {
            const n = String(name || "").trim();
            if (n && out.indexOf(n) === -1) out.push(n);
          });
          return out;
        }

        function appointmentsArchivePersonFieldCandidates() {
          const primary = String(APPOINTMENTS_ARCHIVE_PERSON_FIELD || "PersonnelId").trim();
          const out = [];
          if (primary) out.push(primary);
          const alts = Array.isArray(APPOINTMENTS_ARCHIVE_PERSON_FIELD_ALT)
            ? APPOINTMENTS_ARCHIVE_PERSON_FIELD_ALT
            : [];
          alts.forEach(function (name) {
            const n = String(name || "").trim();
            if (n && out.indexOf(n) === -1) out.push(n);
          });
          return out;
        }

        function appointmentPersonnelIdFromItem(item) {
          if (!item || typeof item !== "object") return null;
          const keys = appointmentsPersonFilterFieldCandidates().concat([
            "PeronnelId",
            "PersonnelId",
            "PersonnelID",
            "PersonnelIdId",
            "PeronnelIdId",
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

        function appointmentRowsMatchPersonnel(rows, personnelId) {
          const want = String(personnelId);
          return (Array.isArray(rows) ? rows : []).filter(function (item) {
            return appointmentPersonnelIdFromItem(item) === want;
          });
        }

        async function discoverAppointmentsPersonFilterFields(seg, pw) {
          if (hubSession.appointmentsPersonFilterFields && hubSession.appointmentsPersonFilterFields.length) {
            return hubSession.appointmentsPersonFilterFields.slice();
          }
          const out = [];
          try {
            const data = await spFetch(
              `/_api/web/${seg}/fields?$select=InternalName,Title,EntityPropertyName,StaticName,TypeAsString&$filter=Hidden eq false&$top=200`,
              {},
              pw,
            );
            const fields = (data && data.value) || [];
            const want = String(APPOINTMENTS_PERSON_FIELD || "PeronnelId").trim().toLowerCase();
            let hit = fields.find(function (f) {
              const title = String(f.Title || "").trim().toLowerCase();
              const internal = String(f.InternalName || f.StaticName || f.EntityPropertyName || "").trim();
              const internalLower = internal.toLowerCase();
              return (
                internalLower === want ||
                title === want ||
                title === "personnel id" ||
                title === "peronnel id" ||
                internalLower === "personnelid" ||
                internalLower === "peronnelid" ||
                internalLower === "personnelidid" ||
                internalLower === "peronnelidid"
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
          hubSession.appointmentsPersonFilterFields = out.slice();
          return out;
        }

        async function fetchAppointmentsForPersonnel(seg, pw, personnelId) {
          const id = parseInt(String(personnelId), 10);
          if (!id || isNaN(id)) return [];
          const orderByClause = String(APPOINTMENTS_ITEMS_ORDERBY || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          const tried = [];
          const candidates = appointmentsPersonFilterFieldCandidates().slice();

          const discovered = await discoverAppointmentsPersonFilterFields(seg, pw);
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
                data = await spFetch(`/_api/web/${seg}/items?$top=500&$filter=${filter}` + orderByQs, {}, pw);
              } catch (e0) {
                if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
                  data = await spFetch(`/_api/web/${seg}/items?$top=500&$filter=${filter}`, {}, pw);
                } else {
                  throw e0;
                }
              }
              hubSession.appointmentsPersonFilterField = personField;
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
          return appointmentRowsMatchPersonnel((data && data.value) || [], id);
        }

        async function loadPersonAppointments(personId, pw) {
          if (!personId) return;
          if (personAppointmentsWrap) personAppointmentsWrap.hidden = false;
          setAppointmentsState("loading", "Loading appointments...");
          clearAppointmentsTable();
          if (personAppointmentsEmpty) personAppointmentsEmpty.hidden = true;

          const seg = appointmentsListApiPath();
          if (!appointmentsListTitle() && !appointmentsListUsesGuid()) {
            renderAppointmentsTable([]);
            setAppointmentsState("", "");
            if (personAppointmentsEmpty) {
              personAppointmentsEmpty.hidden = false;
              personAppointmentsEmpty.textContent = "Appointments list is not configured.";
            }
            return;
          }

          try {
            let rows = await fetchAppointmentsForPersonnel(seg, pw, personId);
            const archiveResult = await archivePastAppointmentItems(seg, pw, rows);
            rows = archiveResult.remaining.filter(appointmentIsUpcoming);
            const dateKeys = [
              "AppointmentDateTime",
              "Appointment_x0020_Date_x0020_Time",
              "Appointment_x0020_DateTime",
              "AppointmentTimeDate",
              "AppointmentTime_x0020_Date",
              "AppointmentDate",
              "Appointment_x0020_Date",
              "Date",
            ];
            rows.sort(function (a, b) {
              const da = formatCellValue(valueFromItemByKeys(a, dateKeys) || "");
              const db = formatCellValue(valueFromItemByKeys(b, dateKeys) || "");
              if (da !== db) return String(da).localeCompare(String(db));
              const ca = formatCellValue(valueFromItemByKeys(a, ["CreatedAt", "Created"]) || "");
              const cb = formatCellValue(valueFromItemByKeys(b, ["CreatedAt", "Created"]) || "");
              return String(ca).localeCompare(String(cb));
            });
            hubSession.appointmentsSampleRow = rows.length ? rows[0] : hubSession.appointmentsSampleRow;
            renderAppointmentsTable(rows);
            setAppointmentsState("", "");
          } catch (e) {
            renderAppointmentsTable([]);
            setAppointmentsState("warn", "Could not load appointments: " + (e.message || String(e)).slice(0, 220));
            if (personAppointmentsEmpty) {
              personAppointmentsEmpty.hidden = false;
              personAppointmentsEmpty.textContent =
                "No appointments loaded. If PersonnelId exists on the list, check that values match the personnel record Id.";
            }
          }
        }

        function normalizedAppointmentsFormFields() {
          const arr = Array.isArray(APPOINTMENTS_FORM_FIELDS) ? APPOINTMENTS_FORM_FIELDS : [];
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
              return {
                key: key,
                label: label,
                tryKeys: tryKeys,
                inputType: entry.inputType || "text",
                required: !!entry.required,
              };
            })
            .filter(Boolean);
        }

        function resolveAppointmentsWriteKey(field, sampleRow) {
          const tryKeys = field.tryKeys || [field.key];
          if (sampleRow) {
            const hit = tryKeys.find(function (k) {
              return Object.prototype.hasOwnProperty.call(sampleRow, k);
            });
            if (hit) return hit;
          }
          return field.key;
        }

        async function resolveAppointmentsPersonPostKey(seg, pw) {
          if (hubSession.appointmentsPersonPostKey) return hubSession.appointmentsPersonPostKey;
          if (hubSession.appointmentsPersonFilterField) {
            hubSession.appointmentsPersonPostKey = hubSession.appointmentsPersonFilterField;
            return hubSession.appointmentsPersonPostKey;
          }
          const discovered = await discoverAppointmentsPersonFilterFields(seg, pw);
          const key = discovered[0] || String(APPOINTMENTS_PERSON_FIELD || "PersonnelId").trim();
          hubSession.appointmentsPersonPostKey = key;
          return key;
        }

        function parseAppointmentDateTime(val) {
          if (val === null || val === undefined || val === "") return null;
          const m = String(val).match(/^\/Date\((\d+)\)\/$/);
          if (m) {
            const d = new Date(parseInt(m[1], 10));
            return isNaN(d.getTime()) ? null : d;
          }
          const d = new Date(val);
          return isNaN(d.getTime()) ? null : d;
        }

        function formatAppointmentDateTimeDisplay(val) {
          const d = parseAppointmentDateTime(val);
          if (!d) return formatCellValue(val);
          return d.toLocaleString("en-US", {
            month: "short",
            day: "2-digit",
            year: "numeric",
            hour: "numeric",
            minute: "2-digit",
          });
        }

        function archivedAtFromItem(item) {
          return parseAppointmentDateTime(valueFromItemByKeys(item, ["ArchivedAt", "Archived_x0020_At"]));
        }

        function appointmentArchiveRetentionInfo(item) {
          const retentionDays = Number(APPOINTMENTS_ARCHIVE_RETENTION_DAYS) || 30;
          const warnYellow = Number(APPOINTMENTS_ARCHIVE_WARN_YELLOW_DAYS) || 10;
          const warnRed = Number(APPOINTMENTS_ARCHIVE_WARN_RED_DAYS) || 5;
          const archivedAt = archivedAtFromItem(item);
          if (!archivedAt) return { daysUntilDelete: null, tier: null, expired: false };
          const msPerDay = 86400000;
          const deleteAt = archivedAt.getTime() + retentionDays * msPerDay;
          const daysUntilDelete = Math.ceil((deleteAt - Date.now()) / msPerDay);
          let tier = null;
          if (daysUntilDelete <= warnRed) tier = "red";
          else if (daysUntilDelete <= warnYellow) tier = "yellow";
          return {
            daysUntilDelete: daysUntilDelete,
            tier: tier,
            expired: daysUntilDelete <= 0,
          };
        }

        function appointmentArchiveStatusLabel(item, pendingArchive) {
          if (pendingArchive) return "Pending archive";
          const when = formatArchivedAtDisplay(item);
          const info = appointmentArchiveRetentionInfo(item);
          if (info.daysUntilDelete == null) return when;
          if (info.expired) return when + " (expired)";
          if (info.tier) return when + " (deletes in " + info.daysUntilDelete + "d)";
          return when;
        }

        function formatArchivedAtDisplay(item) {
          const d = archivedAtFromItem(item);
          if (!d) {
            return formatCellValue(valueFromItemByKeys(item, ["ArchivedAt", "Archived_x0020_At"]));
          }
          return d.toLocaleString("en-US", {
            month: "short",
            day: "2-digit",
            year: "numeric",
            hour: "numeric",
            minute: "2-digit",
          });
        }

        async function purgeExpiredArchivedAppointments(archSeg, pw, rows) {
          const list = Array.isArray(rows) ? rows.slice() : [];
          if (!list.length) return { remaining: list, purged: 0 };
          const remaining = [];
          let purged = 0;
          for (let i = 0; i < list.length; i++) {
            const item = list[i];
            if (!item || item.Id == null) {
              if (item) remaining.push(item);
              continue;
            }
            const info = appointmentArchiveRetentionInfo(item);
            if (!info.expired) {
              remaining.push(item);
              continue;
            }
            try {
              await spFetch(`/_api/web/${archSeg}/items(${item.Id})`, { method: "DELETE" }, pw);
              purged++;
            } catch (e) {
              log(
                "Expired archive delete failed (Id " + item.Id + "):\n" + (e.message || String(e)),
                "err",
              );
              remaining.push(item);
            }
          }
          return { remaining: remaining, purged: purged };
        }

        function isoDateTimeLocalFromValue(val) {
          const d = parseAppointmentDateTime(val);
          if (!d) return "";
          const y = d.getFullYear();
          const mo = String(d.getMonth() + 1).padStart(2, "0");
          const day = String(d.getDate()).padStart(2, "0");
          const h = String(d.getHours()).padStart(2, "0");
          const mi = String(d.getMinutes()).padStart(2, "0");
          return y + "-" + mo + "-" + day + "T" + h + ":" + mi;
        }

        function personnelById(personnelId) {
          const id = String(personnelId || "").trim();
          if (!id) return null;
          const rows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
          return (
            rows.find(function (r) {
              return r && String(r.Id) === id;
            }) || null
          );
        }

        function officeSymbolForPersonnelId(personnelId) {
          const person = personnelById(personnelId);
          const office = person ? itemFieldText(person, "OfficeSymbol") : "";
          return String(office || "").trim() || "Unassigned";
        }

        function schedulingRowViewModel(item) {
          const pid = appointmentPersonnelIdFromItem(item);
          const person = pid ? personnelById(pid) : null;
          const dateCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "AppointmentDateTime";
          });
          const dateKeys = dateCol ? dateCol.tryKeys : ["AppointmentDateTime"];
          const providerCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "ProviderOffice";
          });
          const reasonCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "ReasonType";
          });
          const instructorCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "InstructorInitials";
          });
          const missedCol = normalizedAppointmentsColumns().find(function (c) {
            return c.key === "MissedAppointment";
          });
          return {
            item: item,
            personnelId: pid,
            office: officeSymbolForPersonnelId(pid),
            name: person ? formatPersonDisplayName(person) : pid ? "Personnel Id " + pid : "Unknown",
            when: formatAppointmentDateTimeDisplay(valueFromItemByKeys(item, dateKeys)),
            whenSort: parseAppointmentDateTime(valueFromItemByKeys(item, dateKeys)),
            location: formatCellValue(
              valueFromItemByKeys(item, providerCol ? providerCol.tryKeys : ["ProviderOffice"]) || "",
            ),
            description: formatCellValue(
              valueFromItemByKeys(item, reasonCol ? reasonCol.tryKeys : ["ReasonType"]) || "",
            ),
            instructor: formatCellValue(
              valueFromItemByKeys(item, instructorCol ? instructorCol.tryKeys : ["InstructorInitials"]) || "",
            ),
            missedFlag: appointmentIsMissed(item),
            missed: appointmentIsMissed(item) ? "Yes" : "No",
            rank: person ? itemFieldText(person, "Rank") : "",
            archivedAt: formatArchivedAtDisplay(item),
            archiveRetentionTier: appointmentArchiveRetentionInfo(item).tier,
            createdAt: formatAppointmentDateTimeDisplay(valueFromItemByKeys(item, appointmentCreatedAtKeys())),
          };
        }

        function schedulingPendingArchiveRows() {
          return (schedulingSession.rows || []).filter(appointmentIsPast);
        }

        function schedulingArchiveTabRows() {
          const out = [];
          schedulingPendingArchiveRows().forEach(function (item) {
            out.push({ item: item, pendingArchive: true });
          });
          (schedulingSession.archiveRows || []).forEach(function (item) {
            out.push({ item: item, pendingArchive: false });
          });
          return out;
        }

        function normalizeSchedulingTableEntry(entry) {
          if (entry && entry.item) {
            return { item: entry.item, pendingArchive: !!entry.pendingArchive };
          }
          return { item: entry, pendingArchive: false };
        }

        function schedulingSourceRowsForView() {
          if (schedulingSession.listView === "archive") {
            return schedulingArchiveTabRows();
          }
          return schedulingSession.rows || [];
        }

        function filterAppointmentRowsForListView(rows) {
          const list = Array.isArray(rows) ? rows.slice() : [];
          if (schedulingSession.listView === "archive") {
            return list;
          }
          if (schedulingSession.listView === "missed") {
            return list.filter(function (item) {
              return appointmentIsUpcoming(item) && appointmentIsMissed(item);
            });
          }
          return list.filter(appointmentIsUpcoming);
        }

        function updateSchedulingListTitle(filteredCount, totalCount) {
          if (!schedulingListTitle) return;
          if (schedulingSession.listView === "missed") {
            schedulingListTitle.textContent =
              "Missed appointments (" + filteredCount + " of " + totalCount + " total)";
          } else if (schedulingSession.listView === "archive") {
            const pendingCount = schedulingPendingArchiveRows().length;
            schedulingListTitle.textContent =
              "Archived appointments (" + filteredCount + ")" +
              (pendingCount ? " - " + pendingCount + " pending archive" : "");
          } else {
            schedulingListTitle.textContent = "Upcoming appointments (" + filteredCount + ")";
          }
        }

        function setSchedulingListView(view) {
          schedulingSession.listView =
            view === "missed" ? "missed" : view === "archive" ? "archive" : "all";
          if (schedulingTabAll) {
            schedulingTabAll.classList.toggle("scheduling-tab--active", schedulingSession.listView === "all");
          }
          if (schedulingTabMissed) {
            schedulingTabMissed.classList.toggle("scheduling-tab--active", schedulingSession.listView === "missed");
          }
          if (schedulingTabArchive) {
            schedulingTabArchive.classList.toggle("scheduling-tab--active", schedulingSession.listView === "archive");
          }
          refreshSchedulingListDisplay();
        }

        function refreshSchedulingListDisplay() {
          const rows = schedulingSourceRowsForView();
          const filtered = filterAppointmentRowsForListView(rows);
          const totalCount =
            schedulingSession.listView === "archive"
              ? filtered.length
              : (schedulingSession.rows || []).length;
          updateSchedulingListTitle(filtered.length, totalCount);
          renderSchedulingTable(filtered, totalCount);
          if (schedulingEmpty) {
            schedulingEmpty.hidden = filtered.length > 0;
            schedulingEmpty.textContent =
              schedulingSession.listView === "missed"
                ? "No missed appointments on file."
                : schedulingSession.listView === "archive"
                  ? "No archived appointments on file."
                  : "No upcoming appointments on file.";
          }
        }

        // ----- Phase-1 Tracker -----
        function phaseOneTrackingListTitle() {
          return String(LIST_PHASE_ONE_TRACKING || "").trim();
        }
        function phaseOneTrackingGuidRaw() {
          return String(LIST_PHASE_ONE_TRACKING_GUID || "").trim();
        }
        function phaseOneTrackingListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(phaseOneTrackingGuidRaw());
        }
        function phaseOneTrackingListApiPath() {
          if (phaseOneTrackingListUsesGuid()) return "lists(guid'" + phaseOneTrackingGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(phaseOneTrackingListTitle()) + "')";
        }
        function phaseOneArchiveListTitle() {
          return String(LIST_PHASE_ONE_ARCHIVE || "").trim();
        }
        function phaseOneArchiveGuidRaw() {
          return String(LIST_PHASE_ONE_ARCHIVE_GUID || "").trim();
        }
        function phaseOneArchiveListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(phaseOneArchiveGuidRaw());
        }
        function phaseOneArchiveListApiPath() {
          if (phaseOneArchiveListUsesGuid()) return "lists(guid'" + phaseOneArchiveGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(phaseOneArchiveListTitle()) + "')";
        }
        function normalizedPhaseOneColumns(view) {
          const arr =
            view === "archive"
              ? Array.isArray(PHASE_ONE_ARCHIVE_COLUMNS)
                ? PHASE_ONE_ARCHIVE_COLUMNS
                : []
              : Array.isArray(PHASE_ONE_TRACKING_COLUMNS)
                ? PHASE_ONE_TRACKING_COLUMNS
                : [];
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
              return { key: key, label: label, tryKeys: tryKeys };
            })
            .filter(Boolean);
        }
        function setPhase1State(kind, message) {
          if (!phase1ReadState) return;
          if (!message) {
            phase1ReadState.hidden = true;
            phase1ReadState.textContent = "";
            return;
          }
          phase1ReadState.hidden = false;
          phase1ReadState.className = "read-state " + kind;
          phase1ReadState.textContent = message;
        }
        function setPhase1AddPanelVisible(visible) {
          if (phase1AddPanel) phase1AddPanel.hidden = !visible;
          if (visible && phase1EditPanel) phase1EditPanel.hidden = true;
        }
        function setPhase1EditPanelVisible(visible) {
          if (phase1EditPanel) phase1EditPanel.hidden = !visible;
          if (visible && phase1AddPanel) phase1AddPanel.hidden = true;
        }
        function addCalendarDays(date, days) {
          const out = new Date(date.getFullYear(), date.getMonth(), date.getDate());
          out.setDate(out.getDate() + days);
          return out;
        }
        function calendarDateKey(d) {
          return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
        }
        function nthWeekdayOfMonth(year, monthIndex, weekday, n) {
          let count = 0;
          for (let day = 1; day <= 31; day++) {
            const cur = new Date(year, monthIndex, day);
            if (cur.getMonth() !== monthIndex) break;
            if (cur.getDay() === weekday) {
              count++;
              if (count === n) return cur;
            }
          }
          return null;
        }
        function lastWeekdayOfMonth(year, monthIndex, weekday) {
          let last = null;
          for (let day = 1; day <= 31; day++) {
            const cur = new Date(year, monthIndex, day);
            if (cur.getMonth() !== monthIndex) break;
            if (cur.getDay() === weekday) last = cur;
          }
          return last;
        }
        function observedFederalHoliday(year, month, day) {
          const actual = new Date(year, month - 1, day);
          const dow = actual.getDay();
          if (dow === 6) return addCalendarDays(actual, -1);
          if (dow === 0) return addCalendarDays(actual, 1);
          return actual;
        }
        const phaseOneFederalHolidayCache = {};
        function federalHolidayKeysForYear(year) {
          if (phaseOneFederalHolidayCache[year]) return phaseOneFederalHolidayCache[year];
          const keys = new Set();
          function addKey(d) {
            if (d) keys.add(calendarDateKey(d));
          }
          addKey(observedFederalHoliday(year, 1, 1));
          addKey(nthWeekdayOfMonth(year, 0, 1, 3));
          addKey(nthWeekdayOfMonth(year, 1, 1, 3));
          addKey(lastWeekdayOfMonth(year, 4, 1));
          addKey(observedFederalHoliday(year, 6, 19));
          addKey(observedFederalHoliday(year, 7, 4));
          addKey(nthWeekdayOfMonth(year, 8, 1, 1));
          addKey(nthWeekdayOfMonth(year, 9, 1, 2));
          addKey(observedFederalHoliday(year, 11, 11));
          addKey(nthWeekdayOfMonth(year, 10, 4, 4));
          addKey(observedFederalHoliday(year, 12, 25));
          phaseOneFederalHolidayCache[year] = keys;
          return keys;
        }
        function isUSFederalHoliday(d) {
          const keys = federalHolidayKeysForYear(d.getFullYear());
          if (keys.has(calendarDateKey(d))) return true;
          const prev = federalHolidayKeysForYear(d.getFullYear() - 1);
          const next = federalHolidayKeysForYear(d.getFullYear() + 1);
          const key = calendarDateKey(d);
          return prev.has(key) || next.has(key);
        }
        function isWeekendDate(d) {
          const day = d.getDay();
          return day === 0 || day === 6;
        }
        function hasWeekdayImmediatelyAfter(d) {
          const next = addCalendarDays(d, 1);
          if (isWeekendDate(next)) return false;
          if (isUSFederalHoliday(next)) return false;
          return true;
        }
        function mondayAfterDate(d) {
          let cur = addCalendarDays(d, 1);
          while (cur.getDay() !== 1) {
            cur = addCalendarDays(cur, 1);
          }
          return cur;
        }
        function phaseOneStartDateAfterHouseHunting(dateArrived) {
          const days = parseInt(String(PHASE_ONE_HOUSE_HUNTING_DAYS || "10"), 10) || 10;
          let candidate = addCalendarDays(dateArrived, days);
          for (let guard = 0; guard < 21; guard++) {
            if (isWeekendDate(candidate)) {
              candidate = candidate.getDay() === 6 ? addCalendarDays(candidate, 2) : addCalendarDays(candidate, 1);
              continue;
            }
            if (isUSFederalHoliday(candidate) && !hasWeekdayImmediatelyAfter(candidate)) {
              candidate = mondayAfterDate(candidate);
              continue;
            }
            break;
          }
          return candidate;
        }
        function spDateTimeFromCalendarDate(d) {
          return isoDateFromCalendarDate(d) + "T00:00:00Z";
        }
        function spDateTimeNowIso() {
          return new Date().toISOString();
        }
        function personnelRecordDate(person) {
          const raw = valueFromItemByKeys(person, [
            "RecordDate",
            "Record_x0020_Date",
            "Record0",
            "DateArrived",
            "Date_x0020_Arrived",
          ]);
          const d = parseWeaponsCertCalendarDate(raw);
          return d || new Date();
        }
        function phaseOnePersonnelIdFromItem(item) {
          if (!item || typeof item !== "object") return null;
          const keys = [
            "PersonnelId",
            "PersonnelID",
            "PersonnelIdId",
            "Personnel_x0020_Id",
            "Personnel_x0020_ID",
          ];
          for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (Object.prototype.hasOwnProperty.call(item, k) && item[k] !== null && item[k] !== "") {
              return String(item[k]);
            }
          }
          if (item.Personnel && typeof item.Personnel === "object" && item.Personnel.Id != null) {
            return String(item.Personnel.Id);
          }
          return null;
        }
        function phaseOneIsActive(item) {
          const raw = valueFromItemByKeys(item, ["IsActive", "Active"]);
          if (raw === true || raw === 1 || String(raw).toLowerCase() === "yes" || String(raw) === "1") return true;
          if (raw === false || raw === 0 || String(raw).toLowerCase() === "no" || String(raw) === "0") return false;
          return true;
        }
        function phaseOneDisplayDate(val) {
          if (val === null || val === undefined || val === "") return "";
          return formatWeaponsCertDisplayDate(val) || formatCellValue(val);
        }
        function phaseOneFieldText(item, col) {
          const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
          if (col.key === "DateArrived" || col.key === "PhaseOneStartDate" || col.key === "DateReleased" || col.key === "ArchivedAt") {
            return phaseOneDisplayDate(raw);
          }
          return displayCellText(raw !== undefined && raw !== null ? formatCellValue(raw) : "");
        }
        async function resolvePhaseOnePersonPostKey(seg, pw) {
          if (phase1Session.personPostKey) return phase1Session.personPostKey;
          const filterField = String(PHASE_ONE_TRACKING_PERSON_FIELD || "PersonnelId").trim();
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
          phase1Session.personPostKey = postKey;
          return postKey;
        }
        async function fetchPhaseOneListRows(seg, pw, orderByConst) {
          const orderByClause = String(orderByConst || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=5000` + orderByQs, {}, pw);
          } catch (e0) {
            if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
              data = await spFetch(`/_api/web/${seg}/items?$top=5000`, {}, pw);
            } else {
              throw e0;
            }
          }
          return (data && data.value) || [];
        }
        function clearPhase1Table() {
          if (phase1Thead) phase1Thead.innerHTML = "";
          if (phase1TableBody) phase1TableBody.innerHTML = "";
        }
        function setPhase1ListView(view) {
          phase1Session.listView = view === "archive" ? "archive" : "active";
          if (phase1TabActive) phase1TabActive.classList.toggle("scheduling-tab--active", phase1Session.listView === "active");
          if (phase1TabArchive) phase1TabArchive.classList.toggle("scheduling-tab--active", phase1Session.listView === "archive");
          if (phase1ListTitle) {
            phase1ListTitle.textContent = phase1Session.listView === "archive" ? "Phase 1 archive" : "Active Phase 1";
          }
          refreshPhase1ListDisplay();
        }
        function refreshPhase1ListDisplay() {
          const isArchive = phase1Session.listView === "archive";
          const rows = isArchive
            ? phase1Session.archiveRows || []
            : (phase1Session.trackingRows || []).filter(function (item) {
                return phaseOneIsActive(item);
              });
          renderPhase1Table(rows, isArchive);
        }
        function renderPhase1Table(rows, isArchive) {
          clearPhase1Table();
          const columns = normalizedPhaseOneColumns(isArchive ? "archive" : "active");
          if (!phase1Thead || !phase1TableBody || !columns.length) return;
          const showActions = !isArchive && !!hubSession.pw;
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
            trHead.appendChild(thAct);
          }
          phase1Thead.appendChild(trHead);
          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              td.textContent = phaseOneFieldText(item, col);
              tr.appendChild(td);
            });
            if (showActions && item.Id != null) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              const editBtn = document.createElement("button");
              editBtn.type = "button";
              editBtn.className = "btn-record";
              editBtn.textContent = "Edit";
              editBtn.addEventListener("click", function () {
                void openPhase1EditPanel(item);
              });
              tdAct.appendChild(editBtn);
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          phase1TableBody.appendChild(frag);
          if (phase1Empty) {
            phase1Empty.hidden = rows.length > 0;
            phase1Empty.textContent = isArchive
              ? "No archived Phase 1 records on file."
              : "No active Phase 1 records on file.";
          }
        }
        function personnelAlreadyInActivePhase1(personId) {
          const rows = phase1Session.trackingRows || [];
          const want = String(personId);
          return rows.some(function (item) {
            return phaseOneIsActive(item) && String(phaseOnePersonnelIdFromItem(item) || "") === want;
          });
        }
        function populatePhase1PersonSelect() {
          if (!phase1PersonSelect) return;
          while (phase1PersonSelect.options.length > 1) phase1PersonSelect.remove(1);
          const rows = Array.isArray(hubSession.rows) ? hubSession.rows.slice() : [];
          rows.sort(function (a, b) {
            return formatPersonDisplayName(a).localeCompare(formatPersonDisplayName(b), undefined, { sensitivity: "base" });
          });
          rows.forEach(function (person) {
            if (!person || person.Id == null) return;
            if (personnelAlreadyInActivePhase1(person.Id)) return;
            const o = document.createElement("option");
            o.value = String(person.Id);
            o.textContent = formatPersonDisplayName(person);
            phase1PersonSelect.appendChild(o);
          });
          applySelectAutosize(phase1PersonSelect);
        }
        function openPhase1AddPanel() {
          if (!Array.isArray(hubSession.rows) || !hubSession.rows.length) {
            setPhase1State("warn", "Personnel roster is empty. Add personnel first.");
            return;
          }
          populatePhase1PersonSelect();
          setPhase1AddPanelVisible(true);
          setPhase1State("", "");
        }
        async function enrollPersonInPhase1(houseHunting) {
          const pw = hubSession.pw;
          if (!pw || !phase1PersonSelect) return;
          const personId = parseInt(String(phase1PersonSelect.value || ""), 10);
          if (!personId || isNaN(personId)) {
            setPhase1State("err", "Select personnel to enter Phase 1.");
            return;
          }
          const person = (hubSession.rows || []).find(function (r) {
            return r && String(r.Id) === String(personId);
          });
          if (!person) {
            setPhase1State("err", "Selected personnel not found on roster.");
            return;
          }
          if (personnelAlreadyInActivePhase1(personId)) {
            setPhase1State("warn", formatPersonDisplayName(person) + " is already in active Phase 1.");
            return;
          }
          const seg = phaseOneTrackingListApiPath();
          const dateArrived = personnelRecordDate(person);
          const phaseOneStart = houseHunting ? phaseOneStartDateAfterHouseHunting(dateArrived) : dateArrived;
          const dutyStatus = houseHunting ? PHASE_ONE_DUTY_STATUS_HOUSE_HUNTING : PHASE_ONE_DUTY_STATUS_ACTIVE;
          const dutySection = itemFieldText(person, "OfficeSymbol") || itemFieldText(person, "Squadron") || "";
          const personKey = await resolvePhaseOnePersonPostKey(seg, pw);
          const payload = {
            PersonnelName: formatPersonDisplayName(person),
            DutyStatus: dutyStatus,
            DutySection: dutySection,
            DateArrived: spDateTimeFromCalendarDate(dateArrived),
            PhaseOneStartDate: spDateTimeFromCalendarDate(phaseOneStart),
            IsActive: true,
            Notes: "",
          };
          payload[personKey] = personId;
          payload.Title = formatPersonDisplayName(person);
          try {
            setPhase1State("loading", "Entering Phase 1...");
            await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            setPhase1AddPanelVisible(false);
            if (phase1PersonSelect) phase1PersonSelect.value = "";
            await loadPhase1TrackerData();
            setPhase1State(
              "ok",
              formatPersonDisplayName(person) +
                " entered Phase 1" +
                (houseHunting ? " (house hunting)." : "."),
            );
            window.setTimeout(function () {
              setPhase1State("", "");
            }, 2500);
          } catch (e) {
            setPhase1State("err", "Could not enter Phase 1: " + (e.message || String(e)).slice(0, 280));
          }
        }
        function buildPhase1EditFormFields(item) {
          if (!phase1EditFields) return;
          phase1EditFields.innerHTML = "";
          (Array.isArray(PHASE_ONE_EDIT_FIELDS) ? PHASE_ONE_EDIT_FIELDS : []).forEach(function (field) {
            if (!field || !field.key) return;
            const fwrap = document.createElement("div");
            fwrap.className = "add-field";
            if (field.inputType === "textarea") fwrap.classList.add("add-field--notes");
            const lab = document.createElement("label");
            lab.setAttribute("for", "p1e_" + field.key);
            lab.textContent = field.label || field.key;
            fwrap.appendChild(lab);
            let input;
            if (field.inputType === "textarea") {
              input = document.createElement("textarea");
            } else if (field.inputType === "select") {
              input = document.createElement("select");
              input.className = "add-field-select--autosize";
              const opt0 = document.createElement("option");
              opt0.value = "";
              opt0.textContent = "(select)";
              input.appendChild(opt0);
            } else if (field.inputType === "date") {
              input = document.createElement("input");
              input.type = "date";
            } else {
              input = document.createElement("input");
              input.type = "text";
            }
            input.id = "p1e_" + field.key;
            input.dataset.writeKey = field.key;
            const raw = item ? valueFromItemByKeys(item, [field.key]) : "";
            if (field.inputType === "date") input.value = isoDateForDateInput(raw);
            else if (field.inputType !== "select") {
              input.value = raw !== undefined && raw !== null ? formatCellValue(raw) : "";
            }
            fwrap.appendChild(input);
            phase1EditFields.appendChild(fwrap);
          });
        }
        async function populatePhase1EditDropdowns(item) {
          const pw = hubSession.pw;
          if (!pw) return;
          const seg = phaseOneTrackingListApiPath();
          const fields = Array.isArray(PHASE_ONE_EDIT_FIELDS) ? PHASE_ONE_EDIT_FIELDS : [];
          for (let i = 0; i < fields.length; i++) {
            const field = fields[i];
            if (!field || field.inputType !== "select") continue;
            const sel = document.getElementById("p1e_" + field.key);
            if (!sel || sel.tagName !== "SELECT") continue;
            const col = { key: field.key, tryKeys: [field.key] };
            await fillDropdownSelect(sel, col, seg, pw, item, item);
            if (item) {
              const raw = valueFromItemByKeys(item, [field.key]);
              const display = formatCellValue(raw);
              ensureSelectIncludesValue(sel, display);
              sel.value = display;
            }
          }
          if (phase1EditForm) applyAllSelectAutosizes(phase1EditForm);
        }
        async function openPhase1EditPanel(item) {
          if (!item || item.Id == null) return;
          phase1Session.editItem = item;
          if (phase1EditTitle) {
            phase1EditTitle.textContent =
              "Edit Phase 1 - " + (phaseOneFieldText(item, { key: "PersonnelName", tryKeys: ["PersonnelName", "Title"] }) || "Personnel");
          }
          buildPhase1EditFormFields(item);
          await populatePhase1EditDropdowns(item);
          setPhase1EditPanelVisible(true);
          setPhase1State("", "");
        }
        async function submitPhase1EditSave() {
          const item = phase1Session.editItem;
          const pw = hubSession.pw;
          if (!item || item.Id == null || !pw) return;
          const seg = phaseOneTrackingListApiPath();
          const payload = {};
          (Array.isArray(PHASE_ONE_EDIT_FIELDS) ? PHASE_ONE_EDIT_FIELDS : []).forEach(function (field) {
            if (!field || !field.key) return;
            const el = document.getElementById("p1e_" + field.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || field.key;
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          try {
            setPhase1State("loading", "Saving...");
            if (phase1EditSaveBtn) phase1EditSaveBtn.disabled = true;
            await spFetch(`/_api/web/${seg}/items(${item.Id})`, { method: "MERGE", body: payload }, pw);
            setPhase1EditPanelVisible(false);
            phase1Session.editItem = null;
            await loadPhase1TrackerData();
            setPhase1State("ok", "Phase 1 record saved.");
            window.setTimeout(function () {
              setPhase1State("", "");
            }, 2500);
          } catch (e) {
            setPhase1State("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
          } finally {
            if (phase1EditSaveBtn) phase1EditSaveBtn.disabled = false;
          }
        }
        function phaseOneDatePayloadValue(raw) {
          if (raw === null || raw === undefined || raw === "") return null;
          const d = parseWeaponsCertCalendarDate(raw);
          if (d) return spDateTimeFromCalendarDate(d);
          const text = String(formatCellValue(raw) || "").trim();
          return text || null;
        }
        function phaseOneArchivePayloadFromItem(item) {
          const payload = {
            PersonnelName: formatCellValue(valueFromItemByKeys(item, ["PersonnelName", "Title"]) || ""),
            DutyStatus: formatCellValue(valueFromItemByKeys(item, ["DutyStatus"]) || ""),
            DutySection: formatCellValue(valueFromItemByKeys(item, ["DutySection"]) || ""),
            DateArrived: phaseOneDatePayloadValue(valueFromItemByKeys(item, ["DateArrived"])),
            PhaseOneStartDate: phaseOneDatePayloadValue(valueFromItemByKeys(item, ["PhaseOneStartDate"])),
            ProjectedOfficeFlight: formatCellValue(valueFromItemByKeys(item, ["ProjectedOfficeFlight"]) || ""),
            DateReleased: phaseOneDatePayloadValue(valueFromItemByKeys(item, ["DateReleased"])),
            Notes: formatCellValue(valueFromItemByKeys(item, ["Notes"]) || ""),
            ArchivedAt: spDateTimeNowIso(),
            Title: formatCellValue(valueFromItemByKeys(item, ["PersonnelName", "Title"]) || "Phase 1 archive"),
          };
          return payload;
        }
        async function archivePhase1Record() {
          const item = phase1Session.editItem;
          const pw = hubSession.pw;
          if (!item || item.Id == null || !pw) return;
          if (!phaseOneArchiveListTitle() && !phaseOneArchiveListUsesGuid()) {
            setPhase1State("err", "PhaseOneArchive list is not configured.");
            return;
          }
          const name = phaseOneFieldText(item, { key: "PersonnelName", tryKeys: ["PersonnelName", "Title"] }) || "this member";
          if (!confirm("Archive Phase 1 record for " + name + "? This moves the record to PhaseOneArchive.")) return;
          const trackSeg = phaseOneTrackingListApiPath();
          const archSeg = phaseOneArchiveListApiPath();
          const archivePayload = phaseOneArchivePayloadFromItem(item);
          const personId = phaseOnePersonnelIdFromItem(item);
          if (personId) {
            phase1Session.personPostKey = null;
            const archPersonKey = await resolvePhaseOnePersonPostKey(archSeg, pw);
            archivePayload[archPersonKey] = parseInt(personId, 10);
            phase1Session.personPostKey = null;
          }
          try {
            setPhase1State("loading", "Archiving...");
            if (phase1ArchiveBtn) phase1ArchiveBtn.disabled = true;
            await spFetch(`/_api/web/${archSeg}/items`, { method: "POST", body: archivePayload }, pw);
            await spFetch(
              `/_api/web/${trackSeg}/items(${item.Id})`,
              { method: "MERGE", body: { IsActive: false } },
              pw,
            );
            setPhase1EditPanelVisible(false);
            phase1Session.editItem = null;
            await loadPhase1TrackerData();
            setPhase1State("ok", "Record archived.");
            window.setTimeout(function () {
              setPhase1State("", "");
            }, 2500);
          } catch (e) {
            setPhase1State("err", "Archive failed: " + (e.message || String(e)).slice(0, 280));
          } finally {
            if (phase1ArchiveBtn) phase1ArchiveBtn.disabled = false;
          }
        }
        async function loadPhase1TrackerData() {
          const pw = hubSession.pw;
          if (!pw) return;
          setPhase1State("loading", "Loading Phase 1 records...");
          clearPhase1Table();
          if (phase1Empty) phase1Empty.hidden = true;
          try {
            const trackSeg = phaseOneTrackingListApiPath();
            const rows = await fetchPhaseOneListRows(trackSeg, pw, PHASE_ONE_TRACKING_ITEMS_ORDERBY);
            phase1Session.trackingRows = rows.slice();
            if (phaseOneArchiveListTitle() || phaseOneArchiveListUsesGuid()) {
              const archSeg = phaseOneArchiveListApiPath();
              phase1Session.archiveRows = await fetchPhaseOneListRows(archSeg, pw, PHASE_ONE_ARCHIVE_ITEMS_ORDERBY);
            } else {
              phase1Session.archiveRows = [];
            }
            refreshPhase1ListDisplay();
            const activeCount = rows.filter(function (r) {
              return phaseOneIsActive(r);
            }).length;
            setPhase1State(
              "ok",
              "Loaded " + activeCount + " active Phase 1 record(s)" +
                (phase1Session.archiveRows.length ? ", " + phase1Session.archiveRows.length + " archived." : "."),
            );
            window.setTimeout(function () {
              setPhase1State("", "");
            }, 2200);
          } catch (e) {
            phase1Session.trackingRows = null;
            phase1Session.archiveRows = null;
            refreshPhase1ListDisplay();
            setPhase1State("err", "Could not load Phase 1 data: " + (e.message || String(e)).slice(0, 280));
          }
        }

        async function fetchAllAppointmentsRows(seg, pw) {
          const orderByClause = String(APPOINTMENTS_ITEMS_ORDERBY || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=5000` + orderByQs, {}, pw);
          } catch (e0) {
            if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
              data = await spFetch(`/_api/web/${seg}/items?$top=5000`, {}, pw);
            } else {
              throw e0;
            }
          }
          const rows = (data && data.value) || [];
          hubSession.appointmentsSampleRow = rows.length ? rows[0] : hubSession.appointmentsSampleRow || null;
          return rows;
        }

        async function fetchAllArchivedAppointmentsRows(seg, pw) {
          const orderByClause = String(APPOINTMENTS_ARCHIVE_ITEMS_ORDERBY || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=5000` + orderByQs, {}, pw);
          } catch (e0) {
            if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
              data = await spFetch(`/_api/web/${seg}/items?$top=5000`, {}, pw);
            } else {
              throw e0;
            }
          }
          const rows = (data && data.value) || [];
          hubSession.appointmentsArchiveSampleRow = rows.length
            ? rows[0]
            : hubSession.appointmentsArchiveSampleRow || null;
          return rows;
        }

        function clearSchedulingTable() {
          if (schedulingThead) schedulingThead.innerHTML = "";
          if (schedulingTableBody) schedulingTableBody.innerHTML = "";
        }

        function renderSchedulingTable(rows, totalCount) {
          clearSchedulingTable();
          if (!schedulingThead || !schedulingTableBody) return;
          const isArchive = schedulingSession.listView === "archive";
          const headers = ["Office", "Personnel", "Date / time", "Location", "Description", "Instructor", "Missed"];
          if (isArchive) headers.push("Archived");
          headers.push("Created");
          const trHead = document.createElement("tr");
          headers.forEach(function (label) {
            const th = document.createElement("th");
            th.textContent = label;
            trHead.appendChild(th);
          });
          const thAct = document.createElement("th");
          thAct.className = "roster-actions";
          thAct.textContent = " ";
          thAct.title = isArchive ? "Memorandum" : "Mark missed / memorandum";
          trHead.appendChild(thAct);
          schedulingThead.appendChild(trHead);

          const views = (Array.isArray(rows) ? rows : []).map(function (entry) {
            const normalized = normalizeSchedulingTableEntry(entry);
            const view = schedulingRowViewModel(normalized.item);
            view.pendingArchive = normalized.pendingArchive;
            return view;
          });
          views.sort(function (a, b) {
            if (isArchive && !!a.pendingArchive !== !!b.pendingArchive) {
              return a.pendingArchive ? -1 : 1;
            }
            const oa = String(a.office || "").toLowerCase();
            const ob = String(b.office || "").toLowerCase();
            if (oa !== ob) return oa < ob ? -1 : oa > ob ? 1 : 0;
            const ta = a.whenSort ? a.whenSort.getTime() : 0;
            const tb = b.whenSort ? b.whenSort.getTime() : 0;
            if (ta !== tb) return isArchive ? tb - ta : ta - tb;
            return String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" });
          });

          const frag = document.createDocumentFragment();
          views.forEach(function (view) {
            const tr = document.createElement("tr");
            if (view.pendingArchive) tr.className = "scheduling-row--pending-archive";
            else if (view.archiveRetentionTier === "red") tr.className = "scheduling-row--archive-delete-red";
            else if (view.archiveRetentionTier === "yellow") tr.className = "scheduling-row--archive-delete-yellow";
            else if (view.missedFlag) tr.className = "scheduling-row--missed";
            const cells = [
              view.office,
              view.name,
              view.when,
              view.location,
              view.description,
              view.instructor,
              view.missed,
            ];
            if (isArchive) {
              cells.push(appointmentArchiveStatusLabel(view.item, view.pendingArchive));
            }
            cells.push(view.createdAt);
            cells.forEach(function (text) {
              const td = document.createElement("td");
              td.textContent = displayCellText(text);
              tr.appendChild(td);
            });
            tr.appendChild(
              buildAppointmentActionsCell(
                view.item,
                async function () {
                  await loadSchedulingAppointmentsList();
                },
                { memoOnly: isArchive },
              ),
            );
            frag.appendChild(tr);
          });
          schedulingTableBody.appendChild(frag);
          if (typeof totalCount === "number") updateSchedulingListTitle(rows.length, totalCount);
          else if (schedulingEmpty) schedulingEmpty.hidden = views.length > 0;
        }

        async function loadSchedulingAppointmentsList() {
          const pw = hubSession.pw;
          const seg = appointmentsListApiPath();
          if (!pw) return;
          setSchedulingState("loading", "Loading appointments...");
          clearSchedulingTable();
          if (schedulingEmpty) schedulingEmpty.hidden = true;
          try {
            let rows = await fetchAllAppointmentsRows(seg, pw);
            const archiveResult = await archivePastAppointmentItems(seg, pw, rows);
            rows = archiveResult.remaining;
            schedulingSession.rows = rows.slice();
            if (appointmentsArchiveListTitle() || appointmentsArchiveListUsesGuid()) {
              const archSeg = appointmentsArchiveListApiPath();
              let archiveRows = await fetchAllArchivedAppointmentsRows(archSeg, pw);
              const purgeResult = await purgeExpiredArchivedAppointments(archSeg, pw, archiveRows);
              archiveRows = purgeResult.remaining;
              schedulingSession.archiveRows = archiveRows;
              if (purgeResult.purged > 0) {
                schedulingSession.lastArchivePurged = purgeResult.purged;
              } else {
                schedulingSession.lastArchivePurged = 0;
              }
            } else {
              schedulingSession.archiveRows = [];
              schedulingSession.lastArchivePurged = 0;
            }
            refreshSchedulingListDisplay();
            const upcomingCount = rows.filter(appointmentIsUpcoming).length;
            const pendingCount = rows.filter(appointmentIsPast).length;
            const missedCount = rows.filter(function (r) {
              return appointmentIsUpcoming(r) && appointmentIsMissed(r);
            }).length;
            let msg =
              "Loaded " + upcomingCount + " upcoming appointment(s)" +
              (missedCount ? " (" + missedCount + " missed)." : ".");
            if (pendingCount) {
              msg += " " + pendingCount + " past on file pending archive.";
            }
            if (schedulingSession.archiveRows.length) {
              msg += " " + schedulingSession.archiveRows.length + " archived.";
            }
            if (schedulingSession.lastArchivePurged > 0) {
              msg +=
                " Removed " + schedulingSession.lastArchivePurged + " expired archived appointment(s) (30+ days).";
            }
            if (archiveResult.archived > 0) {
              msg += " Moved " + archiveResult.archived + " past appointment(s) to archive.";
            } else if (archiveResult.pastHidden > 0) {
              msg +=
                " " +
                archiveResult.pastHidden +
                " past appointment(s) still on file (create the AppointmentsArchive list to move them).";
            }
            if (archiveResult.failures && archiveResult.failures.length) {
              msg +=
                " " +
                archiveResult.failures.length +
                " past appointment(s) could not be archived - check the Archive tab (highlighted in yellow).";
              setSchedulingState("warn", msg);
            } else {
              setSchedulingState("ok", msg);
            }
            window.setTimeout(function () {
              setSchedulingState("", "");
            }, 2200);
          } catch (e) {
            schedulingSession.rows = null;
            schedulingSession.archiveRows = null;
            refreshSchedulingListDisplay();
            setSchedulingState("err", "Could not load appointments: " + (e.message || String(e)).slice(0, 220));
          }
        }

        function setSchedulingMemoOverlayVisible(show) {
          if (schedulingMemoOverlay) schedulingMemoOverlay.hidden = !show;
          if (!show && schedulingMemoForm) schedulingMemoForm.reset();
          if (show) {
            populateMemoSignatureDropdowns();
            void populateMemoLetterheadFields();
          }
        }

        function hubLetterheadGuidRaw() {
          return String(LIST_HUB_LETTERHEAD_GUID || "").trim();
        }

        function hubLetterheadListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(
            hubLetterheadGuidRaw(),
          );
        }

        function hubLetterheadListApiPath() {
          if (hubLetterheadListUsesGuid()) return "lists(guid'" + hubLetterheadGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(String(LIST_HUB_LETTERHEAD || "HubLetterhead")) + "')";
        }

        function resolveMemoLetterheadPdfUrl() {
          const explicit =
            typeof MEMO_LETTERHEAD_PDF_URL === "string" ? MEMO_LETTERHEAD_PDF_URL.trim() : "";
          if (explicit) return explicit;
          const root = (typeof PERSONNEL_SITE_ROOT_URL === "string" ? PERSONNEL_SITE_ROOT_URL : "")
            .trim()
            .replace(/\/$/, "");
          const rel = (
            typeof MEMO_LETTERHEAD_PDF_SITE_RELATIVE_PATH === "string"
              ? MEMO_LETTERHEAD_PDF_SITE_RELATIVE_PATH
              : ""
          )
            .trim()
            .replace(/^\//, "");
          if (!root || !rel) return "";
          return root + "/" + rel;
        }

        function updateMemoLetterheadPdfLink() {
          if (!schedulingMemoLetterheadPdfLink) return;
          const url = resolveMemoLetterheadPdfUrl();
          if (!url) {
            schedulingMemoLetterheadPdfLink.hidden = true;
            schedulingMemoLetterheadPdfLink.removeAttribute("href");
            return;
          }
          schedulingMemoLetterheadPdfLink.href = url;
          schedulingMemoLetterheadPdfLink.hidden = false;
        }

        function memoLetterheadFromStorage() {
          try {
            const raw = localStorage.getItem(MEMO_LETTERHEAD_STORAGE_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            if (!parsed || typeof parsed !== "object") return null;
            return {
              topMarginIn:
                parsed.topMarginIn != null && parsed.topMarginIn !== ""
                  ? Number(parsed.topMarginIn)
                  : DEFAULT_MEMO_PRINT_TOP_MARGIN_IN,
              footer: String(parsed.footer || ""),
            };
          } catch (e) {
            return null;
          }
        }

        function saveMemoLetterheadToStorage(topMarginIn, footer) {
          try {
            localStorage.setItem(
              MEMO_LETTERHEAD_STORAGE_KEY,
              JSON.stringify({
                topMarginIn: topMarginIn != null ? Number(topMarginIn) : DEFAULT_MEMO_PRINT_TOP_MARGIN_IN,
                footer: String(footer || ""),
              }),
            );
          } catch (e) {
            /* ignore */
          }
        }

        async function loadMemoLetterhead(pw, force) {
          if (hubSession.memoLetterheadLoaded && !force) return hubSession.memoLetterhead;
          let topMarginIn = DEFAULT_MEMO_PRINT_TOP_MARGIN_IN;
          let footer = String(DEFAULT_MEMO_LETTERHEAD_FOOTER || "");
          const stored = memoLetterheadFromStorage();
          if (stored) {
            if (stored.topMarginIn != null && !isNaN(stored.topMarginIn)) topMarginIn = stored.topMarginIn;
            if (stored.footer !== undefined) footer = stored.footer;
          }
          hubSession.memoLetterhead = { topMarginIn: topMarginIn, footer: footer };
          hubSession.memoLetterheadItemId = null;
          hubSession.memoLetterheadUsesSharePoint = false;
          hubSession.memoLetterheadLoaded = true;
          return hubSession.memoLetterhead;
        }

        async function saveMemoLetterhead(pw, topMarginIn, footer) {
          topMarginIn =
            topMarginIn != null && topMarginIn !== "" ? Number(topMarginIn) : DEFAULT_MEMO_PRINT_TOP_MARGIN_IN;
          footer = String(footer || "");
          saveMemoLetterheadToStorage(topMarginIn, footer);
          hubSession.memoLetterhead = { topMarginIn: topMarginIn, footer: footer };
          hubSession.memoLetterheadLoaded = true;
          return { savedToSharePoint: false };
        }

        function memoLetterheadFieldText(item, keys) {
          const raw = valueFromItemByKeys(item, keys);
          if (raw === null || raw === undefined) return "";
          return String(raw);
        }

        function mqlSignatureFromStorage() {
          try {
            const raw = localStorage.getItem(MQL_SIGNATURE_STORAGE_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            if (!parsed || typeof parsed !== "object") return null;
            return {
              nameLine: String(parsed.nameLine || ""),
              titleLine: String(parsed.titleLine || ""),
            };
          } catch (e) {
            return null;
          }
        }

        function saveMqlSignatureToStorage(nameLine, titleLine) {
          try {
            localStorage.setItem(
              MQL_SIGNATURE_STORAGE_KEY,
              JSON.stringify({
                nameLine: String(nameLine || ""),
                titleLine: String(titleLine || ""),
              }),
            );
          } catch (e) {
            /* ignore */
          }
        }

        async function loadMqlSignature(pw, force) {
          if (hubSession.mqlSignatureLoaded && !force) return hubSession.mqlSignature;
          let nameLine = String(DEFAULT_MQL_SIGNATURE_NAME || "");
          let titleLine = String(DEFAULT_MQL_SIGNATURE_TITLE || "");
          let usesSharePoint = false;
          if (pw) {
            try {
              const seg = hubLetterheadListApiPath();
              const data = await spFetch(`/_api/web/${seg}/items?$top=1&$orderby=Id asc`, {}, pw);
              const item = data && data.value && data.value[0];
              if (item) {
                const n = memoLetterheadFieldText(item, [
                  "MqlSignatureName",
                  "MqlSigName",
                  "MqlSignature_x0020_Name",
                  "MqlSignatureLine",
                ]);
                const t = memoLetterheadFieldText(item, [
                  "MqlSignatureTitle",
                  "MqlSigTitle",
                  "MqlSignature_x0020_Title",
                ]);
                if (n) nameLine = n;
                if (t) titleLine = t;
                hubSession.mqlSignatureItemId = item.Id;
                usesSharePoint = !!(n || t);
              }
            } catch (e) {
              log("MQL signature load skipped:\n" + (e.message || String(e)), "warn");
            }
          }
          if (!usesSharePoint) {
            const stored = mqlSignatureFromStorage();
            if (stored) {
              if (stored.nameLine) nameLine = stored.nameLine;
              if (stored.titleLine) titleLine = stored.titleLine;
            }
          }
          hubSession.mqlSignature = { nameLine: nameLine, titleLine: titleLine };
          hubSession.mqlSignatureUsesSharePoint = usesSharePoint;
          hubSession.mqlSignatureLoaded = true;
          return hubSession.mqlSignature;
        }

        async function saveMqlSignature(pw, nameLine, titleLine) {
          nameLine = String(nameLine || "");
          titleLine = String(titleLine || "");
          let savedToSharePoint = false;
          if (pw) {
            try {
              const seg = hubLetterheadListApiPath();
              const payload = {
                Title: "Default",
                MqlSignatureName: nameLine,
                MqlSignatureTitle: titleLine,
              };
              if (hubSession.mqlSignatureItemId) {
                await spFetch(
                  `/_api/web/${seg}/items(${hubSession.mqlSignatureItemId})`,
                  { method: "MERGE", body: payload },
                  pw,
                );
              } else if (hubSession.memoLetterheadItemId) {
                await spFetch(
                  `/_api/web/${seg}/items(${hubSession.memoLetterheadItemId})`,
                  { method: "MERGE", body: payload },
                  pw,
                );
                hubSession.mqlSignatureItemId = hubSession.memoLetterheadItemId;
              } else {
                const created = await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
                if (created && created.Id != null) hubSession.mqlSignatureItemId = created.Id;
              }
              savedToSharePoint = true;
              hubSession.mqlSignatureUsesSharePoint = true;
            } catch (e) {
              log("MQL signature save to HubLetterhead failed; using browser storage:\n" + (e.message || String(e)), "warn");
            }
          }
          saveMqlSignatureToStorage(nameLine, titleLine);
          hubSession.mqlSignature = { nameLine: nameLine, titleLine: titleLine };
          hubSession.mqlSignatureLoaded = true;
          return { savedToSharePoint: savedToSharePoint };
        }

        function currentMqlSignatureForPrint(panel) {
          const root = panel || document;
          const nameInput = root.querySelector ? root.querySelector(".mql-sig-name-input") : null;
          const titleInput = root.querySelector ? root.querySelector(".mql-sig-title-input") : null;
          const cached = hubSession.mqlSignature || {};
          return {
            nameLine:
              (nameInput ? String(nameInput.value || "").trim() : "") ||
              cached.nameLine ||
              String(DEFAULT_MQL_SIGNATURE_NAME || ""),
            titleLine:
              (titleInput ? String(titleInput.value || "").trim() : "") ||
              cached.titleLine ||
              String(DEFAULT_MQL_SIGNATURE_TITLE || ""),
          };
        }

        function buildMqlSignatureSettingsPanel() {
          const details = document.createElement("details");
          details.className = "reports-mql-signature-panel";
          details.open = false;
          const summary = document.createElement("summary");
          summary.textContent = "Signature block (signed MQL only)";
          details.appendChild(summary);
          const hint = document.createElement("p");
          hint.className = "reports-hint";
          hint.textContent =
            "Edit the commander signature shown at the bottom right of each printed page. Saves to HubLetterhead when available.";
          details.appendChild(hint);
          const fields = document.createElement("div");
          fields.className = "reports-mql-signature-fields";
          const nameField = document.createElement("div");
          nameField.className = "add-field";
          const nameLabel = document.createElement("label");
          nameLabel.textContent = "Name / rank line";
          const nameInput = document.createElement("input");
          nameInput.type = "text";
          nameInput.className = "mql-sig-name-input";
          nameInput.placeholder = DEFAULT_MQL_SIGNATURE_NAME;
          nameField.appendChild(nameLabel);
          nameField.appendChild(nameInput);
          const titleField = document.createElement("div");
          titleField.className = "add-field";
          const titleLabel = document.createElement("label");
          titleLabel.textContent = "Title line";
          const titleInput = document.createElement("input");
          titleInput.type = "text";
          titleInput.className = "mql-sig-title-input";
          titleInput.placeholder = DEFAULT_MQL_SIGNATURE_TITLE;
          titleField.appendChild(titleLabel);
          titleField.appendChild(titleInput);
          fields.appendChild(nameField);
          fields.appendChild(titleField);
          details.appendChild(fields);
          const saveBtn = document.createElement("button");
          saveBtn.type = "button";
          saveBtn.className = "btn-secondary mql-sig-save-btn";
          saveBtn.textContent = "Save signature block";
          details.appendChild(saveBtn);
          saveBtn.addEventListener("click", function () {
            void (async function () {
              const pw = hubSession.pw || ethosSession.pw;
              if (!pw) {
                setReportsState("warn", "Load the Personnel Roster first.");
                return;
              }
              try {
                setReportsState("loading", "Saving signature block...");
                const result = await saveMqlSignature(
                  pw,
                  nameInput.value,
                  titleInput.value,
                );
                setReportsState(
                  "ok",
                  result.savedToSharePoint
                    ? "Signature block saved to HubLetterhead."
                    : "Signature block saved in this browser.",
                );
                window.setTimeout(function () {
                  setReportsState("", "");
                }, 2800);
              } catch (e) {
                setReportsState("err", "Could not save signature: " + (e.message || String(e)).slice(0, 200));
              }
            })();
          });
          return details;
        }

        async function populateMqlSignaturePanel(panel) {
          if (!panel) return;
          const pw = hubSession.pw || ethosSession.pw;
          const sig = await loadMqlSignature(pw, false);
          const nameInput = panel.querySelector(".mql-sig-name-input");
          const titleInput = panel.querySelector(".mql-sig-title-input");
          if (nameInput) nameInput.value = sig.nameLine || "";
          if (titleInput) titleInput.value = sig.titleLine || "";
        }

        function buildMqlSignedPrintHeader(title, options) {
          options = options || {};
          const header = document.createElement("div");
          header.className = "mql-signed-print-header";
          const top = document.createElement("div");
          top.className = "mql-signed-print-header-top";
          const unit = document.createElement("span");
          unit.className = "mql-signed-print-unit";
          unit.textContent = options.unitLabel || String(AF_REPORT_UNIT_LABEL || "88th Security Forces Squadron");
          const dtWrap = document.createElement("span");
          dtWrap.className = "mql-signed-print-datetime";
          const now =
            options.generatedAt instanceof Date && !isNaN(options.generatedAt.getTime())
              ? options.generatedAt
              : new Date();
          const datePart = now.toLocaleDateString("en-US", {
            weekday: "short",
            month: "short",
            day: "numeric",
            year: "numeric",
          });
          const timePart = now.toLocaleTimeString("en-US", {
            hour: "numeric",
            minute: "2-digit",
            second: "2-digit",
            hour12: true,
          });
          dtWrap.textContent = datePart + "  " + timePart;
          top.appendChild(unit);
          top.appendChild(dtWrap);
          header.appendChild(top);
          const h1 = document.createElement("h1");
          h1.className = "mql-signed-print-title";
          h1.textContent = title || MQL_REPORT_PRINT_TITLE;
          header.appendChild(h1);
          const sortLine = document.createElement("p");
          sortLine.className = "mql-signed-print-sortline";
          sortLine.textContent = options.sortLine || MQL_SIGNED_SORT_LINE;
          header.appendChild(sortLine);
          return header;
        }

        function buildMqlPrintSignatureBlock(sig, asPageStamp) {
          sig = sig || hubSession.mqlSignature || {};
          const block = document.createElement("div");
          block.className = asPageStamp ? "mql-print-signature-stamp" : "mql-print-doc-signature";
          if (asPageStamp) {
            const signSpace = document.createElement("div");
            signSpace.className = "mql-print-sig-signing-space";
            signSpace.setAttribute("aria-hidden", "true");
            block.appendChild(signSpace);
          }
          const line = document.createElement("div");
          line.className = "mql-print-sig-line";
          block.appendChild(line);
          const name = document.createElement("p");
          name.className = "mql-print-sig-name";
          name.textContent = sig.nameLine || String(DEFAULT_MQL_SIGNATURE_NAME || "");
          block.appendChild(name);
          const title = document.createElement("p");
          title.className = "mql-print-sig-title";
          title.textContent = sig.titleLine || String(DEFAULT_MQL_SIGNATURE_TITLE || "");
          block.appendChild(title);
          return block;
        }

        async function populateMemoLetterheadFields() {
          const pw = hubSession.pw;
          const letterhead = await loadMemoLetterhead(pw, false);
          if (schedulingMemoPrintTopMargin) {
            schedulingMemoPrintTopMargin.value =
              letterhead.topMarginIn != null ? String(letterhead.topMarginIn) : String(DEFAULT_MEMO_PRINT_TOP_MARGIN_IN);
          }
          if (schedulingMemoLetterheadFooter) {
            schedulingMemoLetterheadFooter.value = letterhead.footer || "";
          }
          updateMemoLetterheadPdfLink();
        }

        function currentMemoLetterheadForPrint() {
          const marginRaw = schedulingMemoPrintTopMargin
            ? String(schedulingMemoPrintTopMargin.value || "").trim()
            : "";
          const footer = schedulingMemoLetterheadFooter
            ? String(schedulingMemoLetterheadFooter.value || "")
            : "";
          const cached = hubSession.memoLetterhead || {};
          const topMarginIn =
            marginRaw !== ""
              ? Number(marginRaw)
              : cached.topMarginIn != null
                ? cached.topMarginIn
                : DEFAULT_MEMO_PRINT_TOP_MARGIN_IN;
          return {
            topMarginIn: isNaN(topMarginIn) ? DEFAULT_MEMO_PRINT_TOP_MARGIN_IN : topMarginIn,
            footer: footer || cached.footer || String(DEFAULT_MEMO_LETTERHEAD_FOOTER || ""),
            pdfUrl: resolveMemoLetterheadPdfUrl(),
          };
        }

        function applyMqlCertCellStyle(td, tone) {
          if (!td || !tone || tone === "unknown" || tone === "ok") return;
          td.classList.add("mql-cert-cell", "mql-cert-cell--" + tone);
        }

        function setMqlLandscapePrintMode(enabled) {
          let el = document.getElementById(MQL_PRINT_STYLE_ID);
          if (enabled) {
            if (!el) {
              el = document.createElement("style");
              el.id = MQL_PRINT_STYLE_ID;
              el.textContent =
                "@media print {" +
                "@page { size: landscape; margin: 0.35in 0.4in 1.5in 0.4in; }" +
                "body { margin: 0; }" +
                "}";
              document.head.appendChild(el);
            }
          } else if (el) {
            el.remove();
          }
        }

        function defaultMemoSignatureBlocks() {
          return {
            sig1Title: "Training NCOIC",
            sig1Squadron: String(APPOINTMENTS_SQUADRON_LABEL || "88 SFS"),
            sig2Title: "Supervisor",
            sig3Title: "Commander",
            sig3Squadron: String(APPOINTMENTS_SQUADRON_LABEL || "88 SFS"),
          };
        }

        function collectPersonnelFieldValues(fieldKey) {
          const values = new Set();
          const rows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
          rows.forEach(function (person) {
            const v = String(itemFieldText(person, fieldKey) || "").trim();
            if (v) values.add(v);
          });
          return Array.from(values).sort(function (a, b) {
            return a.localeCompare(b, undefined, { sensitivity: "base" });
          });
        }

        function memoTitleOptions() {
          const set = new Set(Array.isArray(MEMO_SIGNATURE_TITLE_PRESETS) ? MEMO_SIGNATURE_TITLE_PRESETS : []);
          collectPersonnelFieldValues("OfficeSymbol").forEach(function (v) {
            set.add(v);
          });
          collectPersonnelFieldValues("Status").forEach(function (v) {
            set.add(v);
          });
          return Array.from(set).sort(function (a, b) {
            return a.localeCompare(b, undefined, { sensitivity: "base" });
          });
        }

        function memoSquadronOptions() {
          const set = new Set([String(APPOINTMENTS_SQUADRON_LABEL || "88 SFS")]);
          collectPersonnelFieldValues("Squadron").forEach(function (v) {
            set.add(v);
          });
          return Array.from(set).sort(function (a, b) {
            return a.localeCompare(b, undefined, { sensitivity: "base" });
          });
        }

        function memoSignatureSquadronFromPerson(person) {
          if (!person) return String(APPOINTMENTS_SQUADRON_LABEL || "88 SFS");
          return itemFieldText(person, "Squadron") || String(APPOINTMENTS_SQUADRON_LABEL || "88 SFS");
        }

        function memoSignatureTitleFromPerson(person) {
          if (!person) return "";
          return itemFieldText(person, "OfficeSymbol") || itemFieldText(person, "Status") || "";
        }

        function setMemoSelectValue(sel, value) {
          if (!sel) return;
          const v = String(value || "").trim();
          ensureSelectIncludesValue(sel, v);
          sel.value = v;
        }

        function populateMemoSelectOptions(sel, options, emptyLabel) {
          if (!sel) return;
          const prior = String(sel.value || "").trim();
          while (sel.options.length > 1) sel.remove(1);
          (Array.isArray(options) ? options : []).forEach(function (opt) {
            const text = String(opt || "").trim();
            if (!text) return;
            const o = document.createElement("option");
            o.value = text;
            o.textContent = text;
            sel.appendChild(o);
          });
          if (prior) setMemoSelectValue(sel, prior);
          else if (emptyLabel && sel.options.length) sel.options[0].textContent = emptyLabel;
        }

        function applyMemoSignatureFieldsFromPerson(person, titleSelect, squadronSelect) {
          if (!person) return;
          setMemoSelectValue(titleSelect, memoSignatureTitleFromPerson(person));
          setMemoSelectValue(squadronSelect, memoSignatureSquadronFromPerson(person));
        }

        function populateMemoSignatureDropdowns() {
          const titleOptions = memoTitleOptions();
          const squadronOptions = memoSquadronOptions();
          const rows = Array.isArray(hubSession.rows) ? hubSession.rows.slice() : [];
          rows.sort(function (a, b) {
            return formatPersonDisplayName(a).localeCompare(formatPersonDisplayName(b), undefined, {
              sensitivity: "base",
            });
          });
          schedulingMemoSignatureBlocks.forEach(function (block) {
            populateMemoSelectOptions(block.titleSelect, titleOptions, "Select title...");
            populateMemoSelectOptions(block.squadronSelect, squadronOptions, "Select squadron...");
            const sel = block.personSelect;
            if (!sel) return;
            const prior = String(sel.value || "").trim();
            while (sel.options.length > 1) sel.remove(1);
            rows.forEach(function (person) {
              if (!person || person.Id == null) return;
              const displayName = formatPersonDisplayName(person);
              if (!displayName) return;
              const o = document.createElement("option");
              o.value = String(person.Id);
              o.textContent = displayName;
              sel.appendChild(o);
            });
            if (prior) sel.value = prior;
          });
        }

        function setMemoSignatureBlock(block, opts) {
          if (!block) return;
          const defaults = defaultMemoSignatureBlocks();
          const blockDefaults = [
            { title: defaults.sig1Title, squadron: defaults.sig1Squadron },
            { title: defaults.sig2Title, squadron: "" },
            { title: defaults.sig3Title, squadron: defaults.sig3Squadron },
          ];
          const blockIndex = schedulingMemoSignatureBlocks.indexOf(block);
          const fallback = blockDefaults[blockIndex] || { title: "", squadron: "" };
          if (block.personSelect) {
            block.personSelect.value = opts && opts.personId ? String(opts.personId) : "";
          }
          setMemoSelectValue(block.titleSelect, (opts && opts.title) || fallback.title || "");
          setMemoSelectValue(block.squadronSelect, (opts && opts.squadron) || fallback.squadron || "");
          if (block.personSelect && block.personSelect.value) {
            const person = personnelById(block.personSelect.value);
            if (person) applyMemoSignatureFieldsFromPerson(person, block.titleSelect, block.squadronSelect);
          }
        }

        function fillSchedulingMemoForm(view, options) {
          const opts = options || {};
          const office = view ? view.office : "";
          const name = view ? view.name : "";
          const defaults = defaultMemoSignatureBlocks();
          if (schedulingMemoFor) {
            schedulingMemoFor.value = office && office !== "Unassigned" ? office + " Leadership" : "";
          }
          if (schedulingMemoFrom) {
            schedulingMemoFrom.value = String(APPOINTMENTS_MEMO_DEFAULT_FROM || "88 SFS/S3T");
          }
          if (schedulingMemoSubject) {
            schedulingMemoSubject.value = name
              ? "Missed Training Appointment - " + name
              : "Missed Training Appointment";
          }
          if (schedulingMemoBody) {
            if (view) {
              schedulingMemoBody.value =
                "1. On " +
                (view.when || "the scheduled date") +
                ", " +
                name +
                " failed to report for a scheduled training appointment" +
                (view.location ? " at " + view.location : "") +
                "." +
                (view.description ? " Purpose: " + view.description + "." : "") +
                "\n\n2. Provide corrective action and document follow-up in accordance with unit policy.";
            } else {
              schedulingMemoBody.value =
                "1. [State the facts of the missed appointment.]\n\n2. [State recommended actions or follow-up.]";
            }
          }
          setMemoSignatureBlock(schedulingMemoSignatureBlocks[0], {
            personId: opts.sig1PersonId,
            title: opts.sig1Title || defaults.sig1Title,
            squadron: opts.sig1Squadron || defaults.sig1Squadron,
          });
          setMemoSignatureBlock(schedulingMemoSignatureBlocks[1], {
            personId: opts.sig2PersonId,
            title: opts.sig2Title || defaults.sig2Title,
            squadron: opts.sig2Squadron || "",
          });
          setMemoSignatureBlock(schedulingMemoSignatureBlocks[2], {
            personId: opts.sig3PersonId,
            title: opts.sig3Title || defaults.sig3Title,
            squadron: opts.sig3Squadron || defaults.sig3Squadron,
          });
        }

        function openSchedulingMemoForAppointment(item) {
          if (!item) return;
          const view = schedulingRowViewModel(item);
          setSchedulingMemoOverlayVisible(true);
          fillSchedulingMemoForm(view, {});
          if (schedulingMemoFor) schedulingMemoFor.focus();
        }

        function openBlankSchedulingMemo() {
          setSchedulingMemoOverlayVisible(true);
          fillSchedulingMemoForm(null, {});
          if (schedulingMemoFor) schedulingMemoFor.focus();
        }

        function readMemoSignatureBlockData(block) {
          function val(el) {
            return el ? String(el.value || "").trim() : "";
          }
          if (!block) return { name: "", title: "", org: "" };
          const person =
            block.personSelect && block.personSelect.value
              ? personnelById(block.personSelect.value)
              : null;
          return {
            name: person ? formatPersonDisplayName(person) : "",
            title: val(block.titleSelect),
            org: val(block.squadronSelect),
          };
        }

        function readSchedulingMemoFormData() {
          function val(el) {
            return el ? String(el.value || "").trim() : "";
          }
          return {
            memorandumFor: val(schedulingMemoFor),
            from: val(schedulingMemoFrom),
            subject: val(schedulingMemoSubject),
            body: val(schedulingMemoBody),
            signatures: schedulingMemoSignatureBlocks.map(readMemoSignatureBlockData),
          };
        }

        function renderSchedulingMemoPrintDocument(data) {
          const wrap = document.createElement("div");
          wrap.className = "scheduling-memo-document";
          const letterhead = currentMemoLetterheadForPrint();
          if (letterhead.topMarginIn > 0) {
            wrap.style.paddingTop = String(letterhead.topMarginIn) + "in";
          }
          const title = document.createElement("p");
          title.className = "scheduling-memo-doc-title";
          title.textContent = "MEMORANDUM";
          wrap.appendChild(title);

          function headerRow(label, value) {
            const row = document.createElement("div");
            row.className = "scheduling-memo-doc-row";
            const lab = document.createElement("span");
            lab.className = "scheduling-memo-doc-label";
            lab.textContent = label;
            const val = document.createElement("span");
            val.className = "scheduling-memo-doc-value";
            val.textContent = value || " ";
            row.appendChild(lab);
            row.appendChild(val);
            return row;
          }

          wrap.appendChild(headerRow("Memorandum for", data.memorandumFor));
          wrap.appendChild(headerRow("From", data.from));
          wrap.appendChild(headerRow("Subject", data.subject));

          const body = document.createElement("div");
          body.className = "scheduling-memo-doc-body";
          body.textContent = data.body || "";
          wrap.appendChild(body);

          const sigWrap = document.createElement("div");
          sigWrap.className = "scheduling-memo-doc-sigs";
          (data.signatures || []).forEach(function (sig) {
            if (!sig || (!sig.name && !sig.title && !sig.org)) return;
            const block = document.createElement("div");
            block.className = "scheduling-memo-doc-sig";
            const line = document.createElement("div");
            line.className = "scheduling-memo-doc-sig-line";
            block.appendChild(line);
            if (sig.name) {
              const name = document.createElement("p");
              name.className = "scheduling-memo-doc-sig-name";
              name.textContent = sig.name;
              block.appendChild(name);
            }
            if (sig.title) {
              const titleLine = document.createElement("p");
              titleLine.className = "scheduling-memo-doc-sig-title";
              titleLine.textContent = sig.title;
              block.appendChild(titleLine);
            }
            if (sig.org) {
              const org = document.createElement("p");
              org.className = "scheduling-memo-doc-sig-org";
              org.textContent = sig.org;
              block.appendChild(org);
            }
            sigWrap.appendChild(block);
          });
          wrap.appendChild(sigWrap);
          if (letterhead.footer) {
            const letterheadFooter = document.createElement("div");
            letterheadFooter.className = "scheduling-memo-doc-letterhead-footer";
            letterheadFooter.textContent = letterhead.footer;
            wrap.appendChild(letterheadFooter);
          }
          return wrap;
        }

        function triggerSchedulingPrint(node, options) {
          options = options || {};
          if (!schedulingPrintSurface || !node) return;
          const root = document.getElementById("sp-pip-ui");
          schedulingPrintSurface.className = "scheduling-print-surface";
          if (options.mqlLandscape) {
            schedulingPrintSurface.classList.add("scheduling-print-surface--mql");
            setMqlLandscapePrintMode(true);
          }
          schedulingPrintSurface.innerHTML = "";
          schedulingPrintSurface.appendChild(node);
          if (options.mqlSignature) {
            schedulingPrintSurface.appendChild(buildMqlPrintSignatureBlock(options.mqlSignature, true));
          }
          schedulingPrintSurface.hidden = false;
          if (root) root.classList.add("scheduling-print-active");
          window.setTimeout(function () {
            window.print();
            window.setTimeout(function () {
              if (root) root.classList.remove("scheduling-print-active");
              setMqlLandscapePrintMode(false);
              schedulingPrintSurface.className = "scheduling-print-surface";
              schedulingPrintSurface.hidden = true;
            }, 500);
          }, 120);
        }

        function triggerMqlPrint(node, options) {
          options = options || {};
          const signature =
            options.signed
              ? options.signature || hubSession.mqlSignature || { nameLine: DEFAULT_MQL_SIGNATURE_NAME, titleLine: DEFAULT_MQL_SIGNATURE_TITLE }
              : null;
          triggerSchedulingPrint(node, {
            mqlLandscape: true,
            mqlSignature: signature,
          });
        }

        function printSchedulingMemo() {
          const data = readSchedulingMemoFormData();
          if (!data.subject && !data.body) {
            setSchedulingState("warn", "Enter a subject or body before printing the memorandum.");
            return;
          }
          setSchedulingMemoOverlayVisible(false);
          triggerSchedulingPrint(renderSchedulingMemoPrintDocument(data));
        }

        function setSchedulingState(kind, message) {
          if (!schedulingReadState) return;
          if (!message) {
            schedulingReadState.hidden = true;
            schedulingReadState.textContent = "";
            return;
          }
          schedulingReadState.hidden = false;
          schedulingReadState.className = "read-state " + kind;
          schedulingReadState.textContent = message;
        }

        function setSchedulingAddPanelVisible(show) {
          if (schedulingAddPanel) schedulingAddPanel.hidden = !show;
          if (!show && schedulingAddForm) schedulingAddForm.reset();
        }

        function populateSchedulingPersonSelect() {
          if (!schedulingPersonSelect) return;
          const prior = String(schedulingPersonSelect.value || "").trim();
          while (schedulingPersonSelect.options.length > 1) schedulingPersonSelect.remove(1);
          const rows = Array.isArray(hubSession.rows) ? hubSession.rows.slice() : [];
          rows.sort(function (a, b) {
            return formatPersonDisplayName(a).localeCompare(formatPersonDisplayName(b), undefined, { sensitivity: "base" });
          });
          rows.forEach(function (person) {
            if (!person || person.Id == null) return;
            const displayName = formatPersonDisplayName(person);
            if (!displayName) return;
            const o = document.createElement("option");
            o.value = String(person.Id);
            o.textContent = displayName;
            schedulingPersonSelect.appendChild(o);
          });
          if (prior) schedulingPersonSelect.value = prior;
          applySelectAutosize(schedulingPersonSelect);
        }

        function buildSchedulingAddFormFields() {
          if (!schedulingAddFields) return;
          schedulingAddFields.innerHTML = "";
          normalizedAppointmentsFormFields().forEach(function (field) {
            const fwrap = document.createElement("div");
            fwrap.className = "add-field";
            const lab = document.createElement("label");
            lab.setAttribute("for", "sf_" + field.key);
            lab.textContent = field.label + (field.required ? " *" : "");
            fwrap.appendChild(lab);
            let input;
            if (field.inputType === "checkbox") {
              input = document.createElement("input");
              input.type = "checkbox";
              input.id = "sf_" + field.key;
              input.checked = false;
            } else if (field.inputType === "datetime-local") {
              input = document.createElement("input");
              input.type = "datetime-local";
              input.id = "sf_" + field.key;
              if (field.required) input.required = true;
            } else {
              input = document.createElement("input");
              input.type = "text";
              input.id = "sf_" + field.key;
              if (field.required) input.required = true;
            }
            fwrap.appendChild(input);
            schedulingAddFields.appendChild(fwrap);
          });
        }

        async function submitSchedulingAppointmentAdd(ev) {
          if (ev && ev.preventDefault) ev.preventDefault();
          const pw = hubSession.pw;
          const seg = appointmentsListApiPath();
          if (!pw || !seg) {
            setSchedulingState("err", "Hub not ready. Refresh the roster first.");
            return;
          }
          const personId = schedulingPersonSelect
            ? parseInt(String(schedulingPersonSelect.value || "").trim(), 10)
            : 0;
          if (!personId || isNaN(personId)) {
            setSchedulingState("err", "Select a person from the Personnel Roster.");
            return;
          }
          const person = personnelById(personId);
          if (!person) {
            setSchedulingState("err", "Person not found in roster. Refresh the list and try again.");
            return;
          }

          const sampleRow = hubSession.appointmentsSampleRow;
          const payload = {};
          let missingRequired = "";
          normalizedAppointmentsFormFields().forEach(function (field) {
            const el = document.getElementById("sf_" + field.key);
            if (!el) return;
            const writeKey = resolveAppointmentsWriteKey(field, sampleRow);
            el.dataset.writeKey = writeKey;
            const v = formFieldPayloadValue(el, writeKey);
            if (field.required && (v === null || v === "" || v === false)) {
              missingRequired = field.label;
              return;
            }
            if (field.inputType === "checkbox") {
              payload[writeKey] = !!v;
            } else if (v !== null && v !== "") {
              payload[writeKey] = v;
            }
          });
          if (missingRequired) {
            setSchedulingState("err", missingRequired + " is required.");
            return;
          }

          const personKey = await resolveAppointmentsPersonPostKey(seg, pw);
          payload[personKey] = personId;
          const createdKeys = ["CreatedAt", "Created_x0020_At", "Created"];
          if (sampleRow) {
            const createdHit = createdKeys.find(function (k) {
              return Object.prototype.hasOwnProperty.call(sampleRow, k);
            });
            if (createdHit) payload[createdHit] = new Date().toISOString();
          }

          try {
            setSchedulingState("loading", "Saving appointment...");
            if (schedulingSaveBtn) schedulingSaveBtn.disabled = true;
            await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            setSchedulingAddPanelVisible(false);
            await loadSchedulingAppointmentsList();
            setSchedulingState("ok", "Appointment saved for " + formatPersonDisplayName(person) + ".");
            window.setTimeout(function () {
              setSchedulingState("", "");
            }, 2800);
          } catch (e) {
            setSchedulingState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
            log("Appointment add failed:\n" + (e.message || String(e)), "err");
          } finally {
            if (schedulingSaveBtn) schedulingSaveBtn.disabled = false;
          }
        }

        function formatAfShortDate(d) {
          if (!d || isNaN(d.getTime())) return "";
          const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
          return (
            String(d.getDate()).padStart(2, "0") +
            "-" +
            months[d.getMonth()] +
            "-" +
            String(d.getFullYear()).slice(-2)
          );
        }

        function formatAfTime(d) {
          if (!d || isNaN(d.getTime())) return "";
          return String(d.getHours()) + ":" + String(d.getMinutes()).padStart(2, "0");
        }

        function buildAfOfficialDocHeader(title, options) {
          options = options || {};
          const header = document.createElement("div");
          header.className = "af-official-doc-header";
          const h1 = document.createElement("h1");
          h1.textContent = title || "Report";
          header.appendChild(h1);
          const now = options.generatedAt instanceof Date && !isNaN(options.generatedAt.getTime()) ? options.generatedAt : new Date();
          const dateLine = document.createElement("p");
          dateLine.className = "af-official-doc-dateline";
          dateLine.textContent = now.toLocaleDateString("en-US", {
            weekday: "long",
            month: "long",
            day: "numeric",
            year: "numeric",
          });
          header.appendChild(dateLine);
          const timeLine = document.createElement("p");
          timeLine.className = "af-official-doc-timeline";
          timeLine.textContent = now.toLocaleTimeString("en-US", {
            hour: "numeric",
            minute: "2-digit",
            second: options.includeSeconds ? "2-digit" : undefined,
            hour12: true,
          });
          header.appendChild(timeLine);
          if (options.subtitle) {
            const sub = document.createElement("p");
            sub.className = "af-official-doc-subtitle";
            sub.textContent = options.subtitle;
            header.appendChild(sub);
          }
          if (options.sortLine) {
            const sort = document.createElement("p");
            sort.className = "af-official-doc-sortline";
            sort.textContent = options.sortLine;
            header.appendChild(sort);
          }
          return header;
        }

        function groupAppointmentsByOffice(rows) {
          const groups = new Map();
          (Array.isArray(rows) ? rows : []).forEach(function (item) {
            const view = schedulingRowViewModel(item);
            const office = view.office || "Unassigned";
            if (!groups.has(office)) groups.set(office, []);
            groups.get(office).push(view);
          });
          const out = [];
          groups.forEach(function (items, office) {
            items.sort(function (a, b) {
              const ta = a.whenSort ? a.whenSort.getTime() : 0;
              const tb = b.whenSort ? b.whenSort.getTime() : 0;
              if (ta !== tb) return ta - tb;
              return String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" });
            });
            out.push({ office: office, items: items });
          });
          out.sort(function (a, b) {
            return String(a.office).localeCompare(String(b.office), undefined, { sensitivity: "base" });
          });
          return out;
        }

        function renderSchedulingUnitPrintReport(officeGroups, reportTitle) {
          const wrap = document.createElement("div");
          wrap.className = "scheduling-unit-report af-official-doc";

          wrap.appendChild(
            buildAfOfficialDocHeader(reportTitle || "Appointment Roster Report", { includeSeconds: true }),
          );

          if (!officeGroups.length) {
            const empty = document.createElement("p");
            empty.className = "scheduling-unit-empty";
            empty.textContent = "No appointments on file for this unit.";
            wrap.appendChild(empty);
            return wrap;
          }

          const columns = ["Name", "Rank", "Date", "Time", "Appointment", "Remarks", "Date Entered"];
          const sections = officeGroups.map(function (group) {
            return { section: group.office, items: group.items };
          });
          renderOfficialSectionedTables(wrap, {
            columns: columns,
            sections: sections,
            compact: true,
            buildCells: function (view) {
              const whenDate = view.whenSort;
              const createdDate = parseAppointmentDateTime(valueFromItemByKeys(view.item, appointmentCreatedAtKeys()));
              const appointmentText = [view.description, view.location].filter(Boolean).join("; ");
              const remarks = view.instructor || "";
              return [
                view.name,
                view.rank,
                formatAfShortDate(whenDate) || "-",
                formatAfTime(whenDate) || "-",
                appointmentText || "-",
                remarks || "-",
                formatAfShortDate(createdDate) || displayCellText(view.createdAt) || "-",
              ];
            },
          });

          return wrap;
        }

        async function printSchedulingUnitReport(missedOnly) {
          const pw = hubSession.pw;
          const seg = appointmentsListApiPath();
          if (!pw) {
            setSchedulingState("warn", "Load the Personnel Roster first.");
            return;
          }
          try {
            setSchedulingState("loading", missedOnly ? "Preparing missed report..." : "Preparing unit report...");
            let rows = schedulingSession.rows;
            if (!rows) {
              rows = await fetchAllAppointmentsRows(seg, pw);
              const archiveResult = await archivePastAppointmentItems(seg, pw, rows);
              rows = archiveResult.remaining;
            }
            schedulingSession.rows = rows.slice();
            if (missedOnly) {
              rows = rows.filter(function (item) {
                return appointmentIsUpcoming(item) && appointmentIsMissed(item);
              });
            } else {
              rows = rows.filter(appointmentIsUpcoming);
            }
            const officeGroups = groupAppointmentsByOffice(rows);
            setSchedulingState("", "");
            triggerSchedulingPrint(
              renderSchedulingUnitPrintReport(
                officeGroups,
                missedOnly ? "Missed Appointments Report" : "Appointment Roster Report",
              ),
            );
          } catch (e) {
            setSchedulingState("err", "Could not prepare report: " + (e.message || String(e)).slice(0, 220));
          }
        }


        function setHubListViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = !visible;
          });
          if (personDetailSection) personDetailSection.hidden = visible;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = true;
          if (phase1Section) phase1Section.hidden = true;
          hideEthosSections();
        }

        function setInstructorsViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = !visible;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = true;
          if (phase1Section) phase1Section.hidden = true;
          hideEthosSections();
        }

        function setReportsViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = !visible;
          if (schedulingSection) schedulingSection.hidden = true;
          if (phase1Section) phase1Section.hidden = true;
          hideEthosSections();
        }

        function setSchedulingViewVisible(visible) {
          document.querySelectorAll(".hub-section--form, .hub-section--status, .hub-section--roster").forEach(function (el) {
            el.hidden = visible;
          });
          if (personDetailSection) personDetailSection.hidden = true;
          if (instructorsSection) instructorsSection.hidden = true;
          if (reportsSection) reportsSection.hidden = true;
          if (schedulingSection) schedulingSection.hidden = !visible;
          if (phase1Section) phase1Section.hidden = true;
          hideEthosSections();
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
          if (reportsExportBtn) reportsExportBtn.hidden = true;
          renderReportsHubGrid();
        }

        function updateReportsPrintButtons(reportDef) {
          if (!reportDef) return;
          const id = reportDef.id;
          const signedIds = ["mql-report", "heavy-weapons", "cleo-report", "post-rc-report"];
          const showSigned = reportDef.printable && signedIds.indexOf(id) >= 0;
          const showExport = reportDef.printable && id === "mql-pdf";
          if (reportsPrintBtn) {
            if (id === "status-of-training") {
              updateSotPrintButtonLabel();
            } else {
              reportsPrintBtn.textContent =
                id === "heavy-weapons" ? "Print heavy weapons report" : "Print signed report";
              reportsPrintBtn.hidden = !showSigned;
            }
          }
          if (reportsExportBtn) {
            reportsExportBtn.textContent = "Generate PDF";
            reportsExportBtn.hidden = !showExport;
          }
        }

        function showReportsDetail(reportDef) {
          if (!reportDef) return;
          reportsSession.activeReportId = reportDef.id;
          if (reportsHubPanel) reportsHubPanel.hidden = true;
          if (reportsDetailPanel) reportsDetailPanel.hidden = false;
          if (reportsDetailTitle) reportsDetailTitle.textContent = reportDef.title || "Report";
          if (reportsDetailSubtitle) reportsDetailSubtitle.textContent = reportDef.subtitle || "";
          updateReportsPrintButtons(reportDef);
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

        function personOfficeKey(item) {
          const office = itemFieldText(item, "OfficeSymbol");
          return office || "Unassigned";
        }

        function personDutySectionLabel(item) {
          return itemFieldText(item, "OfficeSymbol") || itemFieldText(item, "DutySection") || "-";
        }

        function personStatusLabel(item) {
          const status = itemFieldText(item, "Status");
          return status || "Unknown";
        }

        function personCountsForSotPosture(item) {
          const allowed = Array.isArray(SOT_PERSONNEL_STATUS_FOR_COUNTS) ? SOT_PERSONNEL_STATUS_FOR_COUNTS : [];
          if (!allowed.length) return true;
          const status = personStatusLabel(item);
          return allowed.some(function (entry) {
            return String(entry).trim().toLowerCase() === status.toLowerCase();
          });
        }

        function weaponsQualDateKeys() {
          const col = normalizedWeaponsCertColumns().find(function (c) {
            return c.key === "QualDate";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["QualDate"];
        }

        function bylawQualDateKeys() {
          const col = normalizedBylawTrainingColumns().find(function (c) {
            return c.key === "QualDate";
          });
          return col && col.tryKeys ? col.tryKeys.slice() : ["QualDate"];
        }

        function yearMonthKeyFromDate(d) {
          if (!d || isNaN(d.getTime())) return "";
          return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
        }

        function currentYearMonthKey() {
          return yearMonthKeyFromDate(new Date());
        }

        function formatYearMonthLabel(yyyyMm) {
          const parts = String(yyyyMm || "").split("-");
          if (parts.length !== 2) return String(yyyyMm || "");
          const y = parts[0];
          const m = parseInt(parts[1], 10);
          if (!y || !m || m < 1 || m > 12) return String(yyyyMm || "");
          const mon = WEAPONS_CERT_MONTH_ABBR[m - 1] || "";
          return mon + " " + y;
        }

        function buildSotMonthOptions() {
          const out = [];
          const now = new Date();
          const count = Math.max(1, parseInt(String(SOT_REPORT_MONTHS_BACK || 24), 10) || 24);
          for (let i = 0; i < count; i++) {
            const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const key = yearMonthKeyFromDate(d);
            if (!key) continue;
            out.push({ key: key, label: formatYearMonthLabel(key) });
          }
          return out;
        }

        function trainingRowQualInMonth(row, yyyyMm, qualKeys) {
          const raw = valueFromItemByKeys(row, qualKeys);
          const d = parseWeaponsCertCalendarDate(raw);
          return yearMonthKeyFromDate(d) === yyyyMm;
        }

        function weaponLabelFromRow(row) {
          const raw = valueFromItemByKeys(row, ["Weapon", "WeaponName", "Weapon_x0020_Name", "Title"]);
          return formatCellValue(raw) || "Weapon";
        }

        function bylawLabelFromRow(row) {
          const raw = valueFromItemByKeys(row, BYLAW_TRAINING_ITEM_SORT_KEYS);
          return formatCellValue(raw) || "Training";
        }

        function tallyActivityLabels(rows, labelFn) {
          const counts = {};
          (Array.isArray(rows) ? rows : []).forEach(function (row) {
            const label = String(labelFn(row) || "Unknown").trim() || "Unknown";
            counts[label] = (counts[label] || 0) + 1;
          });
          return Object.keys(counts)
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .map(function (label) {
              return { label: label, count: counts[label] };
            });
        }

        function incrementToneBucket(bucket, tone) {
          const key = tone && bucket[tone] != null ? tone : "unknown";
          bucket[key] = (bucket[key] || 0) + 1;
        }

        function postureBucketToLines(bucket) {
          const order = [
            { key: "ok", label: "Qualified" },
            { key: "warn", label: "Due 31-60d" },
            { key: "urgent", label: "Due <=30d" },
            { key: "expired", label: "Expired" },
            { key: "unknown", label: "No records" },
          ];
          const lines = [];
          order.forEach(function (entry) {
            const n = bucket[entry.key] || 0;
            if (n > 0) lines.push(entry.label + ": " + n);
          });
          return lines.length ? lines.join(" | ") : "No records";
        }

        function statusCountsToText(statusCounts) {
          return Object.keys(statusCounts)
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .map(function (status) {
              return status + ": " + statusCounts[status];
            })
            .join(" | ");
        }

        function sotStatusCategoryRules() {
          return [
            { label: "Deployed / TDY / Leave", match: /deploy|tdy|leave|pcs|tour/i },
            { label: "TAW Personnel", match: /\btaw\b/i },
            { label: "IMA", match: /\bima\b/i },
            { label: "Profile", match: /profile/i },
          ];
        }

        function sotStatusCategoryLabel(status) {
          const s = String(status || "").trim();
          if (!s) return "Unknown";
          const rules = sotStatusCategoryRules();
          for (let i = 0; i < rules.length; i++) {
            if (rules[i].match.test(s)) return rules[i].label;
          }
          return s;
        }

        function buildStatusCategoryCounts(statusCounts) {
          const out = {};
          Object.keys(statusCounts || {}).forEach(function (status) {
            const label = sotStatusCategoryLabel(status);
            out[label] = (out[label] || 0) + statusCounts[status];
          });
          return out;
        }

        function renderSotDataTable(columns, rows, emptyText) {
          const wrap = document.createElement("div");
          wrap.className = "roster-wrap";
          if (!rows || !rows.length) {
            const p = document.createElement("p");
            p.className = "reports-office-empty";
            p.textContent = emptyText || "No data.";
            wrap.appendChild(p);
            return wrap;
          }
          const table = document.createElement("table");
          table.className = "reports-sot-table";
          const thead = document.createElement("thead");
          const trHead = document.createElement("tr");
          columns.forEach(function (col) {
            const th = document.createElement("th");
            th.textContent = col.label;
            trHead.appendChild(th);
          });
          thead.appendChild(trHead);
          table.appendChild(thead);
          const tbody = document.createElement("tbody");
          rows.forEach(function (row) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              if (col.numeric) td.className = "reports-sot-num";
              const val = row[col.key];
              td.textContent = val === null || val === undefined || val === "" ? "-" : String(val);
              tr.appendChild(td);
            });
            tbody.appendChild(tr);
          });
          table.appendChild(tbody);
          wrap.appendChild(table);
          return wrap;
        }

        function buildOfficeWeaponTypeStats(personRows, weaponsRows, officeKey) {
          const people = (Array.isArray(personRows) ? personRows : []).filter(function (person) {
            return person && personOfficeKey(person) === officeKey;
          });
          const posturePeople = people.filter(personCountsForSotPosture);
          const required = posturePeople.length;
          const weaponStats = {};

          posturePeople.forEach(function (person) {
            const pid = String(person.Id);
            const rows = weaponsRows.filter(function (row) {
              return weaponsPersonnelIdFromItem(row) === pid;
            });
            const byWeapon = {};
            rows.forEach(function (row) {
              const label = weaponLabelFromRow(row);
              const tone = computeWeaponsCertStatus(row).tone;
              if (!byWeapon[label] || trainingStatusToneRank(tone) > trainingStatusToneRank(byWeapon[label])) {
                byWeapon[label] = tone;
              }
            });
            Object.keys(byWeapon).forEach(function (label) {
              if (!weaponStats[label]) {
                weaponStats[label] = { qualified: 0, dueSoon: 0, overdue: 0 };
              }
              const tone = byWeapon[label];
              if (tone === "ok") weaponStats[label].qualified += 1;
              else if (tone === "warn") weaponStats[label].dueSoon += 1;
              else if (tone === "urgent" || tone === "expired") weaponStats[label].overdue += 1;
            });
          });

          return Object.keys(weaponStats)
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .map(function (label) {
              const s = weaponStats[label];
              return {
                weapon: label,
                required: required,
                qualified: s.qualified,
                dueSoon: s.dueSoon,
                overdue: s.overdue,
              };
            });
        }

        function buildOfficeBylawItemStats(personRows, bylawRows, officeKey, yyyyMm) {
          const people = (Array.isArray(personRows) ? personRows : []).filter(function (person) {
            return person && personOfficeKey(person) === officeKey;
          });
          const posturePeople = people.filter(personCountsForSotPosture);
          const required = posturePeople.length;
          const bylawQualKeys = bylawQualDateKeys();
          const itemStats = {};

          posturePeople.forEach(function (person) {
            const pid = String(person.Id);
            const rows = bylawRows.filter(function (row) {
              return bylawPersonnelIdFromItem(row) === pid;
            });
            const byItem = {};
            rows.forEach(function (row) {
              const label = bylawLabelFromRow(row);
              const tone = computeBylawTrainingStatus(row).tone;
              if (!byItem[label]) {
                byItem[label] = { tone: tone, completedMonth: 0 };
              }
              if (trainingStatusToneRank(tone) > trainingStatusToneRank(byItem[label].tone)) {
                byItem[label].tone = tone;
              }
              if (trainingRowQualInMonth(row, yyyyMm, bylawQualKeys)) {
                byItem[label].completedMonth += 1;
              }
            });
            Object.keys(byItem).forEach(function (label) {
              if (!itemStats[label]) {
                itemStats[label] = { qualified: 0, dueSoon: 0, overdue: 0, completedMonth: 0 };
              }
              const tone = byItem[label].tone;
              if (tone === "ok") itemStats[label].qualified += 1;
              else if (tone === "warn") itemStats[label].dueSoon += 1;
              else if (tone === "urgent" || tone === "expired") itemStats[label].overdue += 1;
              itemStats[label].completedMonth += byItem[label].completedMonth;
            });
          });

          return Object.keys(itemStats)
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .map(function (label) {
              const s = itemStats[label];
              return {
                item: label,
                required: required,
                qualified: s.qualified,
                dueSoon: s.dueSoon,
                overdue: s.overdue,
                completedMonth: s.completedMonth,
                monthRange: formatYearMonthLabel(yyyyMm),
              };
            });
        }

        function buildSquadronMonthlyRollup(officeGroups) {
          const rollup = {
            headcount: 0,
            postureIncluded: 0,
            weaponsPosture: { ok: 0, warn: 0, urgent: 0, expired: 0, unknown: 0 },
            bylawPosture: { ok: 0, warn: 0, urgent: 0, expired: 0, unknown: 0 },
            statusCounts: {},
            monthWeaponsTally: [],
            monthBylawTally: [],
            totalWeaponsOverdue: 0,
            totalBylawOverdue: 0,
          };
          const weaponsMonth = {};
          const bylawMonth = {};

          (Array.isArray(officeGroups) ? officeGroups : []).forEach(function (group) {
            rollup.headcount += group.headcount || 0;
            rollup.postureIncluded += group.postureIncluded || 0;
            ["ok", "warn", "urgent", "expired", "unknown"].forEach(function (tone) {
              rollup.weaponsPosture[tone] += (group.weaponsPosture && group.weaponsPosture[tone]) || 0;
              rollup.bylawPosture[tone] += (group.bylawPosture && group.bylawPosture[tone]) || 0;
            });
            Object.keys(group.statusCounts || {}).forEach(function (status) {
              rollup.statusCounts[status] = (rollup.statusCounts[status] || 0) + group.statusCounts[status];
            });
            (group.monthWeaponsTally || []).forEach(function (entry) {
              weaponsMonth[entry.label] = (weaponsMonth[entry.label] || 0) + entry.count;
            });
            (group.monthBylawTally || []).forEach(function (entry) {
              bylawMonth[entry.label] = (bylawMonth[entry.label] || 0) + entry.count;
            });
            rollup.totalWeaponsOverdue +=
              ((group.weaponsPosture && group.weaponsPosture.urgent) || 0) +
              ((group.weaponsPosture && group.weaponsPosture.expired) || 0);
            rollup.totalBylawOverdue +=
              ((group.bylawPosture && group.bylawPosture.urgent) || 0) +
              ((group.bylawPosture && group.bylawPosture.expired) || 0);
          });

          rollup.monthWeaponsTally = Object.keys(weaponsMonth)
            .sort()
            .map(function (label) {
              return { label: label, count: weaponsMonth[label] };
            });
          rollup.monthBylawTally = Object.keys(bylawMonth)
            .sort()
            .map(function (label) {
              return { label: label, count: bylawMonth[label] };
            });
          return rollup;
        }

        function buildMonthlyOfficeGroups(personRows, weaponsRows, bylawRows, yyyyMm) {
          const offices = {};
          const weaponsQualKeys = weaponsQualDateKeys();
          const bylawQualKeys = bylawQualDateKeys();

          function ensureOffice(key) {
            if (!offices[key]) {
              offices[key] = {
                office: key,
                statusCounts: {},
                headcount: 0,
                postureIncluded: 0,
                weaponsPosture: { ok: 0, warn: 0, urgent: 0, expired: 0, unknown: 0 },
                bylawPosture: { ok: 0, warn: 0, urgent: 0, expired: 0, unknown: 0 },
                monthWeapons: [],
                monthBylaw: [],
              };
            }
            return offices[key];
          }

          (Array.isArray(personRows) ? personRows : []).forEach(function (person) {
            if (!person || person.Id == null) return;
            const officeKey = personOfficeKey(person);
            const group = ensureOffice(officeKey);
            const status = personStatusLabel(person);
            group.statusCounts[status] = (group.statusCounts[status] || 0) + 1;
            group.headcount += 1;

            if (!personCountsForSotPosture(person)) return;
            group.postureIncluded += 1;

            const pid = String(person.Id);
            const personWeapons = weaponsRows.filter(function (row) {
              return weaponsPersonnelIdFromItem(row) === pid;
            });
            const personBylaw = bylawRows.filter(function (row) {
              return bylawPersonnelIdFromItem(row) === pid;
            });

            const weaponsStatus = summarizeTrainingRows(personWeapons, computeWeaponsCertStatus);
            const bylawStatus = summarizeTrainingRows(personBylaw, computeBylawTrainingStatus);
            incrementToneBucket(group.weaponsPosture, weaponsStatus.tone);
            incrementToneBucket(group.bylawPosture, bylawStatus.tone);

            personWeapons.forEach(function (row) {
              if (trainingRowQualInMonth(row, yyyyMm, weaponsQualKeys)) group.monthWeapons.push(row);
            });
            personBylaw.forEach(function (row) {
              if (trainingRowQualInMonth(row, yyyyMm, bylawQualKeys)) group.monthBylaw.push(row);
            });
          });

          return Object.keys(offices)
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .map(function (key) {
              const group = offices[key];
              group.monthWeaponsTally = tallyActivityLabels(group.monthWeapons, weaponLabelFromRow);
              group.monthBylawTally = tallyActivityLabels(group.monthBylaw, bylawLabelFromRow);
              group.statusCategoryCounts = buildStatusCategoryCounts(group.statusCounts);
              group.weaponTypeStats = buildOfficeWeaponTypeStats(personRows, weaponsRows, key);
              group.bylawItemStats = buildOfficeBylawItemStats(personRows, bylawRows, key, yyyyMm);
              return group;
            });
        }

        function renderActivityList(items, emptyText) {
          if (!items || !items.length) {
            const p = document.createElement("p");
            p.className = "reports-office-empty";
            p.textContent = emptyText;
            return p;
          }
          const ul = document.createElement("ul");
          ul.className = "reports-office-list";
          items.forEach(function (entry) {
            const li = document.createElement("li");
            li.textContent = entry.label + " (" + entry.count + ")";
            ul.appendChild(li);
          });
          return ul;
        }

        function renderMonthlyOfficeBlock(group, yyyyMm) {
          const block = document.createElement("section");
          block.className = "reports-office-block";

          const title = document.createElement("h3");
          title.className = "reports-office-title";
          title.textContent = group.office;
          block.appendChild(title);

          const meta = document.createElement("p");
          meta.className = "reports-office-meta";
          meta.textContent =
            "Assigned Personnel: " +
            group.headcount +
            (group.postureIncluded !== group.headcount
              ? " | Training posture counts: " + group.postureIncluded + " (Status filter)"
              : "");
          block.appendChild(meta);

          const statusSection = document.createElement("div");
          statusSection.className = "reports-sot-section-block";
          const statusTitle = document.createElement("h4");
          statusTitle.className = "reports-office-subtitle";
          statusTitle.textContent = "Personnel status";
          statusSection.appendChild(statusTitle);
          const statusRows = Object.keys(group.statusCategoryCounts || {})
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .map(function (label) {
              return { status: label, count: group.statusCategoryCounts[label] };
            });
          statusSection.appendChild(
            renderSotDataTable(
              [
                { key: "status", label: "Status" },
                { key: "count", label: "Count", numeric: true },
              ],
              statusRows,
              "No Personnel status data.",
            ),
          );
          block.appendChild(statusSection);

          const weaponsSection = document.createElement("div");
          weaponsSection.className = "reports-sot-section-block";
          const weaponsTitle = document.createElement("h4");
          weaponsTitle.className = "reports-office-subtitle";
          weaponsTitle.textContent = "Monthly required weapons qualifications";
          weaponsSection.appendChild(weaponsTitle);
          weaponsSection.appendChild(
            renderSotDataTable(
              [
                { key: "weapon", label: "Weapon" },
                { key: "required", label: "Required", numeric: true },
                { key: "qualified", label: "Qualified", numeric: true },
                { key: "dueSoon", label: "Due 31-60d", numeric: true },
                { key: "overdue", label: "Overdue", numeric: true },
              ],
              group.weaponTypeStats || [],
              "No weapons qualification records for this office.",
            ),
          );
          block.appendChild(weaponsSection);

          const bylawSection = document.createElement("div");
          bylawSection.className = "reports-sot-section-block";
          const bylawTitle = document.createElement("h4");
          bylawTitle.className = "reports-office-subtitle";
          bylawTitle.textContent = "Scheduled annual AUoF / By-Law training";
          bylawSection.appendChild(bylawTitle);
          bylawSection.appendChild(
            renderSotDataTable(
              [
                { key: "item", label: "Training" },
                { key: "monthRange", label: "Training dates" },
                { key: "required", label: "Required", numeric: true },
                { key: "completedMonth", label: "Completed trng", numeric: true },
                { key: "qualified", label: "Qualified", numeric: true },
                { key: "overdue", label: "Overdue", numeric: true },
              ],
              group.bylawItemStats || [],
              "No By-Law training records for this office.",
            ),
          );
          block.appendChild(bylawSection);

          const monthSection = document.createElement("div");
          monthSection.className = "reports-sot-section-block";
          const monthTitle = document.createElement("h4");
          monthTitle.className = "reports-office-subtitle";
          monthTitle.textContent = "Completed in " + formatYearMonthLabel(yyyyMm);
          monthSection.appendChild(monthTitle);

          const weaponsMonth = document.createElement("h4");
          weaponsMonth.className = "reports-office-subtitle";
          weaponsMonth.textContent = "Weapons qualifications";
          monthSection.appendChild(weaponsMonth);
          monthSection.appendChild(
            renderActivityList(group.monthWeaponsTally, "No weapons qualifications recorded this month."),
          );

          const bylawMonth = document.createElement("h4");
          bylawMonth.className = "reports-office-subtitle";
          bylawMonth.textContent = "By-Law training";
          monthSection.appendChild(bylawMonth);
          monthSection.appendChild(
            renderActivityList(group.monthBylawTally, "No By-Law training recorded this month."),
          );
          block.appendChild(monthSection);

          return block;
        }

        function renderSquadronMonthlySummary(rollup, yyyyMm) {
          const box = document.createElement("div");
          box.className = "reports-sot-squadron-summary";

          const title = document.createElement("h4");
          title.className = "reports-office-subtitle";
          title.textContent = "Squadron totals";
          box.appendChild(title);

          const meta = document.createElement("p");
          meta.className = "reports-office-meta";
          meta.textContent =
            "Personnel: " +
            rollup.headcount +
            " | Training posture: " +
            rollup.postureIncluded +
            " | Weapons overdue: " +
            rollup.totalWeaponsOverdue +
            " | By-Law overdue: " +
            rollup.totalBylawOverdue;
          box.appendChild(meta);

          const postureSection = document.createElement("div");
          postureSection.className = "reports-sot-section-block";
          const postureTitle = document.createElement("h4");
          postureTitle.className = "reports-office-subtitle";
          postureTitle.textContent = "Training posture (all offices)";
          postureSection.appendChild(postureTitle);
          postureSection.appendChild(
            renderSotDataTable(
              [
                { key: "area", label: "Area" },
                { key: "qualified", label: "Qualified", numeric: true },
                { key: "dueSoon", label: "Due 31-60d", numeric: true },
                { key: "urgent", label: "Due <=30d", numeric: true },
                { key: "expired", label: "Expired", numeric: true },
                { key: "unknown", label: "No records", numeric: true },
              ],
              [
                {
                  area: "Weapons",
                  qualified: rollup.weaponsPosture.ok || 0,
                  dueSoon: rollup.weaponsPosture.warn || 0,
                  urgent: rollup.weaponsPosture.urgent || 0,
                  expired: rollup.weaponsPosture.expired || 0,
                  unknown: rollup.weaponsPosture.unknown || 0,
                },
                {
                  area: "By-Law",
                  qualified: rollup.bylawPosture.ok || 0,
                  dueSoon: rollup.bylawPosture.warn || 0,
                  urgent: rollup.bylawPosture.urgent || 0,
                  expired: rollup.bylawPosture.expired || 0,
                  unknown: rollup.bylawPosture.unknown || 0,
                },
              ],
              "",
            ),
          );
          box.appendChild(postureSection);

          const monthSection = document.createElement("div");
          monthSection.className = "reports-sot-section-block";
          const monthTitle = document.createElement("h4");
          monthTitle.className = "reports-office-subtitle";
          monthTitle.textContent = "Total completed in " + formatYearMonthLabel(yyyyMm);
          monthSection.appendChild(monthTitle);
          monthSection.appendChild(
            renderActivityList(rollup.monthWeaponsTally, "No squadron weapons qualifications this month."),
          );
          monthSection.appendChild(
            renderActivityList(rollup.monthBylawTally, "No squadron By-Law training this month."),
          );
          box.appendChild(monthSection);

          return box;
        }

        function renderMonthlySotReport(officeGroups, yyyyMm) {
          const wrap = document.createElement("div");
          wrap.className = "reports-sot-monthly-print";

          const header = document.createElement("div");
          header.className = "reports-month-print-header";
          const docTitle = document.createElement("h3");
          docTitle.className = "reports-sot-doc-title";
          docTitle.textContent = "STATUS OF TRAINING";
          const squadron = document.createElement("p");
          squadron.textContent = "SQUADRON: " + String(SOT_SQUADRON_LABEL || "88 SFS");
          const dataMonth = document.createElement("p");
          dataMonth.textContent = "DATA MONTH: " + formatYearMonthLabel(yyyyMm);
          const generated = document.createElement("p");
          const today = new Date();
          generated.textContent =
            "Report generated: " +
            formatWeaponsCertDisplayDate(isoDateFromCalendarDate(today)) +
            " | Office summary (no names)";
          header.appendChild(docTitle);
          header.appendChild(squadron);
          header.appendChild(dataMonth);
          header.appendChild(generated);
          wrap.appendChild(header);

          if (!officeGroups.length) {
            const empty = document.createElement("p");
            empty.className = "reports-office-empty";
            empty.textContent = "No office groups to display.";
            wrap.appendChild(empty);
            return wrap;
          }

          const rollup = buildSquadronMonthlyRollup(officeGroups);
          wrap.appendChild(renderSquadronMonthlySummary(rollup, yyyyMm));

          const officesTitle = document.createElement("h4");
          officesTitle.className = "reports-office-subtitle";
          officesTitle.textContent = "By office";
          wrap.appendChild(officesTitle);

          officeGroups.forEach(function (group) {
            wrap.appendChild(renderMonthlyOfficeBlock(group, yyyyMm));
          });

          const sig = document.createElement("p");
          sig.className = "reports-office-meta";
          sig.style.marginTop = "18px";
          sig.textContent = "Training NCOIC Sig: _________________________    CC Sig: _________________________";
          wrap.appendChild(sig);

          return wrap;
        }

        function updateSotPrintButtonLabel() {
          if (!reportsPrintBtn) return;
          if (reportsSession.activeReportId !== "status-of-training") return;
          if (reportsSession.sotView === "monthly") {
            reportsPrintBtn.textContent = "Print monthly report";
            reportsPrintBtn.hidden = false;
          } else {
            reportsPrintBtn.textContent = "Print report";
            reportsPrintBtn.hidden = true;
          }
          if (reportsExportBtn) reportsExportBtn.hidden = true;
        }

        function setSotView(view) {
          reportsSession.sotView = view === "monthly" ? "monthly" : "active";
          const activeView = document.getElementById("reportsSotActiveView");
          const monthlyView = document.getElementById("reportsSotMonthlyView");
          const tabActive = document.getElementById("reportsSotTabActive");
          const tabMonthly = document.getElementById("reportsSotTabMonthly");
          if (activeView) activeView.hidden = reportsSession.sotView !== "active";
          if (monthlyView) monthlyView.hidden = reportsSession.sotView !== "monthly";
          if (tabActive) tabActive.classList.toggle("reports-sot-tab--active", reportsSession.sotView === "active");
          if (tabMonthly) tabMonthly.classList.toggle("reports-sot-tab--active", reportsSession.sotView === "monthly");
          updateSotPrintButtonLabel();
        }

        function refreshMonthlySotView() {
          const monthSelect = document.getElementById("reportsSotMonthSelect");
          const monthlyBody = document.getElementById("reportsSotMonthlyBody");
          if (!monthlyBody) return;
          const yyyyMm = monthSelect ? String(monthSelect.value || "").trim() : "";
          reportsSession.selectedMonth = yyyyMm || currentYearMonthKey();
          if (monthSelect && !monthSelect.value) monthSelect.value = reportsSession.selectedMonth;

          const personRows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
          const weaponsRows = Array.isArray(reportsSession.weaponsRows) ? reportsSession.weaponsRows : [];
          const bylawRows = Array.isArray(reportsSession.bylawRows) ? reportsSession.bylawRows : [];
          const officeGroups = buildMonthlyOfficeGroups(personRows, weaponsRows, bylawRows, reportsSession.selectedMonth);
          monthlyBody.innerHTML = "";
          monthlyBody.appendChild(renderMonthlySotReport(officeGroups, reportsSession.selectedMonth));
        }

        function wireStatusOfTrainingControls() {
          const tabActive = document.getElementById("reportsSotTabActive");
          const tabMonthly = document.getElementById("reportsSotTabMonthly");
          const monthSelect = document.getElementById("reportsSotMonthSelect");
          const monthApply = document.getElementById("reportsSotMonthApply");

          if (tabActive && !tabActive.dataset.wired) {
            tabActive.dataset.wired = "1";
            tabActive.addEventListener("click", function () {
              setSotView("active");
            });
          }
          if (tabMonthly && !tabMonthly.dataset.wired) {
            tabMonthly.dataset.wired = "1";
            tabMonthly.addEventListener("click", function () {
              setSotView("monthly");
              refreshMonthlySotView();
            });
          }
          if (monthApply && !monthApply.dataset.wired) {
            monthApply.dataset.wired = "1";
            monthApply.addEventListener("click", function () {
              refreshMonthlySotView();
              setReportsState("ok", "Monthly report updated for " + formatYearMonthLabel(reportsSession.selectedMonth) + ".");
              window.setTimeout(function () {
                setReportsState("", "");
              }, 2200);
            });
          }
          if (monthSelect && !monthSelect.dataset.wired) {
            monthSelect.dataset.wired = "1";
            monthSelect.addEventListener("change", function () {
              refreshMonthlySotView();
            });
          }
        }

        function buildStatusOfTrainingShell(activeChildren, monthSelectOptions) {
          const frag = document.createDocumentFragment();

          const toolbar = document.createElement("div");
          toolbar.className = "reports-sot-toolbar reports-no-print";
          const tabActive = document.createElement("button");
          tabActive.type = "button";
          tabActive.id = "reportsSotTabActive";
          tabActive.className = "btn-secondary reports-sot-tab--active";
          tabActive.textContent = "Active snapshot";
          const tabMonthly = document.createElement("button");
          tabMonthly.type = "button";
          tabMonthly.id = "reportsSotTabMonthly";
          tabMonthly.className = "btn-secondary";
          tabMonthly.textContent = "Monthly report";
          toolbar.appendChild(tabActive);
          toolbar.appendChild(tabMonthly);
          frag.appendChild(toolbar);

          const activeView = document.createElement("div");
          activeView.id = "reportsSotActiveView";
          (Array.isArray(activeChildren) ? activeChildren : []).forEach(function (node) {
            if (node) activeView.appendChild(node);
          });
          frag.appendChild(activeView);

          const monthlyView = document.createElement("div");
          monthlyView.id = "reportsSotMonthlyView";
          monthlyView.hidden = true;

          const controls = document.createElement("div");
          controls.className = "reports-month-controls reports-no-print";
          const field = document.createElement("div");
          field.className = "reports-month-field";
          const lab = document.createElement("label");
          lab.setAttribute("for", "reportsSotMonthSelect");
          lab.textContent = "Report month";
          const sel = document.createElement("select");
          sel.id = "reportsSotMonthSelect";
          (Array.isArray(monthSelectOptions) ? monthSelectOptions : []).forEach(function (opt) {
            const o = document.createElement("option");
            o.value = opt.key;
            o.textContent = opt.label;
            sel.appendChild(o);
          });
          field.appendChild(lab);
          field.appendChild(sel);
          controls.appendChild(field);
          const applyBtn = document.createElement("button");
          applyBtn.type = "button";
          applyBtn.id = "reportsSotMonthApply";
          applyBtn.className = "btn-secondary";
          applyBtn.textContent = "Update month";
          controls.appendChild(applyBtn);
          monthlyView.appendChild(controls);

          const monthlyBody = document.createElement("div");
          monthlyBody.id = "reportsSotMonthlyBody";
          monthlyView.appendChild(monthlyBody);
          frag.appendChild(monthlyView);

          return frag;
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

        function groupReportEntriesBySection(entries, getPerson) {
          const groups = new Map();
          (Array.isArray(entries) ? entries : []).forEach(function (entry) {
            const person = getPerson(entry);
            const section = person ? personOfficeKey(person) : "Unassigned";
            if (!groups.has(section)) groups.set(section, []);
            groups.get(section).push(entry);
          });
          const out = [];
          groups.forEach(function (items, section) {
            items.sort(function (a, b) {
              return formatPersonDisplayName(getPerson(a)).localeCompare(formatPersonDisplayName(getPerson(b)), undefined, {
                sensitivity: "base",
              });
            });
            out.push({ section: section, items: items });
          });
          out.sort(function (a, b) {
            return String(a.section).localeCompare(String(b.section), undefined, { sensitivity: "base" });
          });
          return out;
        }

        function appendOfficialSectionTable(parent, options) {
          renderOfficialSectionedTables(parent, options);
        }

        function officialColumnWidths(columnLabels) {
          const n = columnLabels.length;
          if (!n) return [];
          if (n === 2) return ["44%", "56%"];
          if (n === 3) return ["40%", "35%", "25%"];
          if (n > 3) {
            const certCount = n - 3;
            let dutyPct = 10;
            let namePct = 20;
            let rankPct = 8;
            if (n > 16) {
              dutyPct = 7;
              namePct = 12;
              rankPct = 6;
            } else if (n > 12) {
              dutyPct = 8;
              namePct = 14;
              rankPct = 7;
            } else if (n > 8) {
              dutyPct = 9;
              namePct = 16;
              rankPct = 7;
            }
            const certTotal = 100 - dutyPct - namePct - rankPct;
            const certEach = certTotal / certCount;
            return [dutyPct + "%", namePct + "%", rankPct + "%"].concat(
              Array(certCount)
                .fill(null)
                .map(function () {
                  return certEach + "%";
                }),
            );
          }
          const each = 100 / n;
          return columnLabels.map(function () {
            return each + "%";
          });
        }

        function appendOfficialColgroup(table, widths) {
          const existing = table.querySelector("colgroup");
          if (existing) existing.remove();
          const colgroup = document.createElement("colgroup");
          (widths || []).forEach(function (width) {
            const col = document.createElement("col");
            if (width) col.style.width = width;
            colgroup.appendChild(col);
          });
          if (table.firstChild) table.insertBefore(colgroup, table.firstChild);
          else table.appendChild(colgroup);
        }

        function renderOfficialSectionedTables(parent, options) {
          options = options || {};
          const columnLabels = options.columns || [];
          const sections = options.sections || [];
          const manyCols = columnLabels.length > 14;
          const tableClass =
            (options.compact ? "af-official-table af-official-table--compact" : "af-official-table") +
            (manyCols ? " af-official-table--many-cols" : "");
          const buildCells = options.buildCells;
          const sectionLabelFn =
            typeof options.sectionLabelFn === "function"
              ? options.sectionLabelFn
              : function (group) {
                  return group.section;
                };
          if (!parent || !columnLabels.length || !sections.length || typeof buildCells !== "function") return;

          const widths = officialColumnWidths(columnLabels);
          const headerWrap = document.createElement("div");
          headerWrap.className = "af-official-table-header-wrap";
          const headerTable = document.createElement("table");
          headerTable.className = tableClass + " af-official-table--columns";
          headerTable.style.width = "100%";
          appendOfficialColgroup(headerTable, widths);
          const thead = document.createElement("thead");
          const headRow = document.createElement("tr");
          columnLabels.forEach(function (label) {
            const th = document.createElement("th");
            th.textContent = label;
            headRow.appendChild(th);
          });
          thead.appendChild(headRow);
          headerTable.appendChild(thead);
          headerWrap.appendChild(headerTable);
          parent.appendChild(headerWrap);

          sections.forEach(function (group) {
            const block = document.createElement("div");
            block.className = "af-official-section-block";
            const heading = document.createElement("h3");
            heading.className = "af-official-section-heading";
            heading.textContent = sectionLabelFn(group) || "Unassigned";
            block.appendChild(heading);
            const table = document.createElement("table");
            table.className = tableClass + " af-official-table--section-body";
            table.style.width = "100%";
            appendOfficialColgroup(table, widths);
            const tbody = document.createElement("tbody");
            group.items.forEach(function (entry) {
              const tr = document.createElement("tr");
              (buildCells(entry) || []).forEach(function (cell) {
                const td = document.createElement("td");
                const text = cell && typeof cell === "object" && cell.text != null ? cell.text : cell;
                td.textContent = text === "" || text == null ? "-" : String(text);
                const tone = cell && typeof cell === "object" ? cell.tone : null;
                applyMqlCertCellStyle(td, tone);
                tr.appendChild(td);
              });
              tbody.appendChild(tr);
            });
            table.appendChild(tbody);
            block.appendChild(table);
            parent.appendChild(block);
          });
        }

        function mqlCertPairText(row) {
          if (!row) return "-";
          const qualKeys = weaponsQualDateKeys();
          const expKeys = weaponsCertExpiryDateKeys();
          const qual = formatWeaponsCertDisplayDate(valueFromItemByKeys(row, qualKeys));
          const exp = formatWeaponsCertDisplayDate(valueFromItemByKeys(row, expKeys));
          if (!qual && !exp) return "-";
          if (qual && exp) return qual + " / " + exp;
          return qual || exp;
        }

        function mqlCertQualText(row) {
          if (!row) return "-";
          const qual = formatWeaponsCertDisplayDate(valueFromItemByKeys(row, weaponsQualDateKeys()));
          return qual || "-";
        }

        function mqlCertExpText(row) {
          if (!row) return "-";
          const exp = formatWeaponsCertDisplayDate(valueFromItemByKeys(row, weaponsCertExpiryDateKeys()));
          return exp || "-";
        }

        function mqlCertCell(row, statusFn) {
          if (!row) return { text: "-", tone: "expired" };
          const text = mqlCertPairText(row);
          const status = typeof statusFn === "function" ? statusFn(row) : { tone: "unknown" };
          let tone = status && status.tone ? status.tone : "unknown";
          if (!text || text === "-") tone = "expired";
          return { text: text, tone: tone };
        }

        function mqlBylawCertQualText(row) {
          if (!row) return "-";
          const qual = formatWeaponsCertDisplayDate(valueFromItemByKeys(row, bylawQualDateKeys()));
          return qual || "-";
        }

        function mqlBylawCertExpText(row) {
          if (!row) return "-";
          const exp = formatWeaponsCertDisplayDate(valueFromItemByKeys(row, bylawTrainingExpiryDateKeys()));
          if (exp) return exp;
          const qual = parseWeaponsCertCalendarDate(valueFromItemByKeys(row, bylawQualDateKeys()));
          if (qual) return formatWeaponsCertDisplayDate(expirationDateFromQualDate(qual));
          return "-";
        }

        function mqlExportCertCell(row, part, statusFn, isWeapon) {
          const useWeapon = isWeapon !== false;
          const text =
            part === "qual"
              ? useWeapon
                ? mqlCertQualText(row)
                : mqlBylawCertQualText(row)
              : useWeapon
                ? mqlCertExpText(row)
                : mqlBylawCertExpText(row);
          if (!row) return { text: text, tone: "expired" };
          const status = typeof statusFn === "function" ? statusFn(row) : { tone: "unknown" };
          let tone = status && status.tone ? status.tone : "unknown";
          if (!text || text === "-") tone = "expired";
          return { text: text, tone: tone };
        }

        function mqlIsExcludedCatalogLabel(label) {
          const norm = String(label || "").trim().toLowerCase();
          if (!norm) return true;
          return (Array.isArray(MQL_EXCLUDED_CATALOG_LABELS) ? MQL_EXCLUDED_CATALOG_LABELS : []).some(function (ex) {
            const x = String(ex || "").trim().toLowerCase();
            if (!x) return false;
            return norm === x || norm.indexOf(x) >= 0;
          });
        }

        function mqlIsRequiredWeapon(label) {
          return (Array.isArray(MQL_REQUIRED_WEAPON_LABELS) ? MQL_REQUIRED_WEAPON_LABELS : []).some(function (req) {
            return String(req).toLowerCase() === String(label || "").toLowerCase();
          });
        }

        function mqlMergeBylawCatalogLabels(bylawMap, labels) {
          (Array.isArray(labels) ? labels : []).forEach(function (raw) {
            const label = String(raw || "").trim();
            if (!label || mqlIsExcludedCatalogLabel(label)) return;
            const key = label.toLowerCase();
            if (!bylawMap.has(key)) bylawMap.set(key, label);
          });
        }

        function mqlNormalizeColumnKey(label) {
          const s = String(label || "").trim();
          const upper = s.toUpperCase();
          if (upper === "M4" || upper.indexOf("M4") >= 0) return "M4";
          if (upper === "M18" || upper.indexOf("M18") >= 0) return "M18";
          return s;
        }

        function mqlLookupCertRow(certMap, columnLabel) {
          if (!certMap || !columnLabel) return null;
          if (certMap[columnLabel]) return certMap[columnLabel];
          const colNorm = mqlNormalizeColumnKey(columnLabel).toLowerCase();
          const keys = Object.keys(certMap);
          for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (k.toLowerCase() === String(columnLabel).toLowerCase()) return certMap[k];
            if (mqlNormalizeColumnKey(k).toLowerCase() === colNorm) return certMap[k];
          }
          return null;
        }

        function mqlAugmentCertMapAliases(certMap) {
          if (!certMap || typeof certMap !== "object") return certMap;
          const aliases = {};
          Object.keys(certMap).forEach(function (label) {
            const norm = mqlNormalizeColumnKey(label);
            if (norm && certMap[norm] == null) aliases[norm] = certMap[label];
          });
          Object.keys(aliases).forEach(function (key) {
            certMap[key] = aliases[key];
          });
          return certMap;
        }

        async function fetchMqlBylawItemChoices(pw, listApiPathFn, sampleRow, cacheHolder, cacheField, bylawRows) {
          const field = cacheField || "bylawItemChoices";
          if (cacheHolder && Array.isArray(cacheHolder[field]) && cacheHolder[field].length) {
            return cacheHolder[field];
          }
          const seg = typeof listApiPathFn === "function" ? listApiPathFn() : "";
          let options = [];
          if (seg && pw) {
            const itemCol = normalizedBylawTrainingColumns().find(function (col) {
              return col && col.key === "Item";
            });
            if (itemCol) {
              try {
                const r = await fetchChoiceOptionsForColumn(itemCol, seg, pw, sampleRow);
                options = (r && Array.isArray(r.choices) ? r.choices : [])
                  .map(function (opt) {
                    return String(opt || "").trim();
                  })
                  .filter(Boolean);
              } catch (e) {
                log("MQL by-law choice fetch failed:\n" + (e.message || String(e)), "warn");
              }
            }
          }
          if (!options.length) {
            const seen = new Set();
            (Array.isArray(bylawRows) ? bylawRows : []).forEach(function (row) {
              const label = bylawLabelFromRow(row);
              const clean = String(label || "").trim();
              if (!clean || clean === "Training") return;
              const key = clean.toLowerCase();
              if (!seen.has(key)) {
                seen.add(key);
                options.push(clean);
              }
            });
            options.sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            });
          }
          if (cacheHolder) cacheHolder[field] = options;
          return options;
        }

        function buildMqlCertCatalog(weaponsRows, bylawRows, opts) {
          opts = opts || {};
          const mode = opts.mode === "heavy" ? "heavy" : "standard";
          const weaponMap = new Map();
          if (mode === "standard") {
            (Array.isArray(MQL_REQUIRED_WEAPON_LABELS) ? MQL_REQUIRED_WEAPON_LABELS : []).forEach(function (label) {
              weaponMap.set(String(label).toLowerCase(), label);
            });
          }
          (Array.isArray(weaponsRows) ? weaponsRows : []).forEach(function (row) {
            const label = weaponLabelFromRow(row);
            const norm = mqlNormalizeColumnKey(label);
            const key = String(norm || label).toLowerCase();
            if (!label) return;
            if (mode === "heavy" && mqlIsRequiredWeapon(norm || label)) return;
            if (!weaponMap.has(key)) weaponMap.set(key, norm || label);
          });
          const weapons = [];
          if (mode === "standard") {
            (Array.isArray(MQL_REQUIRED_WEAPON_LABELS) ? MQL_REQUIRED_WEAPON_LABELS : []).forEach(function (label) {
              weapons.push(label);
            });
          }
          Array.from(weaponMap.values())
            .filter(function (label) {
              if (mode === "heavy") return !mqlIsRequiredWeapon(label);
              return !(Array.isArray(MQL_REQUIRED_WEAPON_LABELS) ? MQL_REQUIRED_WEAPON_LABELS : []).some(function (req) {
                return String(req).toLowerCase() === String(label).toLowerCase();
              });
            })
            .sort(function (a, b) {
              return a.localeCompare(b, undefined, { sensitivity: "base" });
            })
            .forEach(function (label) {
              weapons.push(label);
            });

          const bylawMap = new Map();
          if (mode !== "heavy") {
            mqlMergeBylawCatalogLabels(bylawMap, opts.bylawItemChoices);
            (Array.isArray(bylawRows) ? bylawRows : []).forEach(function (row) {
              const label = bylawLabelFromRow(row);
              if (!label || mqlIsExcludedCatalogLabel(label)) return;
              const key = String(label).toLowerCase();
              if (!bylawMap.has(key)) bylawMap.set(key, label);
            });
          }
          const bylaws =
            mode === "heavy"
              ? []
              : Array.from(bylawMap.values()).sort(function (a, b) {
                  return a.localeCompare(b, undefined, { sensitivity: "base" });
                });
          return { weapons: weapons, bylaws: bylaws, mode: mode };
        }

        function indexCertRowsByPersonLabel(rows, personIdFn, labelFn, qualKeys) {
          const out = {};
          const dateKeys = qualKeys || weaponsQualDateKeys();
          (Array.isArray(rows) ? rows : []).forEach(function (row) {
            const pid = personIdFn(row);
            if (!pid) return;
            const label = labelFn(row);
            if (!label) return;
            if (!out[pid]) out[pid] = {};
            const existing = out[pid][label];
            if (!existing) {
              out[pid][label] = row;
              return;
            }
            const da = parseWeaponsCertCalendarDate(valueFromItemByKeys(existing, dateKeys));
            const db = parseWeaponsCertCalendarDate(valueFromItemByKeys(row, dateKeys));
            if (!da || (db && db.getTime() >= da.getTime())) out[pid][label] = row;
          });
          return out;
        }

        function filterHeavyWeaponsMqlRows(rows, catalog) {
          const heavyLabels = catalog && catalog.weapons ? catalog.weapons : [];
          if (!heavyLabels.length) return [];
          return (Array.isArray(rows) ? rows : []).filter(function (entry) {
            return heavyLabels.some(function (label) {
              return !!mqlLookupCertRow(entry.weapons, label);
            });
          });
        }

        function buildMqlReportRows(personRows, weaponsRows, bylawRows, opts) {
          opts = opts || {};
          const weaponsPersonIdFn = opts.weaponsPersonIdFn || weaponsPersonnelIdFromItem;
          const bylawPersonIdFn = opts.bylawPersonIdFn || bylawPersonnelIdFromItem;
          const weaponLabelFn = opts.weaponLabelFn || weaponLabelFromRow;
          const bylawLabelFn = opts.bylawLabelFn || bylawLabelFromRow;
          const weaponsByPerson = indexCertRowsByPersonLabel(weaponsRows, weaponsPersonIdFn, weaponLabelFn, weaponsQualDateKeys());
          const bylawByPerson = indexCertRowsByPersonLabel(
            bylawRows,
            bylawPersonIdFn,
            bylawLabelFn,
            bylawQualDateKeys(),
          );
          return (Array.isArray(personRows) ? personRows : [])
            .filter(function (person) {
              return person && person.Id != null;
            })
            .map(function (person) {
              const pid = String(person.Id);
              return {
                person: person,
                weapons: mqlAugmentCertMapAliases(Object.assign({}, weaponsByPerson[pid] || {})),
                bylaws: mqlAugmentCertMapAliases(Object.assign({}, bylawByPerson[pid] || {})),
              };
            });
        }

        function formatPersonFullName(item) {
          const last = itemFieldText(item, "LastName");
          const first = itemFieldText(item, "FirstName");
          const mi = itemFieldText(item, "MiddleInitial");
          let name = [last, first].filter(Boolean).join(", ");
          if (mi) name = name ? name + " " + mi + "." : mi + ".";
          return name || "-";
        }

        function mqlDisplayOpts(opts) {
          opts = opts || {};
          return {
            displayNameFn: opts.displayNameFn || formatPersonDisplayName,
            statusLabelFn: opts.statusLabelFn || personStatusLabel,
            rankFn: opts.rankFn || function (person) {
              return itemFieldText(person, "Rank") || "-";
            },
            navigateToPerson: opts.navigateToPerson !== false,
          };
        }

        function mqlIsHeavyCatalog(catalog) {
          return !!(catalog && catalog.mode === "heavy");
        }

        function mqlStandardColumnLabels() {
          return (Array.isArray(MQL_STANDARD_COLUMN_DEFS) ? MQL_STANDARD_COLUMN_DEFS : []).map(function (col) {
            return col.label;
          });
        }

        function mqlMatchCatalogLabel(storedLabel, wantedLabel) {
          const norm = function (s) {
            return String(s || "")
              .trim()
              .toLowerCase()
              .replace(/[\s\-_./]/g, "");
          };
          const a = norm(storedLabel);
          const b = norm(wantedLabel);
          if (!a || !b) return false;
          if (a === b || a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return true;
          const digitsA = a.replace(/\D/g, "");
          const digitsB = b.replace(/\D/g, "");
          if (digitsB.length >= 3 && digitsA && (digitsA === digitsB || digitsA.indexOf(digitsB) >= 0)) return true;
          const stripNoise = function (x) {
            return x.replace(/form/g, "").replace(/dd/g, "").replace(/exp/g, "").replace(/date/g, "");
          };
          const sa = stripNoise(a);
          const sb = stripNoise(b);
          return !!(sa && sb && (sa === sb || sa.indexOf(sb) >= 0 || sb.indexOf(sa) >= 0));
        }

        function mqlLookupCertRowByItem(certMap, itemLabel, matchLabels) {
          if (!certMap || !itemLabel) return null;
          const wanted = [itemLabel].concat(Array.isArray(matchLabels) ? matchLabels : []);
          if (certMap[itemLabel]) return certMap[itemLabel];
          const keys = Object.keys(certMap);
          for (let i = 0; i < keys.length; i++) {
            for (let w = 0; w < wanted.length; w++) {
              if (mqlMatchCatalogLabel(keys[i], wanted[w])) return certMap[keys[i]];
            }
          }
          return mqlLookupCertRow(certMap, itemLabel);
        }

        function mqlLookupTrainingRow(entry, colDef) {
          const itemLabel = colDef && colDef.item ? colDef.item : colDef;
          const matchLabels = colDef && colDef.match ? colDef.match : null;
          const preferWeapon = colDef && colDef.preferWeapon;
          const weaponRow = mqlLookupCertRowByItem(entry.weapons, itemLabel, matchLabels);
          const bylawRow = mqlLookupCertRowByItem(entry.bylaws, itemLabel, matchLabels);
          if (preferWeapon !== false) {
            if (weaponRow) return { row: weaponRow, isWeapon: true };
            if (bylawRow) return { row: bylawRow, isWeapon: false };
          } else {
            if (bylawRow) return { row: bylawRow, isWeapon: false };
            if (weaponRow) return { row: weaponRow, isWeapon: true };
          }
          return { row: null, isWeapon: preferWeapon !== false };
        }

        function mqlStandardRowCells(entry, opts) {
          const display = mqlDisplayOpts(opts);
          const person = entry.person;
          const cells = [];
          (Array.isArray(MQL_STANDARD_COLUMN_DEFS) ? MQL_STANDARD_COLUMN_DEFS : []).forEach(function (col) {
            if (col.kind === "person") {
              if (col.field === "status") cells.push({ text: display.statusLabelFn(person), tone: "ok" });
              else if (col.field === "section") cells.push({ text: personDutySectionLabel(person), tone: "ok" });
              else if (col.field === "fullName") cells.push({ text: formatPersonFullName(person), tone: "ok" });
              else if (col.field === "dodid") cells.push({ text: itemFieldText(person, "DoDID") || "-", tone: "ok" });
              return;
            }
            const lookup = mqlLookupTrainingRow(entry, col);
            const statusFn = lookup.isWeapon ? computeMqlWeaponsCertCellStatus : computeMqlBylawCertCellStatus;
            cells.push(mqlExportCertCell(lookup.row, col.part, statusFn, lookup.isWeapon));
          });
          return cells;
        }

        function mqlColumnLabels(catalog) {
          if (!mqlIsHeavyCatalog(catalog)) return mqlStandardColumnLabels();
          const labels = ["Duty Status", "Full Name"];
          (catalog && catalog.weapons ? catalog.weapons : []).forEach(function (w) {
            labels.push(w);
          });
          (catalog && catalog.bylaws ? catalog.bylaws : []).forEach(function (b) {
            labels.push(b);
          });
          return labels;
        }

        function mqlExportColumnLabels(catalog) {
          if (!mqlIsHeavyCatalog(catalog)) return mqlStandardColumnLabels();
          const labels = ["Duty Status", "Full Name"];
          (catalog && catalog.weapons ? catalog.weapons : []).forEach(function (w) {
            labels.push(w + " Qual");
            labels.push(w + " Exp");
          });
          (catalog && catalog.bylaws ? catalog.bylaws : []).forEach(function (b) {
            labels.push(b + " Qual");
            labels.push(b + " Exp");
          });
          return labels;
        }

        function mqlRowCells(entry, catalog, opts) {
          const display = mqlDisplayOpts(opts);
          const person = entry.person;
          const cells = [
            display.statusLabelFn(person),
            display.displayNameFn(person),
          ];
          (catalog && catalog.weapons ? catalog.weapons : []).forEach(function (label) {
            cells.push(mqlCertCell(mqlLookupCertRow(entry.weapons, label), computeMqlWeaponsCertCellStatus));
          });
          (catalog && catalog.bylaws ? catalog.bylaws : []).forEach(function (label) {
            cells.push(mqlCertCell(mqlLookupCertRow(entry.bylaws, label), computeMqlBylawCertCellStatus));
          });
          return cells;
        }

        function mqlExportRowCells(entry, catalog, opts) {
          const display = mqlDisplayOpts(opts);
          const person = entry.person;
          const cells = [
            display.statusLabelFn(person),
            display.displayNameFn(person),
          ];
          (catalog && catalog.weapons ? catalog.weapons : []).forEach(function (label) {
            const row = mqlLookupCertRow(entry.weapons, label);
            cells.push(mqlExportCertCell(row, "qual", computeMqlWeaponsCertCellStatus, true));
            cells.push(mqlExportCertCell(row, "exp", computeMqlWeaponsCertCellStatus, true));
          });
          (catalog && catalog.bylaws ? catalog.bylaws : []).forEach(function (label) {
            const row = mqlLookupCertRow(entry.bylaws, label);
            cells.push(mqlExportCertCell(row, "qual", computeMqlBylawCertCellStatus, false));
            cells.push(mqlExportCertCell(row, "exp", computeMqlBylawCertCellStatus, false));
          });
          return cells;
        }

        function appendMqlTableCell(tr, cell, idx, entry, display) {
          const td = document.createElement("td");
          const text = cell && typeof cell === "object" && cell.text != null ? cell.text : cell;
          const nameColIndex = mqlIsHeavyCatalog(display.catalog) ? 1 : 2;
          if (idx === nameColIndex && display.navigateToPerson && entry.person && entry.person.Id != null) {
            const nameBtn = document.createElement("button");
            nameBtn.type = "button";
            nameBtn.className = "reports-name-link";
            nameBtn.textContent = text;
            nameBtn.title = "Open Personnel Record";
            nameBtn.addEventListener("click", function () {
              navigateToPersonDetail(entry.person.Id);
            });
            td.appendChild(nameBtn);
          } else {
            td.textContent = displayCellText(text);
          }
          tr.appendChild(td);
        }

        function groupMqlEntriesByDutyStatus(entries) {
          const groups = new Map();
          (Array.isArray(entries) ? entries : []).forEach(function (entry) {
            const status = entry && entry.person ? personStatusLabel(entry.person) : "Unknown";
            if (!groups.has(status)) groups.set(status, []);
            groups.get(status).push(entry);
          });
          const out = [];
          groups.forEach(function (items, section) {
            out.push({ section: section, items: sortMqlReportRows(items) });
          });
          out.sort(function (a, b) {
            return String(a.section).localeCompare(String(b.section), undefined, { sensitivity: "base" });
          });
          return out;
        }

        function sortMqlReportRows(rows, opts) {
          return rows.slice().sort(function (a, b) {
            const lastCmp = itemFieldText(a.person, "LastName").localeCompare(itemFieldText(b.person, "LastName"), undefined, {
              sensitivity: "base",
            });
            if (lastCmp !== 0) return lastCmp;
            return itemFieldText(a.person, "FirstName").localeCompare(itemFieldText(b.person, "FirstName"), undefined, {
              sensitivity: "base",
            });
          });
        }

        function buildMqlFlatPrintTable(rows, catalog, display, options) {
          options = options || {};
          const columnLabels = mqlIsHeavyCatalog(catalog) ? mqlExportColumnLabels(catalog) : mqlStandardColumnLabels();
          const buildCells = mqlIsHeavyCatalog(catalog)
            ? function (entry) {
                return mqlExportRowCells(entry, catalog, display);
              }
            : function (entry) {
                return mqlStandardRowCells(entry, display);
              };
          const widths = officialColumnWidths(columnLabels);
          const table = document.createElement("table");
          table.className = "af-official-table af-official-table--columns af-official-table--compact";
          table.style.width = "100%";
          appendOfficialColgroup(table, widths);
          const thead = document.createElement("thead");
          const trHead = document.createElement("tr");
          columnLabels.forEach(function (label) {
            const th = document.createElement("th");
            th.textContent = label;
            trHead.appendChild(th);
          });
          thead.appendChild(trHead);
          table.appendChild(thead);
          const tbody = document.createElement("tbody");
          sortMqlReportRows(rows, display).forEach(function (entry) {
            const tr = document.createElement("tr");
            buildCells(entry).forEach(function (cell) {
              const td = document.createElement("td");
              const text = cell && typeof cell === "object" && cell.text != null ? cell.text : cell;
              td.textContent = text === "" || text == null ? "-" : String(text);
              if (options.colorize) {
                const tone = cell && typeof cell === "object" ? cell.tone : null;
                applyMqlCertCellStyle(td, tone);
              }
              tr.appendChild(td);
            });
            tbody.appendChild(tr);
          });
          table.appendChild(tbody);
          return table;
        }

        async function ensureMqlReportData(pw, opts) {
          opts = opts || {};
          const mode = opts.mode === "heavy" ? "heavy" : "standard";
          reportsSession.mqlMode = mode;
          const cache = await ensureReportsTrainingCache(pw);
          const sampleBylaw = cache.bylawRows && cache.bylawRows.length ? cache.bylawRows[0] : null;
          const bylawItemChoices = await fetchMqlBylawItemChoices(
            pw,
            bylawTrainingListApiPath,
            sampleBylaw,
            reportsSession,
            "bylawItemChoices",
            cache.bylawRows,
          );
          const catalog = buildMqlCertCatalog(cache.weaponsRows, cache.bylawRows, {
            mode: mode,
            bylawItemChoices: bylawItemChoices,
          });
          const personRows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
          let rows = buildMqlReportRows(personRows, cache.weaponsRows, cache.bylawRows, opts);
          if (mode === "heavy") rows = filterHeavyWeaponsMqlRows(rows, catalog);
          reportsSession.mqlCatalog = catalog;
          reportsSession.mqlRows = rows;
          return { rows: rows, catalog: catalog };
        }

        function renderMqlReportTable(rows, catalog, opts) {
          const display = mqlDisplayOpts(opts);
          display.catalog = catalog;
          const wrap = document.createElement("div");
          wrap.className = "roster-wrap reports-mql-wrap";
          const columnLabels = mqlColumnLabels(catalog);
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "reports-office-empty";
            p.textContent = "No Personnel on file.";
            wrap.appendChild(p);
            return wrap;
          }
          const buildCells = mqlIsHeavyCatalog(catalog)
            ? function (entry) {
                return mqlRowCells(entry, catalog, opts);
              }
            : function (entry) {
                return mqlStandardRowCells(entry, opts);
              };
          const appendTable = function (tableRows) {
            const table = document.createElement("table");
            table.className = "roster reports-status-table reports-mql-table";
            table.setAttribute("aria-label", (opts && opts.tableLabel) || "MQL Report");
            const thead = document.createElement("thead");
            const trHead = document.createElement("tr");
            columnLabels.forEach(function (label) {
              const th = document.createElement("th");
              th.textContent = label;
              trHead.appendChild(th);
            });
            thead.appendChild(trHead);
            table.appendChild(thead);
            const tbody = document.createElement("tbody");
            const frag = document.createDocumentFragment();
            tableRows.forEach(function (entry) {
              const tr = document.createElement("tr");
              buildCells(entry).forEach(function (cell, idx) {
                appendMqlTableCell(tr, cell, idx, entry, display);
              });
              frag.appendChild(tr);
            });
            tbody.appendChild(frag);
            table.appendChild(tbody);
            wrap.appendChild(table);
          };
          if (!mqlIsHeavyCatalog(catalog)) {
            groupMqlEntriesByDutyStatus(rows).forEach(function (group) {
              const heading = document.createElement("h3");
              heading.className = "reports-office-title";
              heading.textContent = group.section;
              wrap.appendChild(heading);
              appendTable(group.items);
            });
            return wrap;
          }
          appendTable(sortMqlReportRows(rows, opts));
          return wrap;
        }

        function renderMqlSignedPrintDocument(rows, catalog, opts) {
          opts = opts || {};
          const display = mqlDisplayOpts(opts);
          const title =
            opts.printTitle ||
            (catalog && catalog.mode === "heavy"
              ? "Heavy Weapons Qual / Authorization Report"
              : MQL_REPORT_PRINT_TITLE);
          const wrap = document.createElement("div");
          wrap.className = "af-official-doc af-official-doc--mql-landscape af-official-doc--mql-signed";
          wrap.appendChild(
            buildMqlSignedPrintHeader(title, {
              unitLabel: opts.unitLabel || String(AF_REPORT_UNIT_LABEL || "88th Security Forces Squadron"),
              sortLine: MQL_SIGNED_SORT_LINE,
            }),
          );
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "af-official-doc-empty";
            p.textContent = "No members on file.";
            wrap.appendChild(p);
            return wrap;
          }
          if (!mqlIsHeavyCatalog(catalog)) {
            appendOfficialSectionTable(wrap, {
              columns: mqlStandardColumnLabels(),
              sections: groupMqlEntriesByDutyStatus(rows),
              compact: true,
              buildCells: function (entry) {
                return mqlStandardRowCells(entry, display);
              },
            });
            wrap.appendChild(buildMqlPrintSignatureBlock(opts.signature || hubSession.mqlSignature, false));
            return wrap;
          }
          appendOfficialSectionTable(wrap, {
            columns: mqlColumnLabels(catalog),
            sections: groupReportEntriesBySection(rows, function (entry) {
              return entry.person;
            }),
            compact: true,
            buildCells: function (entry) {
              return mqlRowCells(entry, catalog, display);
            },
          });
          wrap.appendChild(buildMqlPrintSignatureBlock(opts.signature || hubSession.mqlSignature, false));
          return wrap;
        }

        function renderMqlExportPrintDocument(rows, catalog, opts) {
          opts = opts || {};
          const display = mqlDisplayOpts(opts);
          const wrap = document.createElement("div");
          wrap.className = "af-official-doc af-official-doc--mql-landscape mql-export-doc";
          wrap.appendChild(
            buildMqlSignedPrintHeader(opts.printTitle || MQL_EXPORT_PRINT_TITLE, {
              unitLabel: opts.unitLabel || String(AF_REPORT_UNIT_LABEL || "88th Security Forces Squadron"),
              sortLine: MQL_EXPORT_SORT_LINE,
            }),
          );
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "af-official-doc-empty";
            p.textContent = "No members on file.";
            wrap.appendChild(p);
            return wrap;
          }
          wrap.appendChild(buildMqlFlatPrintTable(rows, catalog, display, { colorize: true }));
          return wrap;
        }

        /** @deprecated use renderMqlSignedPrintDocument */
        function renderMqlReportPrintDocument(rows, catalog, opts) {
          return renderMqlSignedPrintDocument(rows, catalog, opts);
        }

        function buildCleoReportRows(personRows) {
          return (Array.isArray(personRows) ? personRows : [])
            .filter(function (person) {
              return person && personIsCleo(person);
            })
            .map(function (person) {
              return {
                person: person,
                cleoDate: formatPersonCleoDateDisplay(person),
              };
            });
        }

        function buildPostRCReportRows(personRows) {
          return (Array.isArray(personRows) ? personRows : [])
            .filter(function (person) {
              return person && person.Id != null && personHasPostRCRecord(person);
            })
            .map(function (person) {
              return {
                person: person,
                postRcQual: formatPersonPostRCQualDisplay(person),
                postRcExp: formatPersonPostRCExpiryDisplay(person),
              };
            });
        }

        function renderCleoReportTable(rows) {
          const wrap = document.createElement("div");
          wrap.className = "roster-wrap";
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "reports-office-empty";
            p.textContent = "No Personnel marked CLEO.";
            wrap.appendChild(p);
            return wrap;
          }
          const table = document.createElement("table");
          table.className = "roster reports-status-table";
          table.setAttribute("aria-label", "CLEO Report");
          const thead = document.createElement("thead");
          const trHead = document.createElement("tr");
          ["Personnel", "Certified date"].forEach(function (label) {
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
            const tdCleoDate = document.createElement("td");
            tdCleoDate.textContent = displayCellText(entry.cleoDate);
            tr.appendChild(tdCleoDate);
            frag.appendChild(tr);
          });
          tbody.appendChild(frag);
          table.appendChild(tbody);
          wrap.appendChild(table);
          return wrap;
        }

        function renderCleoReportPrintDocument(rows) {
          const wrap = document.createElement("div");
          wrap.className = "af-official-doc";
          wrap.appendChild(
            buildAfOfficialDocHeader("CLEO Report", {
              includeSeconds: true,
              subtitle: String(AF_REPORT_UNIT_LABEL || APPOINTMENTS_SQUADRON_LABEL || "88 SFS"),
            }),
          );
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "af-official-doc-empty";
            p.textContent = "No Personnel marked CLEO.";
            wrap.appendChild(p);
            return wrap;
          }
          appendOfficialSectionTable(wrap, {
            columns: ["Personnel", "Certified date"],
            sections: groupReportEntriesBySection(rows, function (entry) {
              return entry.person;
            }),
            compact: true,
            buildCells: function (entry) {
              return [formatPersonDisplayName(entry.person), displayCellText(entry.cleoDate)];
            },
          });
          return wrap;
        }

        function renderPostRCReportTable(rows) {
          const wrap = document.createElement("div");
          wrap.className = "roster-wrap";
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "reports-office-empty";
            p.textContent = "No Personnel with Post RC dates on file.";
            wrap.appendChild(p);
            return wrap;
          }
          const table = document.createElement("table");
          table.className = "roster reports-status-table";
          table.setAttribute("aria-label", "Post RC Report");
          const thead = document.createElement("thead");
          const trHead = document.createElement("tr");
          ["Personnel", "Post RC qual date", "Post RC expiration"].forEach(function (label) {
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
            const tdQual = document.createElement("td");
            tdQual.textContent = displayCellText(entry.postRcQual);
            tr.appendChild(tdQual);
            const tdExp = document.createElement("td");
            tdExp.textContent = displayCellText(entry.postRcExp);
            tr.appendChild(tdExp);
            frag.appendChild(tr);
          });
          tbody.appendChild(frag);
          table.appendChild(tbody);
          wrap.appendChild(table);
          return wrap;
        }

        function renderPostRCReportPrintDocument(rows) {
          const wrap = document.createElement("div");
          wrap.className = "af-official-doc";
          wrap.appendChild(
            buildAfOfficialDocHeader("Post RC Report", {
              includeSeconds: true,
              subtitle: String(AF_REPORT_UNIT_LABEL || APPOINTMENTS_SQUADRON_LABEL || "88 SFS"),
            }),
          );
          if (!rows.length) {
            const p = document.createElement("p");
            p.className = "af-official-doc-empty";
            p.textContent = "No Personnel with Post RC dates on file.";
            wrap.appendChild(p);
            return wrap;
          }
          appendOfficialSectionTable(wrap, {
            columns: ["Personnel", "Post RC qual date", "Post RC expiration"],
            sections: groupReportEntriesBySection(rows, function (entry) {
              return entry.person;
            }),
            compact: true,
            buildCells: function (entry) {
              return [
                formatPersonDisplayName(entry.person),
                displayCellText(entry.postRcQual),
                displayCellText(entry.postRcExp),
              ];
            },
          });
          return wrap;
        }

        async function loadMqlReport(reportDef) {
          const pw = hubSession.pw;
          if (!pw) {
            setReportsState("err", "Hub not ready. Refresh the Personnel Roster first.");
            return;
          }
          if (!reportsDetailBody) return;
          showReportsDetail(reportDef);
          reportsDetailBody.innerHTML = "";
          setReportsState("loading", "Building MQL report...");
          try {
            const data = await ensureMqlReportData(pw, { mode: "standard" });
            const sigPanel = buildMqlSignatureSettingsPanel();
            sigPanel.id = "reportsMqlSignaturePanel";
            const summary = document.createElement("p");
            summary.className = "reports-hint";
            summary.textContent =
              data.rows.length +
              " Personnel | " +
              data.catalog.weapons.length +
              " weapon column(s) | " +
              data.catalog.bylaws.length +
              " By-Law column(s). Use Print signed report for the commander copy.";
            reportsDetailBody.appendChild(sigPanel);
            void populateMqlSignaturePanel(sigPanel);
            reportsDetailBody.appendChild(summary);
            reportsDetailBody.appendChild(renderMqlReportTable(data.rows, data.catalog));
            setReportsState("ok", "MQL report updated.");
            window.setTimeout(function () {
              setReportsState("", "");
            }, 2200);
          } catch (e) {
            setReportsState("err", "Could not build MQL report: " + (e.message || String(e)).slice(0, 220));
          }
        }

        async function loadMqlPdf(reportDef) {
          const pw = hubSession.pw;
          if (!pw) {
            setReportsState("err", "Hub not ready. Refresh the Personnel Roster first.");
            return;
          }
          if (!reportsDetailBody) return;
          showReportsDetail(reportDef);
          reportsDetailBody.innerHTML = "";
          setReportsState("loading", "Preparing MQL export...");
          try {
            const data = await ensureMqlReportData(pw, { mode: "standard" });
            const summary = document.createElement("p");
            summary.className = "reports-hint";
            summary.textContent =
              data.rows.length +
              " Personnel | separate qual/expiration columns for Excel or email. Use Export PDF to print or save.";
            reportsDetailBody.appendChild(summary);
            reportsDetailBody.appendChild(renderMqlReportTable(data.rows, data.catalog));
            setReportsState("ok", "MQL export ready.");
            window.setTimeout(function () {
              setReportsState("", "");
            }, 2200);
          } catch (e) {
            setReportsState("err", "Could not prepare MQL export: " + (e.message || String(e)).slice(0, 220));
          }
        }

        async function loadHeavyWeaponsMqlReport(reportDef) {
          const pw = hubSession.pw;
          if (!pw) {
            setReportsState("err", "Hub not ready. Refresh the Personnel Roster first.");
            return;
          }
          if (!reportsDetailBody) return;
          showReportsDetail(reportDef);
          reportsDetailBody.innerHTML = "";
          setReportsState("loading", "Building Heavy Weapons MQL...");
          try {
            const data = await ensureMqlReportData(pw, { mode: "heavy" });
            const sigPanel = buildMqlSignatureSettingsPanel();
            sigPanel.id = "reportsMqlSignaturePanel";
            const summary = document.createElement("p");
            summary.className = "reports-hint";
            summary.textContent =
              data.rows.length +
              " Personnel with heavy weapon quals | " +
              data.catalog.weapons.length +
              " heavy weapon column(s) (excluding M4/M18).";
            reportsDetailBody.appendChild(sigPanel);
            void populateMqlSignaturePanel(sigPanel);
            reportsDetailBody.appendChild(summary);
            reportsDetailBody.appendChild(
              renderMqlReportTable(data.rows, data.catalog, { tableLabel: "Heavy Weapons MQL" }),
            );
            setReportsState("ok", "Heavy Weapons MQL updated.");
            window.setTimeout(function () {
              setReportsState("", "");
            }, 2200);
          } catch (e) {
            setReportsState("err", "Could not build Heavy Weapons MQL: " + (e.message || String(e)).slice(0, 220));
          }
        }

        async function loadCleoReport(reportDef) {
          const pw = hubSession.pw;
          if (!pw) {
            setReportsState("err", "Hub not ready. Refresh the Personnel Roster first.");
            return;
          }
          if (!reportsDetailBody) return;
          showReportsDetail(reportDef);
          reportsDetailBody.innerHTML = "";
          setReportsState("loading", "Building CLEO report...");
          try {
            const personRows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
            const rows = buildCleoReportRows(personRows);
            const summary = document.createElement("p");
            summary.className = "reports-hint";
            summary.textContent = rows.length + " CLEO member(s) on file.";
            reportsDetailBody.appendChild(summary);
            reportsDetailBody.appendChild(renderCleoReportTable(rows));
            setReportsState("ok", "CLEO report updated.");
            window.setTimeout(function () {
              setReportsState("", "");
            }, 2200);
          } catch (e) {
            setReportsState("err", "Could not build CLEO report: " + (e.message || String(e)).slice(0, 220));
          }
        }

        async function loadPostRCReport(reportDef) {
          const pw = hubSession.pw;
          if (!pw) {
            setReportsState("err", "Hub not ready. Refresh the Personnel Roster first.");
            return;
          }
          if (!reportsDetailBody) return;
          showReportsDetail(reportDef);
          reportsDetailBody.innerHTML = "";
          setReportsState("loading", "Building Post RC report...");
          try {
            const personRows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
            const rows = buildPostRCReportRows(personRows);
            const summary = document.createElement("p");
            summary.className = "reports-hint";
            summary.textContent = rows.length + " Personnel with Post RC dates on file.";
            reportsDetailBody.appendChild(summary);
            reportsDetailBody.appendChild(renderPostRCReportTable(rows));
            setReportsState("ok", "Post RC report updated.");
            window.setTimeout(function () {
              setReportsState("", "");
            }, 2200);
          } catch (e) {
            setReportsState("err", "Could not build Post RC report: " + (e.message || String(e)).slice(0, 220));
          }
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
          reportsSession.sotView = "active";
          reportsSession.selectedMonth = currentYearMonthKey();
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
            reportsSession.sotRows = rows;
            const monthOptions = buildSotMonthOptions();
            const shell = buildStatusOfTrainingShell(
              [renderStatusOfTrainingSummary(rows), renderStatusOfTrainingTable(rows)],
              monthOptions,
            );
            reportsDetailBody.appendChild(shell);
            wireStatusOfTrainingControls();
            setSotView("active");
            refreshMonthlySotView();
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
          if (def.implemented && def.id === "mql-report") {
            await loadMqlReport(def);
            return;
          }
          if (def.implemented && def.id === "mql-pdf") {
            await loadMqlPdf(def);
            return;
          }
          if (def.implemented && def.id === "heavy-weapons") {
            await loadHeavyWeaponsMqlReport(def);
            return;
          }
          if (def.implemented && def.id === "cleo-report") {
            await loadCleoReport(def);
            return;
          }
          if (def.implemented && def.id === "post-rc-report") {
            await loadPostRCReport(def);
            return;
          }
          renderPlaceholderReport(def);
        }

        function invalidateReportsTrainingCache() {
          reportsSession.weaponsRows = null;
          reportsSession.bylawRows = null;
          reportsSession.mqlRows = null;
          reportsSession.mqlCatalog = null;
          reportsSession.bylawItemChoices = null;
          reportsSession.mqlMode = "standard";
        }

        async function printActiveReport() {
          const def = reportTypeById(reportsSession.activeReportId);
          if (!def || !def.printable) return;
          if (def.id === "status-of-training") {
            if (reportsSession.sotView !== "monthly") return;
            window.print();
            return;
          }
          const personRows = Array.isArray(hubSession.rows) ? hubSession.rows : [];
          const pw = hubSession.pw;
          if (def.id === "mql-report") {
            if (!pw) return;
            const data =
              reportsSession.mqlRows && reportsSession.mqlCatalog && reportsSession.mqlMode === "standard"
                ? { rows: reportsSession.mqlRows, catalog: reportsSession.mqlCatalog }
                : await ensureMqlReportData(pw, { mode: "standard" });
            await loadMqlSignature(pw, false);
            const sigPanel =
              reportsDetailBody && reportsDetailBody.querySelector(".reports-mql-signature-panel");
            const sig = currentMqlSignatureForPrint(sigPanel || reportsDetailBody);
            triggerMqlPrint(renderMqlSignedPrintDocument(data.rows, data.catalog, { signature: sig }), {
              signed: true,
              signature: sig,
            });
            return;
          }
          if (def.id === "heavy-weapons") {
            if (!pw) return;
            const data =
              reportsSession.mqlRows && reportsSession.mqlCatalog && reportsSession.mqlMode === "heavy"
                ? { rows: reportsSession.mqlRows, catalog: reportsSession.mqlCatalog }
                : await ensureMqlReportData(pw, { mode: "heavy" });
            await loadMqlSignature(pw, false);
            const sigPanel =
              reportsDetailBody && reportsDetailBody.querySelector(".reports-mql-signature-panel");
            const sig = currentMqlSignatureForPrint(sigPanel || reportsDetailBody);
            triggerMqlPrint(renderMqlSignedPrintDocument(data.rows, data.catalog, { signature: sig }), {
              signed: true,
              signature: sig,
            });
            return;
          }
          if (def.id === "mql-pdf") {
            if (!pw) return;
            const data =
              reportsSession.mqlRows && reportsSession.mqlCatalog && reportsSession.mqlMode === "standard"
                ? { rows: reportsSession.mqlRows, catalog: reportsSession.mqlCatalog }
                : await ensureMqlReportData(pw, { mode: "standard" });
            triggerMqlPrint(renderMqlExportPrintDocument(data.rows, data.catalog), { signed: false });
            return;
          }
          if (def.id === "cleo-report") {
            triggerSchedulingPrint(renderCleoReportPrintDocument(buildCleoReportRows(personRows)));
            return;
          }
          if (def.id === "post-rc-report") {
            triggerSchedulingPrint(renderPostRCReportPrintDocument(buildPostRCReportRows(personRows)));
            return;
          }
          window.print();
        }

        async function exportActiveReport() {
          const def = reportTypeById(reportsSession.activeReportId);
          if (!def || def.id !== "mql-pdf" || !def.printable) return;
          await printActiveReport();
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

        function formatPersonLastName(item) {
          return itemFieldText(item, "LastName") || "-";
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

        function buildCleoDisplayField(label, text, tone) {
          const fwrap = document.createElement("div");
          fwrap.className = "add-field person-detail-field person-cleo-field";
          const lab = document.createElement("div");
          lab.className = "person-detail-label";
          lab.textContent = label;
          const val = document.createElement("div");
          val.className = "person-detail-value";
          val.textContent = displayCellText(text);
          if (tone && tone !== "unknown") val.classList.add("cert-status--" + tone);
          fwrap.appendChild(lab);
          fwrap.appendChild(val);
          return fwrap;
        }

        function buildCleoDetailBlock(item) {
          const block = document.createElement("div");
          block.className = "person-cleo-block";
          const title = document.createElement("div");
          title.className = "person-cleo-block-title";
          title.textContent = "CLEO / Post RC";
          block.appendChild(title);
          const grid = document.createElement("div");
          grid.className = "person-cleo-grid";
          grid.appendChild(buildCleoDisplayField("CLEO", personIsCleo(item) ? "Yes" : "No"));
          grid.appendChild(buildCleoDisplayField("Date became CLEO", formatPersonCleoDateDisplay(item)));
          grid.appendChild(buildCleoDisplayField("Post RC qual date", formatPersonPostRCQualDisplay(item)));
          grid.appendChild(buildCleoDisplayField("Post RC expiration", formatPersonPostRCExpiryDisplay(item)));
          const postRcStatus = computePersonPostRCStatus(item);
          grid.appendChild(buildCleoDisplayField("Post RC status", postRcStatus.text, postRcStatus.tone));
          block.appendChild(grid);
          return block;
        }

        function buildCleoAddFieldWrap(col, sampleRow, idPrefix) {
          if (!col || col.key === "Id") return null;
          const writeKey = resolveWriteKey(col, sampleRow);
          if (col.inputType === "checkbox") {
            const fwrap = document.createElement("div");
            fwrap.className = "add-field person-cleo-check-wrap";
            const lab = document.createElement("label");
            lab.className = "person-cleo-check";
            lab.setAttribute("for", idPrefix + col.key);
            lab.title = "REST write key: " + writeKey;
            const input = document.createElement("input");
            input.type = "checkbox";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
            const span = document.createElement("span");
            span.textContent = col.label;
            lab.appendChild(input);
            lab.appendChild(span);
            fwrap.appendChild(lab);
            return fwrap;
          }
          const fwrap = document.createElement("div");
          fwrap.className = "add-field";
          const lab = document.createElement("label");
          lab.setAttribute("for", idPrefix + col.key);
          lab.textContent = col.label;
          lab.title = "REST write key: " + writeKey;
          fwrap.appendChild(lab);
          let input;
          if (col.inputType === "date") {
            input = document.createElement("input");
            input.type = "date";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
            if (col.readOnly) input.readOnly = true;
          } else {
            input = document.createElement("input");
            input.type = "text";
            input.id = idPrefix + col.key;
            input.dataset.writeKey = writeKey;
          }
          fwrap.appendChild(input);
          return fwrap;
        }

        function buildCleoEditBlock(item, sampleRow) {
          const cols = normalizedPersonnelCleoColumns();
          if (!cols.length) return null;
          const idPrefix = "pf_";
          const block = document.createElement("div");
          block.className = "person-cleo-block";
          const title = document.createElement("div");
          title.className = "person-cleo-block-title";
          title.textContent = "CLEO / Post RC";
          block.appendChild(title);
          const grid = document.createElement("div");
          grid.className = "person-cleo-grid person-cleo-edit-grid";
          cols.forEach(function (col) {
            const fwrap = buildCleoAddFieldWrap(col, sampleRow, idPrefix);
            if (!fwrap) return;
            const input = fwrap.querySelector("input");
            if (input) setFieldValueFromItem(input, col, item);
            grid.appendChild(fwrap);
          });
          if (!grid.childElementCount) return null;
          block.appendChild(grid);
          block.dataset.cleoEdit = "1";
          return block;
        }

        function appendCleoBlockToRightColumn(splitWrap, cleoBlock) {
          if (!splitWrap || !cleoBlock) return;
          const fieldsets = splitWrap.querySelectorAll(":scope > fieldset.add-form-group");
          const rightFs = fieldsets.length > 1 ? fieldsets[1] : fieldsets[0];
          if (rightFs) rightFs.appendChild(cleoBlock);
          else splitWrap.appendChild(cleoBlock);
        }

        function buildDetailFieldWrap(col, item) {
          const tryKeys = col.tryKeys || [col.key];
          const raw = valueFromItemByKeys(item, tryKeys);
          let text = raw !== undefined && raw !== null ? formatCellValue(raw) : "";
          if (col.inputType === "checkbox" || col.key === "IsCLEO") {
            text = parsePersonnelCleoFlag(raw) ? "Yes" : "No";
          }
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
          if (el.type === "checkbox") {
            el.checked = parsePersonnelCleoFlag(raw);
            return;
          }
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

          const cleoBlock = editing ? buildCleoEditBlock(item, sampleRow) : buildCleoDetailBlock(item);
          appendCleoBlockToRightColumn(splitWrap, cleoBlock);

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
            wirePersonPostRCQualAutoExpiry(form, "pf_");
            applyPersonPostRCExpirationFromQual("pf_");
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
          const cleoCols = normalizedPersonnelCleoColumns();
          Object.keys(body || {}).forEach(function (key) {
            if (key === "Title") {
              updated.Title = body[key];
              return;
            }
            const col =
              cols.find(function (c) {
                const writeKey = c.saveKey || c.key;
                const tryKeys = c.tryKeys || [c.key];
                return writeKey === key || tryKeys.indexOf(key) !== -1 || c.key === key;
              }) ||
              cleoCols.find(function (c) {
                const tryKeys = c.tryKeys || [c.key];
                return tryKeys.indexOf(key) !== -1 || c.key === key;
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

          normalizedPersonnelCleoColumns().forEach(function (col) {
            const el = document.getElementById("pf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWriteKey(col, s.sampleRow);
            if (col.inputType === "checkbox") {
              payload[writeKey] = personnelCleoPayloadValue(el, writeKey, s.sampleRow);
              return;
            }
            if (col.key === "PostRCExpirationDate") return;
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          const qualEl = document.getElementById("pf_PostRCQualDate");
          const expCol = normalizedPersonnelCleoColumns().find(function (c) {
            return c.key === "PostRCExpirationDate";
          });
          if (qualEl && qualEl.value && expCol) {
            const qual = parseWeaponsCertCalendarDate(qualEl.value);
            if (qual) {
              const expWriteKey = resolveWriteKey(expCol, s.sampleRow);
              payload[expWriteKey] = isoDateFromCalendarDate(expirationDateFromQualDate(qual)) + "T00:00:00Z";
            }
          }

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
            await loadPersonAppointments(updated.Id, s.pw);
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
          hideEthosSections();
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
          await loadPersonAppointments(item.Id, personDetailSession.pw);
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

        if (hubNavScheduling) {
          hubNavScheduling.addEventListener("click", function () {
            void navigateToScheduling();
          });
        }

        if (hubNavPhase1Tracker) {
          hubNavPhase1Tracker.addEventListener("click", function () {
            void navigateToPhase1();
          });
        }

        if (phase1BackLink) {
          phase1BackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (phase1AddBtn) {
          phase1AddBtn.addEventListener("click", function () {
            if (phase1AddPanel && !phase1AddPanel.hidden) {
              setPhase1AddPanelVisible(false);
              return;
            }
            openPhase1AddPanel();
          });
        }

        if (phase1RefreshBtn) {
          phase1RefreshBtn.addEventListener("click", function () {
            void loadPhase1TrackerData();
          });
        }

        if (phase1HouseHuntingYes) {
          phase1HouseHuntingYes.addEventListener("click", function () {
            void enrollPersonInPhase1(true);
          });
        }

        if (phase1HouseHuntingNo) {
          phase1HouseHuntingNo.addEventListener("click", function () {
            void enrollPersonInPhase1(false);
          });
        }

        if (phase1AddCancelBtn) {
          phase1AddCancelBtn.addEventListener("click", function () {
            setPhase1AddPanelVisible(false);
          });
        }

        if (phase1EditForm) {
          phase1EditForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitPhase1EditSave();
          });
        }

        if (phase1EditCancelBtn) {
          phase1EditCancelBtn.addEventListener("click", function () {
            setPhase1EditPanelVisible(false);
            phase1Session.editItem = null;
          });
        }

        if (phase1ArchiveBtn) {
          phase1ArchiveBtn.addEventListener("click", function () {
            void archivePhase1Record();
          });
        }

        if (phase1TabActive) {
          phase1TabActive.addEventListener("click", function () {
            setPhase1ListView("active");
          });
        }

        if (phase1TabArchive) {
          phase1TabArchive.addEventListener("click", function () {
            setPhase1ListView("archive");
          });
        }

        if (hubNavEthos) {
          hubNavEthos.addEventListener("click", function () {
            void navigateToEthos();
          });
        }

        if (ethosBackLink) {
          ethosBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (ethosMemberDetailBackLink) {
          ethosMemberDetailBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToEthosRoster();
          });
        }

        if (ethosProbeRun) {
          ethosProbeRun.addEventListener("click", function () {
            runEthosProbe().catch(function (e) {
              setEthosReadState("err", "Unexpected error: " + (e.message || String(e)).slice(0, 280));
            });
          });
        }

        if (ethosMqlReportBtn) {
          ethosMqlReportBtn.addEventListener("click", function () {
            void showEthosMqlReport();
          });
        }

        if (ethosMqlPrintBtn) {
          ethosMqlPrintBtn.addEventListener("click", function () {
            printEthosMqlReport();
          });
        }

        if (ethosMqlCloseBtn) {
          ethosMqlCloseBtn.addEventListener("click", function () {
            if (ethosMqlReportPanel) ethosMqlReportPanel.hidden = true;
            if (ethosMqlPrintBtn) ethosMqlPrintBtn.hidden = true;
            if (ethosMqlCloseBtn) ethosMqlCloseBtn.hidden = true;
          });
        }

        if (ethosMemberDetailEditBtn) {
          ethosMemberDetailEditBtn.addEventListener("click", function () {
            if (!ethosDetailSession.item) return;
            void renderEthosMemberDetailView(ethosDetailSession.item, true);
          });
        }

        if (ethosMemberDetailSaveBtn) {
          ethosMemberDetailSaveBtn.addEventListener("click", function () {
            void saveEthosMemberDetailEdits();
          });
        }

        if (ethosMemberDetailCancelBtn) {
          ethosMemberDetailCancelBtn.addEventListener("click", function () {
            if (!ethosDetailSession.item) return;
            void renderEthosMemberDetailView(ethosDetailSession.item, false);
          });
        }

        if (ethosWeaponsAddBtn) {
          ethosWeaponsAddBtn.addEventListener("click", function () {
            if (ethosWeaponsAddPanel && !ethosWeaponsAddPanel.hidden) {
              ethosWeaponsAddPanel.hidden = true;
              if (ethosWeaponsAddForm) ethosWeaponsAddForm.reset();
              return;
            }
            setEthosWeaponsBulkPanelVisible(false);
            void openEthosWeaponsCertAddPanel();
          });
        }

        if (ethosWeaponsBulkAddBtn) {
          ethosWeaponsBulkAddBtn.addEventListener("click", function () {
            if (ethosWeaponsBulkPanel && !ethosWeaponsBulkPanel.hidden) {
              setEthosWeaponsBulkPanelVisible(false);
              if (ethosWeaponsBulkForm) ethosWeaponsBulkForm.reset();
              return;
            }
            void openEthosWeaponsBulkAddPanel();
          });
        }

        if (ethosWeaponsAddForm) {
          ethosWeaponsAddForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitEthosWeaponsCertSave();
          });
        }

        if (ethosWeaponsBulkForm) {
          ethosWeaponsBulkForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitEthosWeaponsBulkSave();
          });
        }

        if (ethosWeaponsBulkSelectAll) {
          ethosWeaponsBulkSelectAll.addEventListener("click", function () {
            setBylawBulkItemChecks(ethosWeaponsBulkItems, true);
          });
        }

        if (ethosWeaponsBulkClearAll) {
          ethosWeaponsBulkClearAll.addEventListener("click", function () {
            setBylawBulkItemChecks(ethosWeaponsBulkItems, false);
          });
        }

        if (ethosWeaponsAddCancelBtn) {
          ethosWeaponsAddCancelBtn.addEventListener("click", function () {
            if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = true;
            if (ethosWeaponsAddForm) ethosWeaponsAddForm.reset();
          });
        }

        if (ethosWeaponsBulkCancelBtn) {
          ethosWeaponsBulkCancelBtn.addEventListener("click", function () {
            setEthosWeaponsBulkPanelVisible(false);
            if (ethosWeaponsBulkForm) ethosWeaponsBulkForm.reset();
          });
        }

        if (ethosBylawAddBtn) {
          ethosBylawAddBtn.addEventListener("click", function () {
            if (ethosBylawAddPanel && !ethosBylawAddPanel.hidden) {
              ethosBylawAddPanel.hidden = true;
              if (ethosBylawAddForm) ethosBylawAddForm.reset();
              return;
            }
            setEthosBylawBulkPanelVisible(false);
            void openEthosBylawTrainingAddPanel();
          });
        }

        if (ethosBylawBulkAddBtn) {
          ethosBylawBulkAddBtn.addEventListener("click", function () {
            if (ethosBylawBulkPanel && !ethosBylawBulkPanel.hidden) {
              setEthosBylawBulkPanelVisible(false);
              if (ethosBylawBulkForm) ethosBylawBulkForm.reset();
              return;
            }
            void openEthosBylawBulkAddPanel();
          });
        }

        if (ethosBylawAddForm) {
          ethosBylawAddForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitEthosBylawTrainingSave();
          });
        }

        if (ethosBylawBulkForm) {
          ethosBylawBulkForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitEthosBylawBulkSave();
          });
        }

        if (ethosBylawBulkSelectAll) {
          ethosBylawBulkSelectAll.addEventListener("click", function () {
            setBylawBulkItemChecks(ethosBylawBulkItems, true);
          });
        }

        if (ethosBylawBulkClearAll) {
          ethosBylawBulkClearAll.addEventListener("click", function () {
            setBylawBulkItemChecks(ethosBylawBulkItems, false);
          });
        }

        if (ethosBylawAddCancelBtn) {
          ethosBylawAddCancelBtn.addEventListener("click", function () {
            if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = true;
            if (ethosBylawAddForm) ethosBylawAddForm.reset();
          });
        }

        if (ethosBylawBulkCancelBtn) {
          ethosBylawBulkCancelBtn.addEventListener("click", function () {
            setEthosBylawBulkPanelVisible(false);
            if (ethosBylawBulkForm) ethosBylawBulkForm.reset();
          });
        }

        if (schedulingBackLink) {
          schedulingBackLink.addEventListener("click", function (ev) {
            ev.preventDefault();
            void navigateToRoster();
          });
        }

        if (schedulingAddBtn) {
          schedulingAddBtn.addEventListener("click", function () {
            if (!hubSession.rows || !hubSession.rows.length) {
              setSchedulingState("warn", "Personnel Roster is empty. Refresh the list first.");
              return;
            }
            buildSchedulingAddFormFields();
            populateSchedulingPersonSelect();
            setSchedulingAddPanelVisible(true);
            setSchedulingState("", "");
            if (schedulingPersonSelect) schedulingPersonSelect.focus();
          });
        }

        if (schedulingAddCancelBtn) {
          schedulingAddCancelBtn.addEventListener("click", function () {
            setSchedulingAddPanelVisible(false);
          });
        }

        if (schedulingAddForm) {
          schedulingAddForm.addEventListener("submit", function (ev) {
            void submitSchedulingAppointmentAdd(ev);
          });
        }

        if (schedulingPrintReportBtn) {
          schedulingPrintReportBtn.addEventListener("click", function () {
            void printSchedulingUnitReport(false);
          });
        }

        if (schedulingPrintMissedReportBtn) {
          schedulingPrintMissedReportBtn.addEventListener("click", function () {
            void printSchedulingUnitReport(true);
          });
        }

        if (schedulingNewMemoBtn) {
          schedulingNewMemoBtn.addEventListener("click", function () {
            openBlankSchedulingMemo();
          });
        }

        if (schedulingTabAll) {
          schedulingTabAll.addEventListener("click", function () {
            setSchedulingListView("all");
          });
        }

        if (schedulingTabMissed) {
          schedulingTabMissed.addEventListener("click", function () {
            setSchedulingListView("missed");
          });
        }

        if (schedulingTabArchive) {
          schedulingTabArchive.addEventListener("click", function () {
            setSchedulingListView("archive");
          });
        }

        if (schedulingMemoPrintBtn) {
          schedulingMemoPrintBtn.addEventListener("click", function () {
            printSchedulingMemo();
          });
        }

        if (schedulingMemoCancelBtn) {
          schedulingMemoCancelBtn.addEventListener("click", function () {
            setSchedulingMemoOverlayVisible(false);
          });
        }

        if (schedulingMemoLetterheadSaveBtn) {
          schedulingMemoLetterheadSaveBtn.addEventListener("click", function () {
            void (async function () {
              const pw = hubSession.pw;
              if (!pw) {
                setSchedulingState("warn", "Load the Personnel Roster first.");
                return;
              }
              const topMarginIn = schedulingMemoPrintTopMargin ? schedulingMemoPrintTopMargin.value : "";
              const footer = schedulingMemoLetterheadFooter ? schedulingMemoLetterheadFooter.value : "";
              try {
                setSchedulingState("loading", "Saving letterhead...");
                await saveMemoLetterhead(pw, topMarginIn, footer);
                updateMemoLetterheadPdfLink();
                setSchedulingState("ok", "Letterhead settings saved in this browser.");
                window.setTimeout(function () {
                  setSchedulingState("", "");
                }, 3200);
              } catch (e) {
                setSchedulingState("err", "Could not save letterhead: " + (e.message || String(e)).slice(0, 220));
              }
            })();
          });
        }

        schedulingMemoSignatureBlocks.forEach(function (block) {
          if (!block.personSelect) return;
          block.personSelect.addEventListener("change", function () {
            const person = block.personSelect.value ? personnelById(block.personSelect.value) : null;
            if (person) applyMemoSignatureFieldsFromPerson(person, block.titleSelect, block.squadronSelect);
          });
        });

        if (reportsPrintBtn) {
          reportsPrintBtn.addEventListener("click", function () {
            void printActiveReport();
          });
        }

        if (reportsExportBtn) {
          reportsExportBtn.addEventListener("click", function () {
            void exportActiveReport();
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
            setWeaponsBulkPanelVisible(false);
            void openWeaponsCertAddPanel();
          });
        }

        if (personWeaponsBulkAddBtn) {
          personWeaponsBulkAddBtn.addEventListener("click", function () {
            if (personWeaponsBulkPanel && !personWeaponsBulkPanel.hidden) {
              setWeaponsBulkPanelVisible(false);
              if (personWeaponsBulkForm) personWeaponsBulkForm.reset();
              return;
            }
            void openWeaponsBulkAddPanel();
          });
        }

        if (personWeaponsAddForm) {
          personWeaponsAddForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitWeaponsCertSave();
          });
        }

        if (personWeaponsBulkForm) {
          personWeaponsBulkForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitWeaponsBulkSave();
          });
        }

        if (personWeaponsBulkSelectAll) {
          personWeaponsBulkSelectAll.addEventListener("click", function () {
            setBylawBulkItemChecks(personWeaponsBulkItems, true);
          });
        }

        if (personWeaponsBulkClearAll) {
          personWeaponsBulkClearAll.addEventListener("click", function () {
            setBylawBulkItemChecks(personWeaponsBulkItems, false);
          });
        }

        if (personWeaponsAddCancelBtn) {
          personWeaponsAddCancelBtn.addEventListener("click", function () {
            setWeaponsCertAddPanelVisible(false);
            if (personWeaponsAddForm) personWeaponsAddForm.reset();
          });
        }

        if (personWeaponsBulkCancelBtn) {
          personWeaponsBulkCancelBtn.addEventListener("click", function () {
            setWeaponsBulkPanelVisible(false);
            if (personWeaponsBulkForm) personWeaponsBulkForm.reset();
          });
        }

        if (personBylawAddBtn) {
          personBylawAddBtn.addEventListener("click", function () {
            if (personBylawAddPanel && !personBylawAddPanel.hidden) {
              setBylawTrainingAddPanelVisible(false);
              if (personBylawAddForm) personBylawAddForm.reset();
              return;
            }
            setBylawBulkPanelVisible(false);
            void openBylawTrainingAddPanel();
          });
        }

        if (personBylawBulkAddBtn) {
          personBylawBulkAddBtn.addEventListener("click", function () {
            if (personBylawBulkPanel && !personBylawBulkPanel.hidden) {
              setBylawBulkPanelVisible(false);
              if (personBylawBulkForm) personBylawBulkForm.reset();
              return;
            }
            void openBylawBulkAddPanel();
          });
        }

        if (personBylawAddForm) {
          personBylawAddForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitBylawTrainingSave();
          });
        }

        if (personBylawBulkForm) {
          personBylawBulkForm.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitBylawBulkSave();
          });
        }

        if (personBylawBulkSelectAll) {
          personBylawBulkSelectAll.addEventListener("click", function () {
            setBylawBulkItemChecks(personBylawBulkItems, true);
          });
        }

        if (personBylawBulkClearAll) {
          personBylawBulkClearAll.addEventListener("click", function () {
            setBylawBulkItemChecks(personBylawBulkItems, false);
          });
        }

        if (personBylawAddCancelBtn) {
          personBylawAddCancelBtn.addEventListener("click", function () {
            setBylawTrainingAddPanelVisible(false);
            if (personBylawAddForm) personBylawAddForm.reset();
          });
        }

        if (personBylawBulkCancelBtn) {
          personBylawBulkCancelBtn.addEventListener("click", function () {
            setBylawBulkPanelVisible(false);
            if (personBylawBulkForm) personBylawBulkForm.reset();
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
          if (el.type === "checkbox") return el.checked;
          const v = String(el.value || "").trim();
          if (v === "") return null;
          if (el.type === "date") return v + "T00:00:00Z";
          if (el.type === "datetime-local") return v.length === 16 ? v + ":00Z" : v;
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

        // ----- ETHOS Members -----
        function normalizedEthosRosterColumns() {
          const arr = Array.isArray(ETHOS_ROSTER_COLUMNS) ? ETHOS_ROSTER_COLUMNS : [];
          const result = [];
          arr.forEach(function (entry) {
            if (!entry || !entry.key) return;
            const key = String(entry.key).trim();
            const label = String(entry.label || key).trim() || key;
            const alt = Array.isArray(entry.altKeys) ? entry.altKeys.map(function (x) { return String(x).trim(); }).filter(Boolean) : [];
            const tryKeys = [key];
            alt.forEach(function (a) {
              if (tryKeys.indexOf(a) === -1) tryKeys.push(a);
            });
            result.push({ key: key, label: label, tryKeys: tryKeys, saveKey: entry.saveKey || null });
          });
          return result;
        }

        function resolveEthosRosterColumnPlan(rows) {
          const explicit = normalizedEthosRosterColumns();
          if (explicit.length > 0) return { mode: "explicit", columns: explicit };
          const keys = collectColumnKeys(rows);
          return { mode: "auto", columns: keys.map(function (k) { return { key: k, label: k, tryKeys: [k], saveKey: null }; }) };
        }

        function sortEthosRowsAlphabetical(rows) {
          if (!Array.isArray(rows) || rows.length < 2) return rows;
          const cols = normalizedEthosRosterColumns();
          function colFor(fieldKey) {
            return cols.find(function (c) { return c.key === fieldKey; }) || { key: fieldKey, tryKeys: sortKeysForPersonnelName(fieldKey) };
          }
          const lastCol = colFor("LastName");
          const firstCol = colFor("FirstName");
          function nameVal(item, col) {
            const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
            return String(raw != null ? raw : "").trim().toLowerCase();
          }
          return rows.slice().sort(function (a, b) {
            const la = nameVal(a, lastCol);
            const lb = nameVal(b, lastCol);
            if (la !== lb) return la < lb ? -1 : 1;
            const fa = nameVal(a, firstCol);
            const fb = nameVal(b, firstCol);
            if (fa !== fb) return fa < fb ? -1 : 1;
            return (a.Id || 0) - (b.Id || 0);
          });
        }

        function ethosMembersListTitle() { return String(LIST_ETHOS_MEMBERS || "").trim(); }
        function ethosMembersGuidRaw() { return String(LIST_ETHOS_MEMBERS_GUID || "").trim(); }
        function ethosMembersListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(ethosMembersGuidRaw());
        }
        function ethosMembersListApiPath() {
          if (ethosMembersListUsesGuid()) return "lists(guid'" + ethosMembersGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(ethosMembersListTitle()) + "')";
        }
        function ethosWeaponsCertListTitle() { return String(LIST_ETHOS_WEAPONS_CERTIFICATIONS || "").trim(); }
        function ethosWeaponsCertGuidRaw() { return String(LIST_ETHOS_WEAPONS_CERTIFICATIONS_GUID || "").trim(); }
        function ethosWeaponsCertListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(ethosWeaponsCertGuidRaw());
        }
        function ethosWeaponsCertListApiPath() {
          if (ethosWeaponsCertListUsesGuid()) return "lists(guid'" + ethosWeaponsCertGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(ethosWeaponsCertListTitle()) + "')";
        }
        function ethosBylawListTitle() { return String(LIST_ETHOS_BYLAW_TRAINING || "").trim(); }
        function ethosBylawGuidRaw() { return String(LIST_ETHOS_BYLAW_TRAINING_GUID || "").trim(); }
        function ethosBylawListUsesGuid() {
          return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(ethosBylawGuidRaw());
        }
        function ethosBylawTrainingListApiPath() {
          if (ethosBylawListUsesGuid()) return "lists(guid'" + ethosBylawGuidRaw() + "')";
          return "lists/getbytitle('" + escListTitle(ethosBylawListTitle()) + "')";
        }

        function normalizedEthosWeaponsCertColumns() {
          const arr = Array.isArray(ETHOS_WEAPONS_CERT_COLUMNS) ? ETHOS_WEAPONS_CERT_COLUMNS : [];
          return arr.map(function (entry) {
            if (!entry || !entry.key) return null;
            const key = String(entry.key).trim();
            const alt = Array.isArray(entry.altKeys) ? entry.altKeys.map(function (x) { return String(x).trim(); }).filter(Boolean) : [];
            const tryKeys = [key];
            alt.forEach(function (a) { if (tryKeys.indexOf(a) === -1) tryKeys.push(a); });
            return { key: key, label: String(entry.label || key).trim() || key, tryKeys: tryKeys, computed: !!entry.computed };
          }).filter(Boolean);
        }
        function ethosWeaponsCertFormColumns() {
          return normalizedEthosWeaponsCertColumns().filter(function (c) { return !c.computed; });
        }
        function ethosWeaponsCertExpiryDateKeys() {
          const col = normalizedEthosWeaponsCertColumns().find(function (c) { return c.key === "ExpirationDate" || c.key === "ExpiryDate"; });
          return col && col.tryKeys ? col.tryKeys.slice() : ["ExpirationDate", "ExpiryDate"];
        }
        function computeEthosWeaponsCertStatus(item) {
          return computeCertStatusFromExpiryKeys(item, ethosWeaponsCertExpiryDateKeys(), certQualDateKeysFromColumns(normalizedEthosWeaponsCertColumns()));
        }
        function normalizedEthosBylawColumns() {
          const arr = Array.isArray(ETHOS_BYLAW_TRAINING_COLUMNS) ? ETHOS_BYLAW_TRAINING_COLUMNS : [];
          return arr.map(function (entry) {
            if (!entry || !entry.key) return null;
            const key = String(entry.key).trim();
            const alt = Array.isArray(entry.altKeys) ? entry.altKeys.map(function (x) { return String(x).trim(); }).filter(Boolean) : [];
            const tryKeys = [key];
            alt.forEach(function (a) { if (tryKeys.indexOf(a) === -1) tryKeys.push(a); });
            return { key: key, label: String(entry.label || key).trim() || key, tryKeys: tryKeys, computed: !!entry.computed };
          }).filter(Boolean);
        }
        function ethosBylawFormColumns() {
          return normalizedEthosBylawColumns().filter(function (c) { return !c.computed; });
        }
        function ethosBylawExpiryDateKeys() {
          const col = normalizedEthosBylawColumns().find(function (c) { return c.key === "ExpirationDate" || c.key === "ExpiryDate"; });
          return col && col.tryKeys ? col.tryKeys.slice() : ["ExpirationDate", "ExpiryDate"];
        }
        function computeEthosBylawTrainingStatus(item) {
          return computeCertStatusFromExpiryKeys(item, ethosBylawExpiryDateKeys(), certQualDateKeysFromColumns(normalizedEthosBylawColumns()));
        }

        function linkedMemberFilterCandidates(primary, alts) {
          const out = [];
          const p = String(primary || "").trim();
          if (p) out.push(p);
          (Array.isArray(alts) ? alts : []).forEach(function (name) {
            const n = String(name || "").trim();
            if (n && out.indexOf(n) === -1) out.push(n);
          });
          return out;
        }

        function linkedMemberIdFromItem(item, candidates, lookupBase) {
          if (!item) return null;
          for (let i = 0; i < candidates.length; i++) {
            const k = candidates[i];
            if (!k || k.indexOf("/") !== -1) continue;
            if (Object.prototype.hasOwnProperty.call(item, k) && item[k] !== null && item[k] !== "") return String(item[k]);
          }
          if (lookupBase && item[lookupBase] && typeof item[lookupBase] === "object" && item[lookupBase].Id != null) {
            return String(item[lookupBase].Id);
          }
          return null;
        }

        async function fetchLinkedCertRowsForMember(seg, pw, memberId, opts) {
          const id = parseInt(String(memberId), 10);
          if (!id || isNaN(id)) return [];
          const orderByClause = String(opts.orderBy || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          const candidates = linkedMemberFilterCandidates(opts.personField, opts.personFieldAlt);
          if (opts.session && opts.session[opts.filterFieldsKey] && opts.session[opts.filterFieldsKey].length) {
            opts.session[opts.filterFieldsKey].forEach(function (name) {
              if (name && candidates.indexOf(name) === -1) candidates.unshift(name);
            });
          }
          const tried = [];
          for (let i = 0; i < candidates.length; i++) {
            const personField = candidates[i];
            if (!personField || tried.indexOf(personField) !== -1) continue;
            tried.push(personField);
            const filter = encodeURIComponent(personField + " eq " + id);
            try {
              let data = null;
              try {
                data = await spFetch(`/_api/web/${seg}/items?$top=200&$filter=${filter}` + orderByQs, {}, pw);
              } catch (e0) {
                if (/\b400\b/.test(String(e0.message || "")) && orderByQs) {
                  data = await spFetch(`/_api/web/${seg}/items?$top=200&$filter=${filter}`, {}, pw);
                } else throw e0;
              }
              if (opts.session) opts.session[opts.filterFieldKey] = personField;
              return (data && data.value) || [];
            } catch (e1) {
              if (!/\b400\b/.test(String(e1.message || "")) && !/does not exist/i.test(String(e1.message || ""))) throw e1;
            }
          }
          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=500` + orderByQs, {}, pw);
          } catch (e2) {
            if (/\b400\b/.test(String(e2.message || "")) && orderByQs) data = await spFetch(`/_api/web/${seg}/items?$top=500`, {}, pw);
            else throw e2;
          }
          const all = (data && data.value) || [];
          return all.filter(function (row) {
            return linkedMemberIdFromItem(row, candidates, opts.lookupBase) === String(id);
          });
        }

        async function resolveLinkedMemberPostKey(seg, pw, session, filterFieldKey, postKeyKey, defaultField) {
          if (session[postKeyKey]) return session[postKeyKey];
          const filterField = String(session[filterFieldKey] || defaultField || "").trim();
          let postKey = filterField;
          try {
            const esc = filterField.replace(/'/g, "''");
            const data = await spFetch(`/_api/web/${seg}/fields/getbyinternalnameortitle('${esc}')?$select=InternalName,TypeAsString`, {}, pw);
            const internal = String(data.InternalName || filterField).trim();
            if (/lookup/i.test(String(data.TypeAsString || ""))) postKey = internal.endsWith("Id") ? internal : internal + "Id";
            else postKey = internal;
          } catch (_) {
            if (/Id$/i.test(filterField)) postKey = filterField;
            else postKey = filterField + "Id";
          }
          session[postKeyKey] = postKey;
          return postKey;
        }

        function setEthosReadState(kind, message) {
          if (!ethosReadState) return;
          if (!message) {
            ethosReadState.hidden = true;
            ethosReadState.textContent = "";
            return;
          }
          ethosReadState.hidden = false;
          ethosReadState.className = "read-state " + (kind || "");
          ethosReadState.textContent = message;
        }

        function clearEthosRosterTable() {
          const thead = document.getElementById("ethosRosterThead");
          if (thead) thead.innerHTML = "";
          if (ethosRosterTableBody) ethosRosterTableBody.innerHTML = "";
        }

        function ethosItemFieldText(item, fieldKey) {
          const col = normalizedEthosRosterColumns().find(function (c) { return c.key === fieldKey; });
          const keys = col ? col.tryKeys || [col.key] : sortKeysForPersonnelName(fieldKey);
          const raw = valueFromItemByKeys(item, keys);
          return raw !== undefined && raw !== null ? formatCellValue(raw) : "";
        }

        function formatEthosMemberDisplayName(item) {
          const rank = ethosItemFieldText(item, "Rank");
          const last = ethosItemFieldText(item, "LastName");
          const first = ethosItemFieldText(item, "FirstName");
          const mi = ethosItemFieldText(item, "MiddleInitial");
          let name = [last, first].filter(Boolean).join(", ");
          if (mi) name = name ? name + " " + mi + "." : mi + ".";
          if (rank && name) return rank + " " + name;
          return rank || name || "ETHOS Member";
        }

        function renderEthosRosterTable(rows, meta, pw, seg) {
          clearEthosRosterTable();
          const thead = document.getElementById("ethosRosterThead");
          if (!ethosRosterTableBody || !thead) return;
          const plan = resolveEthosRosterColumnPlan(rows);
          const columns = plan.columns;
          const showActions = !!(pw && seg);
          const trHead = document.createElement("tr");
          columns.forEach(function (col) {
            const th = document.createElement("th");
            th.textContent = col.label;
            trHead.appendChild(th);
          });
          if (showActions) {
            const thx = document.createElement("th");
            thx.className = "roster-actions";
            thx.textContent = " ";
            trHead.appendChild(thx);
          }
          thead.appendChild(trHead);
          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
              td.textContent = displayCellText(raw !== undefined && raw !== null ? formatCellValue(raw) : "");
              tr.appendChild(td);
            });
            if (showActions && item.Id != null) {
              const tdAct = document.createElement("td");
              tdAct.className = "roster-actions";
              const inner = document.createElement("div");
              inner.className = "roster-actions-inner";
              const recordBtn = document.createElement("button");
              recordBtn.type = "button";
              recordBtn.className = "btn-record";
              recordBtn.textContent = "Record";
              recordBtn.addEventListener("click", function () { void openEthosMemberDetailRecord(item.Id, meta, pw, seg, rows); });
              inner.appendChild(recordBtn);
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              delBtn.addEventListener("click", function () { void deleteEthosMemberRow(item.Id, pw, seg); });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          ethosRosterTableBody.appendChild(frag);
          const listTitle = meta && meta.Title ? String(meta.Title) : ethosMembersListTitle();
          if (rows.length > 0) {
            setEthosReadState("ok", "Loaded " + rows.length + " ETHOS member(s). List \"" + listTitle + "\".");
          } else {
            setEthosReadState(
              "warn",
              "No ETHOS members in list \"" + listTitle + "\". Use Add ETHOS Member below or click Refresh list.",
            );
          }
        }

        async function deleteEthosMemberRow(id, pw, seg) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid)) { setEthosReadState("err", "Invalid row Id for delete."); return; }
          if (!confirm("Delete ETHOS Member Id " + sid + "? This cannot be undone.")) return;
          try {
            setEthosReadState("loading", "Deleting member...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            await runEthosProbe();
            setEthosReadState("ok", "Deleted member Id " + sid + ".");
          } catch (e) {
            setEthosReadState("err", "Delete failed: " + (e.message || String(e)).slice(0, 280));
          }
        }

        function ethosAddFormGroupedKeySet() {
          const s = new Set();
          if (!Array.isArray(ETHOS_ADD_FORM_FIELD_GROUPS)) return s;
          ETHOS_ADD_FORM_FIELD_GROUPS.forEach(function (g) {
            (g.keys || []).forEach(function (k) { s.add(k); });
          });
          return s;
        }

        async function renderEthosAddMemberForm(meta, pw, seg, sampleRow) {
          const panel = document.getElementById("ethosAddMemberPanel");
          if (!panel) return;
          panel.innerHTML = "";
          const cols = normalizedEthosRosterColumns();
          if (!cols.length) return;
          const h = document.createElement("h2");
          h.className = "hub-subtitle";
          h.textContent = "Add ETHOS Member";
          panel.appendChild(h);
          const form = document.createElement("form");
          form.id = "newEthosMemberForm";
          form.className = "add-form-stack";
          form.addEventListener("submit", function (ev) {
            ev.preventDefault();
            void submitNewEthosMember(meta, pw, seg, sampleRow);
          });
          let splitWrap = null;
          if (Array.isArray(ETHOS_ADD_FORM_FIELD_GROUPS)) {
            ETHOS_ADD_FORM_FIELD_GROUPS.forEach(function (group) {
              if (!group || !group.title || !Array.isArray(group.keys)) return;
              const grid = document.createElement("div");
              grid.className = "add-form-grid";
              if (group.layout === "stack") grid.classList.add("add-form-grid--stack");
              group.keys.forEach(function (key) {
                const col = cols.find(function (c) { return c.key === key; });
                if (!col) return;
                const fwrap = buildAddFieldWrap(col, sampleRow, { idPrefix: "enf_" });
                if (fwrap) grid.appendChild(fwrap);
              });
              if (!grid.childElementCount) return;
              const fs = document.createElement("fieldset");
              fs.className = "add-form-group";
              const leg = document.createElement("legend");
              leg.className = "add-form-group-legend";
              leg.textContent = group.title;
              fs.appendChild(leg);
              fs.appendChild(grid);
              if (group.split === "left") {
                splitWrap = document.createElement("div");
                splitWrap.className = "add-form-split add-form-split--primary";
                splitWrap.appendChild(fs);
                form.appendChild(splitWrap);
              } else if (group.split === "right") {
                if (!splitWrap) { splitWrap = document.createElement("div"); splitWrap.className = "add-form-split"; form.appendChild(splitWrap); }
                splitWrap.appendChild(fs);
              } else form.appendChild(fs);
            });
          }
          const bar = document.createElement("div");
          bar.className = "add-form-actions";
          const submitBtn = document.createElement("button");
          submitBtn.type = "submit";
          submitBtn.textContent = "Submit";
          bar.appendChild(submitBtn);
          form.appendChild(bar);
          panel.appendChild(form);
          for (let i = 0; i < cols.length; i++) {
            const col = cols[i];
            if (ETHOS_DROPDOWN_COLUMN_KEYS.indexOf(col.key) === -1) continue;
            const sel = document.getElementById("enf_" + col.key);
            await fillDropdownSelect(sel, col, seg, pw, sampleRow, sampleRow);
          }
          applyAllSelectAutosizes(form);
        }

        async function submitNewEthosMember(meta, pw, seg, sampleRow) {
          const cols = normalizedEthosRosterColumns();
          let lastName = "";
          let firstName = "";
          cols.forEach(function (col) {
            const el = document.getElementById("enf_" + col.key);
            if (!el) return;
            const v = String(el.value || "").trim();
            if (col.key === "LastName") lastName = v;
            if (col.key === "FirstName") firstName = v;
          });
          if (!lastName || !firstName) {
            setEthosReadState("err", "Last name and First name are required.");
            return;
          }
          const payload = {};
          cols.forEach(function (col) {
            const el = document.getElementById("enf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWriteKey(col, sampleRow);
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          if (ETHOS_SET_TITLE_ON_CREATE) payload.Title = [lastName, firstName].filter(Boolean).join(", ") || "ETHOS Member";
          try {
            setEthosReadState("loading", "Submitting...");
            await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            const form = document.getElementById("newEthosMemberForm");
            if (form) form.reset();
            await runEthosProbe();
            setEthosReadState("ok", "ETHOS member submitted. Table refreshed.");
          } catch (e) {
            setEthosReadState("err", "Submit failed: " + (e.message || String(e)).slice(0, 280));
          }
        }

        async function ensureEthosRosterRendered() {
          if (ethosSession.rows == null || !ethosSession.pw || !ethosSession.seg) return;
          renderEthosRosterTable(ethosSession.rows, ethosSession.meta, ethosSession.pw, ethosSession.seg);
          try {
            await renderEthosAddMemberForm(ethosSession.meta, ethosSession.pw, ethosSession.seg, ethosSession.sampleRow);
          } catch (e) {
            setEthosReadState(
              "warn",
              "Roster loaded but add form failed: " + (e.message || String(e)).slice(0, 220),
            );
          }
        }

        async function runEthosProbe() {
          try {
            clearEthosRosterTable();
            const addPanel = document.getElementById("ethosAddMemberPanel");
            if (addPanel) addPanel.innerHTML = "";
            setEthosReadState("loading", "Loading ETHOS Members list from SharePoint...");
            const pw = resolvedPersonnelRestBase() || ethosSession.pw || hubSession.pw;
            if (!pw) {
              setEthosReadState(
                "warn",
                "Load the Personnel Roster first (wait for list to load or click Refresh list).",
              );
              return;
            }
            const seg = ethosMembersListApiPath();
            let meta = null;
            try {
              meta = await spFetch(`/_api/web/${seg}?$select=Title,ItemCount`, {}, pw);
            } catch (e) {
              setEthosReadState("err", "Cannot read EthosMembers list: " + (e.message || String(e)).slice(0, 280));
              return;
            }
            const orderByClause = String(ETHOS_ITEMS_ORDERBY || "").trim();
            const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
            let items = null;
            try {
              items = await spFetch(`/_api/web/${seg}/items?$top=500` + orderByQs, {}, pw);
            } catch (e0) {
              if (/\b400\b/.test(String(e0.message || ""))) {
                items = await spFetch(`/_api/web/${seg}/items?$top=500`, {}, pw);
              } else {
                setEthosReadState("err", "Cannot read ETHOS items: " + (e0.message || String(e0)).slice(0, 280));
                return;
              }
            }
            let rows = sortEthosRowsAlphabetical((items && items.value) || []);
            const sampleRow = rows.length ? rows[0] : null;
            ethosSession = Object.assign({}, ethosSession, { rows: rows, meta: meta, pw: pw, seg: seg, sampleRow: sampleRow });
            renderEthosRosterTable(rows, meta, pw, seg);
            try {
              await renderEthosAddMemberForm(meta, pw, seg, sampleRow);
            } catch (eForm) {
              setEthosReadState(
                "warn",
                "List loaded but add form failed: " + (eForm.message || String(eForm)).slice(0, 220),
              );
            }
          } catch (e) {
            setEthosReadState("err", "Unexpected error loading ETHOS: " + (e.message || String(e)).slice(0, 280));
          }
        }

        function setEthosMemberDetailState(kind, message) {
          if (!ethosMemberDetailReadState) return;
          if (!message) { ethosMemberDetailReadState.hidden = true; ethosMemberDetailReadState.textContent = ""; return; }
          ethosMemberDetailReadState.hidden = false;
          ethosMemberDetailReadState.className = "read-state " + kind;
          ethosMemberDetailReadState.textContent = message;
        }

        function setEthosMemberDetailEditMode(editing) {
          ethosDetailSession.editing = !!editing;
          if (ethosMemberDetailEditBtn) ethosMemberDetailEditBtn.hidden = !!editing;
          if (ethosMemberDetailSaveBtn) ethosMemberDetailSaveBtn.hidden = !editing;
          if (ethosMemberDetailCancelBtn) ethosMemberDetailCancelBtn.hidden = !editing;
        }

        function buildEthosDetailFieldWrap(col, item) {
          const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
          const fwrap = document.createElement("div");
          fwrap.className = "add-field person-detail-field";
          const lab = document.createElement("div");
          lab.className = "person-detail-label";
          lab.textContent = col.label;
          const val = document.createElement("div");
          val.className = "person-detail-value";
          val.textContent = displayCellText(raw !== undefined && raw !== null ? formatCellValue(raw) : "");
          fwrap.appendChild(lab);
          fwrap.appendChild(val);
          return fwrap;
        }

        function buildEthosMemberDetailFields(item, editing) {
          const cols = normalizedEthosRosterColumns();
          const sampleRow = ethosDetailSession.sampleRow || item;
          const splitWrap = document.createElement("div");
          splitWrap.className = "add-form-split add-form-split--primary";
          if (Array.isArray(ETHOS_ADD_FORM_FIELD_GROUPS)) {
            ETHOS_ADD_FORM_FIELD_GROUPS.forEach(function (group) {
              const grid = document.createElement("div");
              grid.className = "add-form-grid add-form-grid--stack";
              group.keys.forEach(function (key) {
                const col = cols.find(function (c) { return c.key === key; });
                if (!col) return;
                if (editing) {
                  const f = buildAddFieldWrap(col, sampleRow, { idPrefix: "ef_" });
                  const input = f && f.querySelector("input, select, textarea");
                  if (input) setFieldValueFromItem(input, col, item);
                  if (f) grid.appendChild(f);
                } else grid.appendChild(buildEthosDetailFieldWrap(col, item));
              });
              if (!grid.childElementCount) return;
              const fs = document.createElement("fieldset");
              fs.className = "add-form-group";
              const leg = document.createElement("legend");
              leg.className = "add-form-group-legend";
              leg.textContent = group.title;
              fs.appendChild(leg);
              fs.appendChild(grid);
              splitWrap.appendChild(fs);
            });
          }
          return splitWrap.childElementCount ? splitWrap : document.createElement("div");
        }

        async function renderEthosMemberDetailView(item, editing) {
          if (!ethosMemberDetailContent || !item) return;
          ethosMemberDetailContent.innerHTML = "";
          setEthosMemberDetailEditMode(!!editing);
          setEthosMemberDetailState("", "");
          if (ethosMemberDetailTitle) ethosMemberDetailTitle.textContent = formatEthosMemberDisplayName(item);
          const fields = buildEthosMemberDetailFields(item, editing);
          if (editing) {
            const form = document.createElement("div");
            form.id = "ethosEditForm";
            form.className = "add-form-stack";
            form.appendChild(fields);
            ethosMemberDetailContent.appendChild(form);
            const cols = normalizedEthosRosterColumns();
            for (let i = 0; i < cols.length; i++) {
              const col = cols[i];
              if (ETHOS_DROPDOWN_COLUMN_KEYS.indexOf(col.key) === -1) continue;
              const sel = form.querySelector("#ef_" + col.key);
              await fillDropdownSelect(sel, col, ethosDetailSession.seg, ethosDetailSession.pw, ethosDetailSession.sampleRow, item);
            }
            applyAllSelectAutosizes(form);
          } else ethosMemberDetailContent.appendChild(fields);
        }

        function syncEthosMemberInSession(updatedItem) {
          if (!updatedItem || updatedItem.Id == null) return;
          if (!Array.isArray(ethosSession.rows)) ethosSession.rows = [];
          const idx = ethosSession.rows.findIndex(function (r) { return r && String(r.Id) === String(updatedItem.Id); });
          if (idx === -1) ethosSession.rows.push(updatedItem);
          else ethosSession.rows[idx] = Object.assign({}, ethosSession.rows[idx], updatedItem);
          ethosSession.rows = sortEthosRowsAlphabetical(ethosSession.rows);
          if (ethosSession.pw && ethosSession.seg) renderEthosRosterTable(ethosSession.rows, ethosSession.meta, ethosSession.pw, ethosSession.seg);
        }

        async function saveEthosMemberDetailEdits() {
          const s = ethosDetailSession;
          if (!s.item || s.item.Id == null || !s.pw || !s.seg) return;
          const cols = normalizedEthosRosterColumns();
          const payload = {};
          cols.forEach(function (col) {
            const el = document.getElementById("ef_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || resolveWriteKey(col, s.sampleRow);
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          try {
            setEthosMemberDetailState("loading", "Saving changes...");
            await spFetch(`/_api/web/${s.seg}/items(${s.item.Id})`, { method: "MERGE", body: payload }, s.pw);
            const updated = mergePayloadIntoItem(s.item, payload);
            updated.Id = s.item.Id;
            ethosDetailSession.item = updated;
            syncEthosMemberInSession(updated);
            await renderEthosMemberDetailView(updated, false);
            setEthosMemberDetailState("ok", "Record saved.");
            window.setTimeout(function () { setEthosMemberDetailState("", ""); }, 2500);
          } catch (e) {
            setEthosMemberDetailState("err", "Save failed: " + (e.message || String(e)).slice(0, 280));
          }
        }

        function setEthosWeaponsBulkPanelVisible(visible) {
          if (ethosWeaponsBulkPanel) ethosWeaponsBulkPanel.hidden = !visible;
          if (visible) {
            if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = true;
            ethosWeaponsCertEditSession.item = null;
          }
        }

        function clearEthosMemberWeaponsSection() {
          if (ethosWeaponsThead) ethosWeaponsThead.innerHTML = "";
          if (ethosWeaponsBody) ethosWeaponsBody.innerHTML = "";
          ethosWeaponsCertEditSession.item = null;
          if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = true;
          setEthosWeaponsBulkPanelVisible(false);
          if (ethosWeaponsWrap) ethosWeaponsWrap.hidden = true;
          if (ethosWeaponsAddForm) ethosWeaponsAddForm.reset();
          if (ethosWeaponsBulkForm) ethosWeaponsBulkForm.reset();
        }

        function setEthosBylawBulkPanelVisible(visible) {
          if (ethosBylawBulkPanel) ethosBylawBulkPanel.hidden = !visible;
          if (visible) {
            if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = true;
            ethosBylawTrainingEditSession.item = null;
          }
        }

        function clearEthosMemberBylawSection() {
          if (ethosBylawThead) ethosBylawThead.innerHTML = "";
          if (ethosBylawBody) ethosBylawBody.innerHTML = "";
          ethosBylawTrainingEditSession.item = null;
          if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = true;
          setEthosBylawBulkPanelVisible(false);
          if (ethosBylawWrap) ethosBylawWrap.hidden = true;
          if (ethosBylawAddForm) ethosBylawAddForm.reset();
          if (ethosBylawBulkForm) ethosBylawBulkForm.reset();
        }

        function clearEthosMemberDetailSection() {
          clearEthosMemberWeaponsSection();
          clearEthosMemberBylawSection();
          setEthosMemberDetailEditMode(false);
        }

        function setEthosWeaponsState(kind, message) {
          if (!ethosWeaponsState) return;
          if (!message) { ethosWeaponsState.hidden = true; ethosWeaponsState.textContent = ""; return; }
          ethosWeaponsState.hidden = false;
          ethosWeaponsState.className = "read-state " + kind;
          ethosWeaponsState.textContent = message;
        }

        function setEthosBylawState(kind, message) {
          if (!ethosBylawState) return;
          if (!message) { ethosBylawState.hidden = true; ethosBylawState.textContent = ""; return; }
          ethosBylawState.hidden = false;
          ethosBylawState.className = "read-state " + kind;
          ethosBylawState.textContent = message;
        }

        function renderEthosWeaponsCertTable(rows) {
          if (ethosWeaponsThead) ethosWeaponsThead.innerHTML = "";
          if (ethosWeaponsBody) ethosWeaponsBody.innerHTML = "";
          const columns = normalizedEthosWeaponsCertColumns();
          if (!ethosWeaponsThead || !ethosWeaponsBody || !columns.length) return;
          const showActions = !!(ethosDetailSession.pw && ethosWeaponsWrap && !ethosWeaponsWrap.hidden);
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
            trHead.appendChild(thAct);
          }
          ethosWeaponsThead.appendChild(trHead);
          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              if (col.computed && col.key === "Status") {
                const status = computeEthosWeaponsCertStatus(item);
                td.textContent = displayCellText(status.text);
                if (status.tone && status.tone !== "unknown") td.className = "weapons-status weapons-status--" + status.tone;
              } else if (isWeaponsCertDateColumn(col)) {
                td.textContent = displayCellText(formatWeaponsCertDisplayDate(valueFromItemByKeys(item, col.tryKeys || [col.key])));
              } else {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                td.textContent = displayCellText(raw !== undefined && raw !== null ? formatCellValue(raw) : "");
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
              requalBtn.addEventListener("click", function () { void openEthosWeaponsCertEditPanel(item); });
              inner.appendChild(requalBtn);
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              delBtn.addEventListener("click", function () {
                const memberId = ethosDetailSession.item && ethosDetailSession.item.Id;
                void deleteEthosWeaponsCertRow(item.Id, ethosDetailSession.pw, ethosWeaponsCertListApiPath(), memberId);
              });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          ethosWeaponsBody.appendChild(frag);
          if (ethosWeaponsEmpty) ethosWeaponsEmpty.hidden = rows.length > 0;
        }

        async function loadEthosMemberWeaponsCertifications(memberId, pw) {
          if (!memberId) return;
          if (ethosWeaponsWrap) ethosWeaponsWrap.hidden = false;
          setEthosWeaponsState("loading", "Loading Weapons Qualifications...");
          const seg = ethosWeaponsCertListApiPath();
          if (!ethosWeaponsCertListTitle() && !ethosWeaponsCertListUsesGuid()) {
            renderEthosWeaponsCertTable([]);
            setEthosWeaponsState("", "");
            return;
          }
          try {
            const rows = await fetchLinkedCertRowsForMember(seg, pw, memberId, {
              personField: ETHOS_WEAPONS_PERSON_FIELD,
              personFieldAlt: ETHOS_WEAPONS_PERSON_FIELD_ALT,
              orderBy: ETHOS_WEAPONS_CERT_ITEMS_ORDERBY,
              session: ethosSession,
              filterFieldKey: "weaponsPersonFilterField",
              filterFieldsKey: "weaponsPersonFilterFields",
              lookupBase: "EthosMember",
            });
            ethosSession.weaponsCertSampleRow = rows.length ? rows[0] : ethosSession.weaponsCertSampleRow;
            ethosSession.weaponsCertRows = rows.slice();
            renderEthosWeaponsCertTable(rows);
            setEthosWeaponsState("", "");
          } catch (e) {
            renderEthosWeaponsCertTable([]);
            setEthosWeaponsState("warn", "Could not load weapons qualifications: " + (e.message || String(e)).slice(0, 200));
          }
        }

        function buildEthosWeaponsCertAddFormFields(options) {
          if (!ethosWeaponsAddFields) return;
          ethosWeaponsAddFields.innerHTML = "";
          const sampleRow = ethosSession.weaponsCertSampleRow;
          ethosWeaponsCertFormColumns().forEach(function (col) {
            const f = buildWeaponsCertFieldWrap(col, sampleRow, Object.assign({ idPrefix: "ewf_" }, options || {}));
            if (f) ethosWeaponsAddFields.appendChild(f);
          });
        }

        async function populateEthosWeaponsCertAddDropdowns() {
          if (!ethosWeaponsAddForm || !ethosDetailSession.pw) return;
          const seg = ethosWeaponsCertListApiPath();
          const pw = ethosDetailSession.pw;
          const sampleRow = ethosSession.weaponsCertSampleRow;
          for (let i = 0; i < normalizedEthosWeaponsCertColumns().length; i++) {
            const col = normalizedEthosWeaponsCertColumns()[i];
            if (ETHOS_WEAPONS_CERT_DROPDOWN_KEYS.indexOf(col.key) === -1) continue;
            const sel = document.getElementById("ewf_" + col.key);
            if (sel && sel.tagName === "SELECT") await fillDropdownSelect(sel, col, seg, pw, sampleRow, null);
          }
          applyAllSelectAutosizes(ethosWeaponsAddForm);
        }

        async function openEthosWeaponsCertAddPanel() {
          if (!ethosDetailSession.item || ethosDetailSession.item.Id == null) return;
          ethosWeaponsCertEditSession.item = null;
          setEthosWeaponsBulkPanelVisible(false);
          if (ethosWeaponsFormTitle) ethosWeaponsFormTitle.textContent = "New Weapons Qualification";
          if (ethosWeaponsSaveBtn) ethosWeaponsSaveBtn.textContent = "Save qualification";
          buildEthosWeaponsCertAddFormFields();
          if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = false;
          await populateEthosWeaponsCertAddDropdowns();
          wireCertQualDateAutoExpiry(ethosWeaponsAddForm, "ewf_", "ethosWeaponsAutoExpiry");
        }

        async function openEthosWeaponsBulkAddPanel() {
          if (!ethosDetailSession.item || ethosDetailSession.item.Id == null) return;
          const pw = ethosDetailSession.pw;
          if (!pw) {
            setEthosWeaponsState("warn", "Member record is still loading. Try again in a moment.");
            return;
          }
          ethosWeaponsCertEditSession.item = null;
          setEthosWeaponsBulkPanelVisible(true);
          setEthosWeaponsState("", "");
          const seg = ethosWeaponsCertListApiPath();
          const sampleRow = ethosSession.weaponsCertSampleRow;
          const columns = ethosWeaponsCertFormColumns();
          buildWeaponsBulkSharedFields(ethosWeaponsBulkSharedFields, "ewb_", sampleRow, columns);
          await populateWeaponsBulkWeaponChecklist(ethosWeaponsBulkItems, seg, pw, sampleRow, columns);
          if (ethosWeaponsBulkForm) {
            delete ethosWeaponsBulkForm.dataset.ethosWeaponsBulkAutoExpiryWired;
            wireCertQualDateAutoExpiry(ethosWeaponsBulkForm, "ewb_", "ethosWeaponsBulkAutoExpiryWired");
          }
          const qualEl = document.getElementById("ewb_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyCertExpirationFromQual("ewb_");
          if (qualEl) qualEl.focus();
        }

        async function submitEthosWeaponsBulkSave() {
          const memberItem = ethosDetailSession.item;
          const pw = ethosDetailSession.pw;
          if (!memberItem || memberItem.Id == null || !pw) return;

          const selected = getSelectedBylawBulkItems(ethosWeaponsBulkItems);
          if (!selected.length) {
            setEthosWeaponsState("err", "Select at least one weapon.");
            return;
          }
          const qualEl = document.getElementById("ewb_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) {
            setEthosWeaponsState("err", "Certification date is required.");
            return;
          }

          const seg = ethosWeaponsCertListApiPath();
          const sampleRow = ethosSession.weaponsCertSampleRow;
          const columns = ethosWeaponsCertFormColumns();
          const sharedPayload = collectWeaponsBulkSharedPayload("ewb_", columns, sampleRow);
          const memberId = parseInt(String(memberItem.Id), 10);
          if (!memberId || isNaN(memberId)) {
            setEthosWeaponsState("err", "Invalid ETHOS member Id.");
            return;
          }
          const personKey = await resolveLinkedMemberPostKey(
            seg,
            pw,
            ethosSession,
            "weaponsPersonFilterField",
            "weaponsPersonPostKey",
            ETHOS_WEAPONS_PERSON_FIELD,
          );
          sharedPayload[personKey] = memberId;

          try {
            if (ethosWeaponsBulkSaveBtn) ethosWeaponsBulkSaveBtn.disabled = true;
            let saved = 0;
            for (let i = 0; i < selected.length; i++) {
              setEthosWeaponsState("loading", "Saving " + (i + 1) + " of " + selected.length + "...");
              const payload = Object.assign({}, sharedPayload);
              payload[selected[i].writeKey] = selected[i].value;
              if (ETHOS_WEAPONS_CERT_SET_TITLE) payload.Title = selected[i].value;
              await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
              saved++;
            }
            setEthosWeaponsBulkPanelVisible(false);
            if (ethosWeaponsBulkForm) ethosWeaponsBulkForm.reset();
            await loadEthosMemberWeaponsCertifications(memberItem.Id, pw);
            setEthosWeaponsState("ok", "Saved " + saved + " qualification(s).");
            window.setTimeout(function () {
              setEthosWeaponsState("", "");
            }, 2500);
          } catch (e) {
            setEthosWeaponsState("err", "Save failed: " + (e.message || String(e)).slice(0, 200));
          } finally {
            if (ethosWeaponsBulkSaveBtn) ethosWeaponsBulkSaveBtn.disabled = false;
          }
        }

        async function openEthosWeaponsCertEditPanel(item) {
          if (!item || item.Id == null) return;
          setEthosWeaponsBulkPanelVisible(false);
          ethosWeaponsCertEditSession.item = item;
          if (ethosWeaponsFormTitle) ethosWeaponsFormTitle.textContent = "Requalify Weapon";
          if (ethosWeaponsSaveBtn) ethosWeaponsSaveBtn.textContent = "Save requalification";
          buildEthosWeaponsCertAddFormFields({ readOnlyKeys: ["Weapon"] });
          if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = false;
          await populateEthosWeaponsCertAddDropdowns();
          ethosWeaponsCertFormColumns().forEach(function (col) {
            const el = document.getElementById("ewf_" + col.key);
            if (!el || !item) return;
            const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
            if (isWeaponsCertDateColumn(col) && col.key !== "QualDate") el.value = isoDateForDateInput(raw);
            else if (col.key !== "QualDate") el.value = formatCellValue(raw);
          });
          const qualEl = document.getElementById("ewf_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyCertExpirationFromQual("ewf_");
          wireCertQualDateAutoExpiry(ethosWeaponsAddForm, "ewf_", "ethosWeaponsAutoExpiry");
        }

        async function deleteEthosWeaponsCertRow(id, pw, seg, memberId) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid) || !confirm("Delete Weapons Qualification Id " + sid + "?")) return;
          try {
            setEthosWeaponsState("loading", "Deleting...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            if (memberId) await loadEthosMemberWeaponsCertifications(memberId, pw);
            if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = true;
            setEthosWeaponsState("ok", "Qualification deleted.");
          } catch (e) {
            setEthosWeaponsState("err", "Delete failed: " + (e.message || String(e)).slice(0, 200));
          }
        }

        async function submitEthosWeaponsCertSave() {
          const memberItem = ethosDetailSession.item;
          const pw = ethosDetailSession.pw;
          if (!memberItem || memberItem.Id == null || !pw) return;
          const editItem = ethosWeaponsCertEditSession.item;
          const isEdit = editItem && editItem.Id != null;
          const seg = ethosWeaponsCertListApiPath();
          const weaponEl = document.getElementById("ewf_Weapon");
          const weaponVal = weaponEl ? String(weaponEl.value || "").trim() : "";
          if (!isEdit && !weaponVal) { setEthosWeaponsState("err", "Weapon is required."); return; }
          const qualEl = document.getElementById("ewf_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) { setEthosWeaponsState("err", "Certification date is required."); return; }
          const payload = {};
          ethosWeaponsCertFormColumns().forEach(function (col) {
            if (isEdit && col.key === "Weapon") return;
            const el = document.getElementById("ewf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || col.key;
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          if (!isEdit) {
            const personKey = await resolveLinkedMemberPostKey(seg, pw, ethosSession, "weaponsPersonFilterField", "weaponsPersonPostKey", ETHOS_WEAPONS_PERSON_FIELD);
            payload[personKey] = parseInt(String(memberItem.Id), 10);
            if (ETHOS_WEAPONS_CERT_SET_TITLE) payload.Title = weaponVal;
          }
          try {
            setEthosWeaponsState("loading", "Saving...");
            if (isEdit) await spFetch(`/_api/web/${seg}/items(${editItem.Id})`, { method: "MERGE", body: payload }, pw);
            else await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            if (ethosWeaponsAddPanel) ethosWeaponsAddPanel.hidden = true;
            if (ethosWeaponsAddForm) ethosWeaponsAddForm.reset();
            await loadEthosMemberWeaponsCertifications(memberItem.Id, pw);
            setEthosWeaponsState("ok", "Saved.");
          } catch (e) {
            setEthosWeaponsState("err", "Save failed: " + (e.message || String(e)).slice(0, 200));
          }
        }

        function renderEthosBylawTrainingTable(rows) {
          if (ethosBylawThead) ethosBylawThead.innerHTML = "";
          if (ethosBylawBody) ethosBylawBody.innerHTML = "";
          const columns = normalizedEthosBylawColumns();
          if (!ethosBylawThead || !ethosBylawBody) return;
          const showActions = !!(ethosDetailSession.pw && ethosBylawWrap && !ethosBylawWrap.hidden);
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
            trHead.appendChild(thAct);
          }
          ethosBylawThead.appendChild(trHead);
          const frag = document.createDocumentFragment();
          rows.forEach(function (item) {
            const tr = document.createElement("tr");
            columns.forEach(function (col) {
              const td = document.createElement("td");
              if (col.computed && col.key === "Status") {
                const status = computeEthosBylawTrainingStatus(item);
                td.textContent = displayCellText(status.text);
                if (status.tone && status.tone !== "unknown") td.className = "cert-status cert-status--" + status.tone;
              } else if (isWeaponsCertDateColumn(col)) {
                td.textContent = displayCellText(formatWeaponsCertDisplayDate(valueFromItemByKeys(item, col.tryKeys || [col.key])));
              } else if (col.key === "Certifier") {
                td.textContent = displayCellText(formatSharePointLookupDisplay(valueFromItemByKeys(item, col.tryKeys || [col.key])));
              } else {
                const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
                td.textContent = displayCellText(raw !== undefined && raw !== null ? formatCellValue(raw) : "");
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
              requalBtn.addEventListener("click", function () { void openEthosBylawTrainingEditPanel(item); });
              inner.appendChild(requalBtn);
              const delBtn = document.createElement("button");
              delBtn.type = "button";
              delBtn.className = "btn-danger";
              delBtn.textContent = "Delete";
              delBtn.addEventListener("click", function () {
                const memberId = ethosDetailSession.item && ethosDetailSession.item.Id;
                void deleteEthosBylawTrainingRow(item.Id, ethosDetailSession.pw, ethosBylawTrainingListApiPath(), memberId);
              });
              inner.appendChild(delBtn);
              tdAct.appendChild(inner);
              tr.appendChild(tdAct);
            }
            frag.appendChild(tr);
          });
          ethosBylawBody.appendChild(frag);
          if (ethosBylawEmpty) ethosBylawEmpty.hidden = rows.length > 0;
        }

        async function loadEthosMemberBylawTraining(memberId, pw) {
          if (!memberId) return;
          if (ethosBylawWrap) ethosBylawWrap.hidden = false;
          setEthosBylawState("loading", "Loading By-Law Training...");
          const seg = ethosBylawTrainingListApiPath();
          if (!ethosBylawListTitle() && !ethosBylawListUsesGuid()) {
            renderEthosBylawTrainingTable([]);
            setEthosBylawState("", "");
            return;
          }
          try {
            const rows = await fetchLinkedCertRowsForMember(seg, pw, memberId, {
              personField: ETHOS_BYLAW_PERSON_FIELD,
              personFieldAlt: ETHOS_BYLAW_PERSON_FIELD_ALT,
              orderBy: ETHOS_BYLAW_TRAINING_ITEMS_ORDERBY,
              session: ethosSession,
              filterFieldKey: "bylawPersonFilterField",
              filterFieldsKey: "bylawPersonFilterFields",
              lookupBase: "EthosMember",
            });
            ethosSession.bylawTrainingSampleRow = rows.length ? rows[0] : ethosSession.bylawTrainingSampleRow;
            ethosSession.bylawTrainingRows = rows.slice();
            renderEthosBylawTrainingTable(rows);
            setEthosBylawState("", "");
          } catch (e) {
            renderEthosBylawTrainingTable([]);
            setEthosBylawState("warn", "Could not load By-Law Training: " + (e.message || String(e)).slice(0, 200));
          }
        }

        function buildEthosBylawAddFormFields(options) {
          if (!ethosBylawAddFields) return;
          ethosBylawAddFields.innerHTML = "";
          const sampleRow = ethosSession.bylawTrainingSampleRow;
          ethosBylawFormColumns().forEach(function (col) {
            const f = buildBylawTrainingFieldWrap(col, sampleRow, Object.assign({ idPrefix: "ebf_" }, options || {}));
            if (f) ethosBylawAddFields.appendChild(f);
          });
        }

        async function populateEthosBylawAddDropdowns() {
          if (!ethosBylawAddForm || !ethosDetailSession.pw) return;
          const seg = ethosBylawTrainingListApiPath();
          const pw = ethosDetailSession.pw;
          const sampleRow = ethosSession.bylawTrainingSampleRow;
          for (let i = 0; i < normalizedEthosBylawColumns().length; i++) {
            const col = normalizedEthosBylawColumns()[i];
            if (ETHOS_BYLAW_TRAINING_DROPDOWN_KEYS.indexOf(col.key) === -1) continue;
            const sel = document.getElementById("ebf_" + col.key);
            if (!sel || sel.tagName !== "SELECT") continue;
            if (col.key === "Certifier") await fillBylawCertifierDropdown(sel, seg, pw, sampleRow);
            else await fillDropdownSelect(sel, col, seg, pw, sampleRow, null);
          }
          applyAllSelectAutosizes(ethosBylawAddForm);
        }

        async function openEthosBylawTrainingAddPanel() {
          if (!ethosDetailSession.item || ethosDetailSession.item.Id == null) return;
          ethosBylawTrainingEditSession.item = null;
          setEthosBylawBulkPanelVisible(false);
          buildEthosBylawAddFormFields();
          if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = false;
          await populateEthosBylawAddDropdowns();
          wireCertQualDateAutoExpiry(ethosBylawAddForm, "ebf_", "ethosBylawAutoExpiry");
        }

        async function openEthosBylawBulkAddPanel() {
          if (!ethosDetailSession.item || ethosDetailSession.item.Id == null) return;
          const pw = ethosDetailSession.pw;
          if (!pw) {
            setEthosBylawState("warn", "Member record is still loading. Try again in a moment.");
            return;
          }
          ethosBylawTrainingEditSession.item = null;
          setEthosBylawBulkPanelVisible(true);
          setEthosBylawState("", "");
          const seg = ethosBylawTrainingListApiPath();
          const sampleRow = ethosSession.bylawTrainingSampleRow;
          const columns = ethosBylawFormColumns();
          buildBylawBulkSharedFields(ethosBylawBulkSharedFields, "ebb_", sampleRow, columns);
          await populateBylawBulkCertifierDropdown("ebb_", seg, pw, sampleRow);
          await populateBylawBulkItemChecklist(ethosBylawBulkItems, seg, pw, sampleRow, columns);
          if (ethosBylawBulkForm) {
            delete ethosBylawBulkForm.dataset.ethosBylawBulkAutoExpiryWired;
            wireCertQualDateAutoExpiry(ethosBylawBulkForm, "ebb_", "ethosBylawBulkAutoExpiryWired");
          }
          const qualEl = document.getElementById("ebb_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyCertExpirationFromQual("ebb_");
          if (ethosBylawBulkForm) applyAllSelectAutosizes(ethosBylawBulkForm);
          if (qualEl) qualEl.focus();
        }

        async function submitEthosBylawBulkSave() {
          const memberItem = ethosDetailSession.item;
          const pw = ethosDetailSession.pw;
          if (!memberItem || memberItem.Id == null || !pw) return;

          const selected = getSelectedBylawBulkItems(ethosBylawBulkItems);
          if (!selected.length) {
            setEthosBylawState("err", "Select at least one training item.");
            return;
          }
          const qualEl = document.getElementById("ebb_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) {
            setEthosBylawState("err", "Certification date is required.");
            return;
          }

          const seg = ethosBylawTrainingListApiPath();
          const sampleRow = ethosSession.bylawTrainingSampleRow;
          const columns = ethosBylawFormColumns();
          const sharedPayload = collectBylawBulkSharedPayload("ebb_", columns, sampleRow);
          const memberId = parseInt(String(memberItem.Id), 10);
          if (!memberId || isNaN(memberId)) {
            setEthosBylawState("err", "Invalid ETHOS member Id.");
            return;
          }
          const personKey = await resolveLinkedMemberPostKey(
            seg,
            pw,
            ethosSession,
            "bylawPersonFilterField",
            "bylawPersonPostKey",
            ETHOS_BYLAW_PERSON_FIELD,
          );
          sharedPayload[personKey] = memberId;

          try {
            if (ethosBylawBulkSaveBtn) ethosBylawBulkSaveBtn.disabled = true;
            let saved = 0;
            for (let i = 0; i < selected.length; i++) {
              setEthosBylawState("loading", "Saving " + (i + 1) + " of " + selected.length + "...");
              const payload = Object.assign({}, sharedPayload);
              payload[selected[i].writeKey] = selected[i].value;
              if (ETHOS_BYLAW_TRAINING_SET_TITLE) payload.Title = selected[i].value;
              await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
              saved++;
            }
            setEthosBylawBulkPanelVisible(false);
            if (ethosBylawBulkForm) ethosBylawBulkForm.reset();
            await loadEthosMemberBylawTraining(memberItem.Id, pw);
            setEthosBylawState("ok", "Saved " + saved + " training record(s).");
            window.setTimeout(function () {
              setEthosBylawState("", "");
            }, 2500);
          } catch (e) {
            setEthosBylawState("err", "Save failed: " + (e.message || String(e)).slice(0, 200));
          } finally {
            if (ethosBylawBulkSaveBtn) ethosBylawBulkSaveBtn.disabled = false;
          }
        }

        async function openEthosBylawTrainingEditPanel(item) {
          if (!item || item.Id == null) return;
          ethosBylawTrainingEditSession.item = item;
          setEthosBylawBulkPanelVisible(false);
          buildEthosBylawAddFormFields({ readOnlyKeys: ["Item"] });
          if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = false;
          await populateEthosBylawAddDropdowns();
          ethosBylawFormColumns().forEach(function (col) {
            const el = document.getElementById("ebf_" + col.key);
            if (!el) return;
            const raw = valueFromItemByKeys(item, col.tryKeys || [col.key]);
            if (isWeaponsCertDateColumn(col) && col.key !== "QualDate") el.value = isoDateForDateInput(raw);
            else if (col.key !== "QualDate") {
              if (el.tagName === "SELECT") ensureSelectIncludesValue(el, formatSharePointLookupDisplay(raw));
              el.value = col.key === "Certifier" ? String(item.CertifierId || raw || "") : formatCellValue(raw);
            }
          });
          const qualEl = document.getElementById("ebf_QualDate");
          if (qualEl) qualEl.value = isoDateFromCalendarDate(new Date());
          applyCertExpirationFromQual("ebf_");
          wireCertQualDateAutoExpiry(ethosBylawAddForm, "ebf_", "ethosBylawAutoExpiry");
        }

        async function deleteEthosBylawTrainingRow(id, pw, seg, memberId) {
          const sid = parseInt(String(id), 10);
          if (!sid || isNaN(sid) || !confirm("Delete By-Law Training record Id " + sid + "?")) return;
          try {
            setEthosBylawState("loading", "Deleting...");
            await spFetch(`/_api/web/${seg}/items(${sid})`, { method: "DELETE" }, pw);
            if (memberId) await loadEthosMemberBylawTraining(memberId, pw);
            if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = true;
            setEthosBylawState("ok", "Record deleted.");
          } catch (e) {
            setEthosBylawState("err", "Delete failed: " + (e.message || String(e)).slice(0, 200));
          }
        }

        async function submitEthosBylawTrainingSave() {
          const memberItem = ethosDetailSession.item;
          const pw = ethosDetailSession.pw;
          if (!memberItem || memberItem.Id == null || !pw) return;
          const editItem = ethosBylawTrainingEditSession.item;
          const isEdit = editItem && editItem.Id != null;
          const seg = ethosBylawTrainingListApiPath();
          const itemEl = document.getElementById("ebf_Item");
          const itemVal = itemEl ? String(itemEl.value || "").trim() : "";
          if (!isEdit && !itemVal) { setEthosBylawState("err", "Item is required."); return; }
          const qualEl = document.getElementById("ebf_QualDate");
          if (!qualEl || !String(qualEl.value || "").trim()) { setEthosBylawState("err", "Certification date is required."); return; }
          const payload = {};
          ethosBylawFormColumns().forEach(function (col) {
            if (isEdit && col.key === "Item") return;
            const el = document.getElementById("ebf_" + col.key);
            if (!el) return;
            const writeKey = el.dataset.writeKey || col.key;
            const v = formFieldPayloadValue(el, writeKey);
            if (v === null) return;
            payload[writeKey] = v;
          });
          if (!isEdit) {
            const personKey = await resolveLinkedMemberPostKey(seg, pw, ethosSession, "bylawPersonFilterField", "bylawPersonPostKey", ETHOS_BYLAW_PERSON_FIELD);
            payload[personKey] = parseInt(String(memberItem.Id), 10);
            if (ETHOS_BYLAW_TRAINING_SET_TITLE) payload.Title = itemVal;
          }
          try {
            setEthosBylawState("loading", "Saving...");
            if (isEdit) await spFetch(`/_api/web/${seg}/items(${editItem.Id})`, { method: "MERGE", body: payload }, pw);
            else await spFetch(`/_api/web/${seg}/items`, { method: "POST", body: payload }, pw);
            if (ethosBylawAddPanel) ethosBylawAddPanel.hidden = true;
            if (ethosBylawAddForm) ethosBylawAddForm.reset();
            await loadEthosMemberBylawTraining(memberItem.Id, pw);
            setEthosBylawState("ok", "Saved.");
          } catch (e) {
            setEthosBylawState("err", "Save failed: " + (e.message || String(e)).slice(0, 200));
          }
        }

        async function openEthosMemberDetailRecord(itemId, meta, pw, seg, cachedRows, sampleRow) {
          if (!ethosMemberDetailSection || !ethosMemberDetailContent) return;
          if (ethosSection) ethosSection.hidden = true;
          ethosMemberDetailSection.hidden = false;
          setEthosMemberDetailEditMode(false);
          const rows = Array.isArray(cachedRows) ? cachedRows : ethosSession.rows;
          let item = Array.isArray(rows) && rows.find(function (r) { return r && String(r.Id) === String(itemId); });
          if (!item) {
            setEthosMemberDetailState("loading", "Loading ETHOS member...");
            try {
              item = await spFetch(`/_api/web/${seg}/items(${itemId})`, {}, pw);
            } catch (e) {
              setEthosMemberDetailState("err", "Cannot load member: " + (e.message || String(e)).slice(0, 280));
              return;
            }
          }
          ethosDetailSession = {
            item: item,
            editing: false,
            meta: meta || ethosSession.meta,
            pw: pw || ethosSession.pw,
            seg: seg || ethosSession.seg,
            sampleRow: sampleRow || (rows && rows.length ? rows[0] : item),
          };
          await renderEthosMemberDetailView(item, false);
          await loadEthosMemberWeaponsCertifications(item.Id, ethosDetailSession.pw);
          await loadEthosMemberBylawTraining(item.Id, ethosDetailSession.pw);
        }

        function buildEthosMqlCertTable(rows, columns, statusFn, itemLabelKey) {
          const table = document.createElement("table");
          table.className = "ethos-mql-table";
          const itemCol = columns.find(function (c) { return c.key === itemLabelKey; }) || columns.find(function (c) { return c.key === "Weapon" || c.key === "Item"; }) || columns[0];
          const thead = document.createElement("thead");
          const htr = document.createElement("tr");
          [itemCol ? itemCol.label : "Item", "Qual Date", "Expiration", "Status"].forEach(function (label) {
            const th = document.createElement("th");
            th.textContent = label;
            htr.appendChild(th);
          });
          thead.appendChild(htr);
          table.appendChild(thead);
          const tbody = document.createElement("tbody");
          if (!rows.length) {
            const tr = document.createElement("tr");
            const td = document.createElement("td");
            td.colSpan = 4;
            td.textContent = "No records on file.";
            tr.appendChild(td);
            tbody.appendChild(tr);
          } else {
            rows.forEach(function (row) {
              const tr = document.createElement("tr");
              const itemCol = columns.find(function (c) { return c.key === itemLabelKey || c.key === "Weapon" || c.key === "Item"; }) || columns[0];
              const qualCol = columns.find(function (c) { return c.key === "QualDate"; });
              const expCol = columns.find(function (c) { return c.key === "ExpirationDate" || c.key === "ExpiryDate"; });
              const status = statusFn(row);
              [itemCol, qualCol, expCol].forEach(function (col, idx) {
                const td = document.createElement("td");
                if (!col) { td.textContent = "-"; tr.appendChild(td); return; }
                const raw = valueFromItemByKeys(row, col.tryKeys || [col.key]);
                td.textContent = displayCellText(isWeaponsCertDateColumn(col) ? formatWeaponsCertDisplayDate(raw) : formatCellValue(raw));
                tr.appendChild(td);
              });
              const tdStatus = document.createElement("td");
              tdStatus.textContent = displayCellText(status.text);
              tr.appendChild(tdStatus);
              tbody.appendChild(tr);
            });
          }
          table.appendChild(tbody);
          return table;
        }

        function renderHubExternalLinks() {
          if (!hubExternalLinksBar) return;
          const links = Array.isArray(HUB_EXTERNAL_LINKS) ? HUB_EXTERNAL_LINKS : [];
          hubExternalLinksBar.innerHTML = "";
          if (!links.length) {
            hubExternalLinksBar.hidden = true;
            hubExternalLinksBar.classList.remove("hub-external-links--visible");
            return;
          }
          const frag = document.createDocumentFragment();
          links.forEach(function (entry) {
            if (!entry || !entry.url) return;
            const a = document.createElement("a");
            a.className = "btn-secondary";
            a.href = String(entry.url);
            a.target = "_blank";
            a.rel = "noopener noreferrer";
            a.textContent = String(entry.label || entry.url);
            frag.appendChild(a);
          });
          hubExternalLinksBar.appendChild(frag);
          hubExternalLinksBar.hidden = false;
          hubExternalLinksBar.classList.add("hub-external-links--visible");
        }

        function ethosLinkedMemberIdFromItem(row, personField, personFieldAlt) {
          const candidates = linkedMemberFilterCandidates(personField, personFieldAlt);
          return linkedMemberIdFromItem(row, candidates, "EthosMember");
        }

        function ethosWeaponsPersonIdFromItem(row) {
          return ethosLinkedMemberIdFromItem(row, ETHOS_WEAPONS_PERSON_FIELD, ETHOS_WEAPONS_PERSON_FIELD_ALT);
        }

        function ethosBylawPersonIdFromItem(row) {
          return ethosLinkedMemberIdFromItem(row, ETHOS_BYLAW_PERSON_FIELD, ETHOS_BYLAW_PERSON_FIELD_ALT);
        }

        function ethosMqlDisplayOpts() {
          return {
            displayNameFn: formatEthosMemberDisplayName,
            statusLabelFn: function (person) {
              return ethosItemFieldText(person, "Status") || "Unknown";
            },
            rankFn: function (person) {
              return ethosItemFieldText(person, "Rank") || "-";
            },
            navigateToPerson: false,
            tableLabel: "ETHOS MQL Report",
            printTitle: "ETHOS MQL Report",
            unitLabel: String(ETHOS_SQUADRON_LABEL || AF_REPORT_UNIT_LABEL || "88 SFS"),
          };
        }

        async function buildEthosMqlReportData(pw, allWeapons, allBylaw) {
          const sampleBylaw = allBylaw && allBylaw.length ? allBylaw[0] : null;
          const bylawItemChoices = await fetchMqlBylawItemChoices(
            pw,
            ethosBylawTrainingListApiPath,
            sampleBylaw,
            ethosSession,
            "mqlBylawItemChoices",
            allBylaw,
          );
          const catalog = buildMqlCertCatalog(allWeapons, allBylaw, {
            mode: "standard",
            bylawItemChoices: bylawItemChoices,
          });
          const members = sortEthosRowsAlphabetical((ethosSession.rows || []).slice());
          const rows = buildMqlReportRows(members, allWeapons, allBylaw, {
            weaponsPersonIdFn: ethosWeaponsPersonIdFromItem,
            bylawPersonIdFn: ethosBylawPersonIdFromItem,
          });
          return { rows: rows, catalog: catalog };
        }

        async function showEthosMqlReport() {
          const pw = ethosSession.pw || hubSession.pw;
          if (!pw) {
            setEthosReadState("warn", "Load ETHOS roster first.");
            return;
          }
          if (!ethosSession.rows || !ethosSession.rows.length) {
            setEthosReadState("warn", "No ETHOS members loaded.");
            return;
          }
          try {
            setEthosReadState("loading", "Building ETHOS MQL report...");
            const weaponsSeg = ethosWeaponsCertListApiPath();
            const bylawSeg = ethosBylawTrainingListApiPath();
            const allWeapons = await fetchAllEthosLinkedRows(weaponsSeg, pw, {
              orderBy: ETHOS_WEAPONS_CERT_ITEMS_ORDERBY,
            });
            const allBylaw = await fetchAllEthosLinkedRows(bylawSeg, pw, {
              orderBy: ETHOS_BYLAW_TRAINING_ITEMS_ORDERBY,
            });
            const data = await buildEthosMqlReportData(pw, allWeapons, allBylaw);
            ethosSession.mqlRows = data.rows;
            ethosSession.mqlCatalog = data.catalog;
            if (ethosMqlReportBody) {
              ethosMqlReportBody.innerHTML = "";
              const sigPanel = buildMqlSignatureSettingsPanel();
              sigPanel.id = "ethosMqlSignaturePanel";
              const summary = document.createElement("p");
              summary.className = "reports-hint";
              summary.textContent =
                data.rows.length +
                " ETHOS member(s) | " +
                data.catalog.weapons.length +
                " weapon column(s) | " +
                data.catalog.bylaws.length +
                " By-Law column(s).";
              ethosMqlReportBody.appendChild(sigPanel);
              void populateMqlSignaturePanel(sigPanel);
              ethosMqlReportBody.appendChild(summary);
              ethosMqlReportBody.appendChild(renderMqlReportTable(data.rows, data.catalog, ethosMqlDisplayOpts()));
            }
            if (ethosMqlReportPanel) ethosMqlReportPanel.hidden = false;
            if (ethosMqlPrintBtn) {
              ethosMqlPrintBtn.hidden = false;
              ethosMqlPrintBtn.textContent = "Print signed report";
            }
            if (ethosMqlCloseBtn) ethosMqlCloseBtn.hidden = false;
            setEthosReadState("", "");
          } catch (e) {
            setEthosReadState("err", "Could not build MQL report: " + (e.message || String(e)).slice(0, 220));
          }
        }

        function printEthosMqlReport() {
          const rows = ethosSession.mqlRows;
          const catalog = ethosSession.mqlCatalog;
          if (!rows || !catalog) {
            setEthosReadState("warn", "Open ETHOS MQL Report first.");
            return;
          }
          const pw = ethosSession.pw || hubSession.pw;
          void (async function () {
            if (pw) await loadMqlSignature(pw, false);
            const sigPanel =
              ethosMqlReportBody && ethosMqlReportBody.querySelector(".reports-mql-signature-panel");
            const sig = currentMqlSignatureForPrint(sigPanel || ethosMqlReportBody);
            triggerMqlPrint(
              renderMqlSignedPrintDocument(rows, catalog, Object.assign({}, ethosMqlDisplayOpts(), { signature: sig })),
              { signed: true, signature: sig },
            );
          })();
        }

        async function fetchAllEthosLinkedRows(seg, pw, opts) {
          opts = opts || {};
          const orderByClause = String(opts.orderBy || "").trim();
          const orderByQs = orderByClause ? "&$orderby=" + encodeURIComponent(orderByClause) : "";
          let data = null;
          try {
            data = await spFetch(`/_api/web/${seg}/items?$top=2000` + orderByQs, {}, pw);
          } catch (e) {
            if (/\b400\b/.test(String(e.message || "")) && orderByQs) data = await spFetch(`/_api/web/${seg}/items?$top=2000`, {}, pw);
            else throw e;
          }
          return (data && data.value) || [];
        }

        function groupLinkedRowsByMember(rows, personField, personFieldAlt, lookupBase) {
          const candidates = linkedMemberFilterCandidates(personField, personFieldAlt);
          const map = {};
          (rows || []).forEach(function (row) {
            const mid = linkedMemberIdFromItem(row, candidates, lookupBase);
            if (!mid) return;
            if (!map[mid]) map[mid] = [];
            map[mid].push(row);
          });
          return map;
        }

        function triggerEthosPrint(node) {
          if (!ethosPrintSurface || !node) return;
          ethosPrintSurface.innerHTML = "";
          ethosPrintSurface.appendChild(node);
          ethosPrintSurface.hidden = false;
          const root = document.getElementById("sp-pip-ui");
          if (root) root.classList.add("ethos-print-active");
          window.setTimeout(function () {
            window.print();
            window.setTimeout(function () {
              if (root) root.classList.remove("ethos-print-active");
              ethosPrintSurface.hidden = true;
            }, 500);
          }, 120);
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
          renderHubExternalLinks();
          if (!hubAlternateNavActive()) {
            setHubListViewVisible(true);
          }
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
