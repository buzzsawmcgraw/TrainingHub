$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "sharepoint-modern-script-editor-ui.html"
$raw = Get-Content $src -Raw -Encoding UTF8

function ConvertTo-SharePointAscii([string]$text) {
  if ([string]::IsNullOrEmpty($text)) { return $text }
  $text = $text.Replace([string][char]0x2026, '...')
  $text = $text.Replace([string][char]0x2014, '-')
  $text = $text.Replace([string][char]0x2013, '-')
  $text = $text.Replace([string][char]0x2192, '->')
  $text = $text.Replace([string][char]0x00B7, '*')
  $text = $text.Replace([string][char]0x2022, '*')
  $text = $text.Replace([string][char]0x2019, "'")
  $text = $text.Replace([string][char]0x2018, "'")
  $text = $text.Replace([string][char]0x201C, '"')
  $text = $text.Replace([string][char]0x201D, '"')
  return $text
}

$cssMatch = [regex]::Match($raw, '(?s)<style>(.*?)</style>')
if (-not $cssMatch.Success) { throw "Could not extract CSS from head" }
$css = ConvertTo-SharePointAscii($cssMatch.Groups[1].Value.Trim())
# Strip local-preview body shell only (kept in source for opening the .html file in a browser).
$css = [regex]::Replace($css, '(?s)/\* Page shell when viewing this file directly.*?\*/\s*body\s*\{[^}]*\}\s*', '')

$hubStart = $raw.IndexOf('<div id="sp-pip-ui">')
$scriptIdx = $raw.IndexOf('<script>', $hubStart)
if ($hubStart -lt 0 -or $scriptIdx -le $hubStart) { throw "Could not locate #sp-pip-ui block" }
$hubHtml = ConvertTo-SharePointAscii($raw.Substring($hubStart, $scriptIdx - $hubStart).TrimEnd())

$jsStart = $raw.LastIndexOf('<script>')
$jsEnd = $raw.LastIndexOf('</script>')
if ($jsStart -lt 0 -or $jsEnd -le $jsStart) { throw "Could not locate script block" }
$scriptBlock = $raw.Substring($jsStart, $jsEnd - $jsStart)
$open = [regex]::Match($scriptBlock, '\(function \(\) \{')
$close = $scriptBlock.LastIndexOf('})();')
if (-not $open.Success -or $close -lt 0) { throw "Could not parse script wrapper" }
$js = ConvertTo-SharePointAscii($scriptBlock.Substring($open.Index, $close - $open.Index + '})();'.Length))

$buildIdMatch = [regex]::Match($js, 'const HUB_BUILD_ID = "([^"]*)"')
$hubBuildId = if ($buildIdMatch.Success) { $buildIdMatch.Groups[1].Value } else { "dev" }

$archive = Join-Path $root "archive"
if (-not (Test-Path $archive)) {
  New-Item -ItemType Directory -Path $archive | Out-Null
}

Set-Content -Path (Join-Path $root "training-hub-styles.txt") -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub-script.txt") -Value $js -Encoding UTF8

$paste = @"
<!-- Training Hub - full paste (all-in-one). Use only if Site Assets loading is not an option. -->
<!-- Source: sharepoint-modern-script-editor-ui.html - run build-sharepoint-deploy.ps1 after edits. -->
<!-- Recommended: sharepoint-script-editor-loader.html + Site Assets .txt files instead. -->
<style>
$css
</style>
$hubHtml
<script>
$js
</script>
"@

Set-Content -Path (Join-Path $archive "sharepoint-script-editor-paste.html") -Value $paste -Encoding UTF8
Set-Content -Path (Join-Path $archive "training-hub.css") -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $archive "training-hub.js") -Value $js -Encoding UTF8

$loaderCss = @'
      #sp-pip-ui {
        box-sizing: border-box;
        max-width: 100%;
        min-height: 120px;
        font-family: Consolas, Monaco, "Courier New", monospace;
        font-size: 14px;
        color: #3cff7a;
        background: #050807;
        border: 2px solid #1e4a36;
        border-radius: 8px;
        overflow: hidden;
        position: relative;
      }
      #sp-pip-ui .hub-loader-status {
        margin: 0;
        padding: 20px 24px;
        font-size: 12px;
        color: rgba(210, 255, 224, 0.88);
      }
      #sp-pip-ui .hub-loader-status[hidden] {
        display: none !important;
      }
      #sp-pip-ui .hub-load-err {
        margin: 0;
        padding: 16px;
        color: #ff5a45;
        font-size: 12px;
        line-height: 1.5;
      }
      #sp-pip-ui .hub-build-badge {
        margin: 8px 0 0;
        padding: 4px 8px;
        display: inline-block;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        color: #0a140c;
        background: #c9a227;
        border: 1px solid #e8c547;
        border-radius: 4px;
      }
      #sp-pip-ui .hub-build-badge #hubBuildIdLabel {
        color: #0a140c;
      }
      #sp-pip-ui .hub-access-gate {
        background: #050807 !important;
      }
      #sp-pip-ui:has(#hubAccessGate:not([hidden])) > .hub-top-banner-wrap,
      #sp-pip-ui:has(#hubAccessGate:not([hidden])) > .hub-header,
      #sp-pip-ui:has(#hubAccessGate:not([hidden])) > .hub-main {
        visibility: hidden !important;
        pointer-events: none !important;
      }
      #sp-pip-ui #hubAccessInput {
        color: rgba(210, 255, 224, 0.95) !important;
        -webkit-text-fill-color: rgba(210, 255, 224, 0.95) !important;
        background: #081610 !important;
      }
      #sp-pip-ui button.btn-danger {
        color: #1a0505 !important;
        background: #ff3b35 !important;
        border: 2px solid #9c1818 !important;
      }
      #sp-pip-ui button.btn-record {
        color: #3cff7a !important;
        background: rgba(8, 24, 14, 0.95) !important;
        border: 1px solid #1f8f47 !important;
      }
      #sp-pip-ui button.reports-hub-card {
        color: rgba(210, 255, 224, 0.92) !important;
        background: rgba(5, 14, 10, 0.95) !important;
        border: 1px solid #1e4a36 !important;
      }
      #sp-pip-ui button.reports-name-link {
        color: #3cff7a !important;
        background: transparent !important;
        border: none !important;
      }
      #sp-pip-ui #hubAccessSubmit,
      #sp-pip-ui #probeRun,
      #sp-pip-ui #ethosProbeRun,
      #sp-pip-ui #phase1RefreshBtn,
      #sp-pip-ui #personDetailSaveBtn,
      #sp-pip-ui #ethosMemberDetailSaveBtn,
      #sp-pip-ui #schedulingMemoLetterheadSaveBtn,
      #sp-pip-ui #reportsSotManualSave,
      #sp-pip-ui #reportsSotUpcomingSave,
      #sp-pip-ui button.mql-sig-save-btn,
      #sp-pip-ui form#newPersonForm button[type="submit"],
      #sp-pip-ui .add-form-actions > button[type="submit"] {
        display: inline-block !important;
        visibility: visible !important;
        opacity: 1 !important;
        min-width: 8em !important;
        min-height: 2.25em !important;
        width: auto !important;
        height: auto !important;
        color: #021208 !important;
        background: linear-gradient(180deg, #5dff8f 0%, #3cff7a 100%) !important;
        border: 1px solid #1f8f47 !important;
        border-radius: 4px !important;
        box-shadow: 0 0 14px rgba(60, 255, 122, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.12) !important;
        font-family: Consolas, Monaco, "Courier New", monospace !important;
        font-size: 12px !important;
        font-weight: 700 !important;
        letter-spacing: 0.1em !important;
        text-transform: uppercase !important;
        cursor: pointer !important;
        -webkit-appearance: none !important;
        appearance: none !important;
      }
'@

$loaderJs = @'
(function () {
  var SITE_ASSETS_BASE = "/sites/88thSFS/SiteAssets";
  var LOADER_BUILD_ID = "LOADER_BUILD_PLACEHOLDER";
  var cssUrl = SITE_ASSETS_BASE + "/training-hub-styles.txt";
  var jsUrl = SITE_ASSETS_BASE + "/training-hub-script.txt";
  var statusEl = document.getElementById("hubLoaderStatus");
  var hubScriptCode = "";

  function ensureBuildBadge(buildId) {
    var label = document.getElementById("hubBuildIdLabel");
    if (!label) {
      var titles = document.querySelector("#sp-pip-ui .hub-header-titles");
      if (titles) {
        var badge = document.createElement("p");
        badge.className = "hub-build-badge";
        badge.title = "Hub script version - verify after Site Assets upload";
        badge.innerHTML = 'Build <span id="hubBuildIdLabel"></span>';
        titles.appendChild(badge);
        label = document.getElementById("hubBuildIdLabel");
      }
    }
    if (label) label.textContent = String(buildId || LOADER_BUILD_ID || "unknown");
  }

  ensureBuildBadge(LOADER_BUILD_ID);
  if (statusEl) {
    statusEl.textContent = "Loading Training Hub (build " + LOADER_BUILD_ID + ")...";
  }

  function fail(msg) {
    if (statusEl) statusEl.hidden = true;
    var root = document.getElementById("sp-pip-ui");
    if (!root) return;
    var p = document.createElement("p");
    p.className = "hub-load-err";
    p.textContent = "Training Hub failed to load: " + String(msg || "unknown error");
    root.appendChild(p);
  }

  function parseScriptConst(code, name, fallback) {
    var re = new RegExp("const " + name + '\\s*=\\s*"([^"]*)"');
    var m = code.match(re);
    return m ? m[1] : fallback;
  }

  function isHubAccessGranted(storageKey) {
    try {
      return sessionStorage.getItem(storageKey) === "1";
    } catch (_) {
      return false;
    }
  }

  function grantHubAccess(storageKey) {
    try {
      sessionStorage.setItem(storageKey, "1");
    } catch (_) {}
  }

  function runHubScript(code) {
    window.__trainingHubSkipAccessGate = true;
    try {
      var blob = new Blob([code], { type: "application/javascript" });
      var blobUrl = URL.createObjectURL(blob);
      var s = document.createElement("script");
      s.src = blobUrl;
      s.onload = function () {
        URL.revokeObjectURL(blobUrl);
      };
      s.onerror = function () {
        URL.revokeObjectURL(blobUrl);
        try {
          (new Function(code))();
        } catch (err) {
          fail("Could not run hub script. Re-upload training-hub-script.txt to Site Assets.");
        }
      };
      document.body.appendChild(s);
    } catch (err) {
      try {
        (new Function(code))();
      } catch (err2) {
        fail("Could not run hub script. Re-upload training-hub-script.txt to Site Assets.");
      }
    }
  }

  function initPasswordGateThenRun(password, storageKey, onReady) {
    var gate = document.getElementById("hubAccessGate");
    var input = document.getElementById("hubAccessInput");
    var submit = document.getElementById("hubAccessSubmit");
    var err = document.getElementById("hubAccessError");
    password = String(password || "").trim();

    if (!password || isHubAccessGranted(storageKey)) {
      if (gate) gate.hidden = true;
      onReady();
      return;
    }

    function tryUnlock() {
      if (!input) return;
      if (String(input.value || "") === password) {
        grantHubAccess(storageKey);
        if (err) err.hidden = true;
        if (gate) gate.hidden = true;
        onReady();
        return;
      }
      if (err) {
        err.hidden = false;
        err.textContent = "Incorrect password.";
      }
      input.value = "";
      input.focus();
    }

    if (gate) gate.hidden = false;
    if (input) {
      input.value = "";
      window.setTimeout(function () {
        input.focus();
      }, 0);
      input.addEventListener("keydown", function (ev) {
        if (ev.key === "Enter") {
          ev.preventDefault();
          tryUnlock();
        }
      });
    }
    if (submit) submit.addEventListener("click", tryUnlock);
  }

  fetch(cssUrl, { credentials: "same-origin", cache: "no-cache" })
    .then(function (res) {
      if (!res.ok) throw new Error("Upload training-hub-styles.txt to Site Assets (" + cssUrl + ").");
      return res.text();
    })
    .then(function (css) {
      var tag = document.createElement("style");
      tag.textContent = css;
      document.head.appendChild(tag);
      return fetch(jsUrl, { credentials: "same-origin", cache: "no-cache" });
    })
    .then(function (res) {
      if (!res.ok) throw new Error("Upload training-hub-script.txt to Site Assets (" + jsUrl + ").");
      return res.text();
    })
    .then(function (code) {
      hubScriptCode = code;
      var scriptBuild = parseScriptConst(code, "HUB_BUILD_ID", "");
      if (scriptBuild) {
        ensureBuildBadge(scriptBuild);
      } else {
        ensureBuildBadge(LOADER_BUILD_ID);
      }
      if (statusEl) statusEl.hidden = true;
      var password = parseScriptConst(code, "HUB_ACCESS_PASSWORD", "Training2026");
      var storageKey = parseScriptConst(code, "HUB_ACCESS_STORAGE_KEY", "trainingHubAccessGranted");
      initPasswordGateThenRun(password, storageKey, function () {
        runHubScript(hubScriptCode);
      });
    })
    .catch(function (err) {
      fail(err && err.message ? err.message : err);
    });
})();
'@

$loaderJs = $loaderJs.Replace(
  'var cssUrl = SITE_ASSETS_BASE + "/training-hub-styles.txt";',
  "var cssUrl = SITE_ASSETS_BASE + ""/training-hub-styles.txt?v=$hubBuildId"";"
)
$loaderJs = $loaderJs.Replace(
  'var jsUrl = SITE_ASSETS_BASE + "/training-hub-script.txt";',
  "var jsUrl = SITE_ASSETS_BASE + ""/training-hub-script.txt?v=$hubBuildId"";"
)
$loaderJs = $loaderJs.Replace("LOADER_BUILD_PLACEHOLDER", $hubBuildId)

$hubHtml = $hubHtml -replace 'id="hubBuildIdLabel">pending', ('id="hubBuildIdLabel">' + $hubBuildId)

$hubHtmlLoader = [regex]::Replace(
  $hubHtml,
  '(<div id="sp-pip-ui">)',
  '$1' + "`n      <p id=""hubLoaderStatus"" class=""hub-loader-status"">Loading Training Hub...</p>",
  1
)

$loader = @"
<!-- Training Hub - SMALL loader for Modern Script Editor (recommended). -->
<!-- STEP 1: Upload training-hub-styles.txt and training-hub-script.txt to Site Assets on this site. -->
<!-- STEP 2: Paste this entire file into Modern Script Editor. Change SITE_ASSETS_BASE below if needed. -->
<style>
$loaderCss
</style>
$hubHtmlLoader
<script>
$loaderJs
</script>
"@

Set-Content -Path (Join-Path $root "sharepoint-script-editor-loader.html") -Value $loader -Encoding UTF8

$legacyRootFiles = @(
  "sharepoint-script-editor-paste.html",
  "training-hub.css",
  "training-hub.js"
)
foreach ($name in $legacyRootFiles) {
  $legacyPath = Join-Path $root $name
  if (Test-Path $legacyPath) {
    Remove-Item $legacyPath -Force
  }
}

$loaderBytes = (Get-Item (Join-Path $root "sharepoint-script-editor-loader.html")).Length
$stylesBytes = (Get-Item (Join-Path $root "training-hub-styles.txt")).Length
$scriptBytes = (Get-Item (Join-Path $root "training-hub-script.txt")).Length
$pasteBytes = (Get-Item (Join-Path $archive "sharepoint-script-editor-paste.html")).Length
Write-Host "Wrote sharepoint-script-editor-loader.html ($loaderBytes bytes, paste this + Site Assets)"
Write-Host "Wrote training-hub-styles.txt ($stylesBytes bytes) and training-hub-script.txt ($scriptBytes bytes)"
Write-Host "Hub build id: $hubBuildId (shown in Diagnostics after deploy)"
Write-Host "Wrote archive/sharepoint-script-editor-paste.html ($pasteBytes bytes, all-in-one fallback)"
Write-Host "Wrote archive/training-hub.css and archive/training-hub.js (local preview helpers)"
