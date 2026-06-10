$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "sharepoint-modern-script-editor-ui.html"
$raw = Get-Content $src -Raw

$cssMatch = [regex]::Match($raw, '(?s)<style>(.*?)</style>')
if (-not $cssMatch.Success) { throw "Could not extract CSS from head" }
$css = $cssMatch.Groups[1].Value.Trim()
# Strip local-preview body shell — must not ship to SharePoint (breaks Script Editor canvas).
$css = [regex]::Replace($css, '(?s)/\* Page shell when viewing this file directly.*?\*/\s*body\s*\{[^}]*\}\s*', '')

$hubStart = $raw.IndexOf('<div id="sp-pip-ui">')
$scriptIdx = $raw.IndexOf('<script>', $hubStart)
if ($hubStart -lt 0 -or $scriptIdx -le $hubStart) { throw "Could not locate #sp-pip-ui block" }
$hubHtml = $raw.Substring($hubStart, $scriptIdx - $hubStart).TrimEnd()

$jsStart = $raw.LastIndexOf('<script>')
$jsEnd = $raw.LastIndexOf('</script>')
if ($jsStart -lt 0 -or $jsEnd -le $jsStart) { throw "Could not locate script block" }
$scriptBlock = $raw.Substring($jsStart, $jsEnd - $jsStart)
$open = [regex]::Match($scriptBlock, '\(function \(\) \{')
$close = $scriptBlock.LastIndexOf('})();')
if (-not $open.Success -or $close -lt 0) { throw "Could not parse script wrapper" }
$js = $scriptBlock.Substring($open.Index, $close - $open.Index + '})();'.Length)

Set-Content -Path (Join-Path $root "training-hub.css") -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub.js") -Value $js -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub-styles.txt") -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub-script.txt") -Value $js -Encoding UTF8

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
      }
      #sp-pip-ui .hub-authoring-placeholder {
        margin: 0;
        padding: 22px 24px 24px;
        font-size: 12px;
        line-height: 1.55;
        color: rgba(210, 255, 224, 0.88);
        background: rgba(5, 14, 10, 0.55);
      }
      #sp-pip-ui .hub-authoring-placeholder strong {
        display: block;
        margin-bottom: 8px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: #ffcc66;
      }
      #sp-pip-ui.hub-authoring-mode .hub-main,
      #sp-pip-ui.hub-authoring-mode .hub-boot-overlay {
        display: none !important;
      }
      #sp-pip-ui .hub-load-err {
        margin: 0;
        padding: 16px;
        color: #ff5a45;
        font-size: 12px;
        line-height: 1.5;
      }
'@

$loaderJs = @'
(function () {
  var SITE_ASSETS_BASE = "/sites/88thSFS/SiteAssets";
  var cssUrl = SITE_ASSETS_BASE + "/training-hub-styles.txt";
  var jsUrl = SITE_ASSETS_BASE + "/training-hub-script.txt";

  function isSharePointAuthoringContext() {
    var urlPatterns = [/mode=edit/i, /action=edit/i, /layoutview=edit/i, /preview=1/i, /[?&]e=[a-f0-9-]{8,}/i];
    var wins = [window];
    try {
      if (window.parent && window.parent !== window) wins.push(window.parent);
    } catch (e1) {}
    try {
      if (window.top && window.top !== window) wins.push(window.top);
    } catch (e2) {}
    for (var w = 0; w < wins.length; w++) {
      try {
        var loc = wins[w].location;
        var combined = String(loc && loc.href ? loc.href : "") + String(loc && loc.hash ? loc.hash : "");
        for (var p = 0; p < urlPatterns.length; p++) {
          if (urlPatterns[p].test(combined)) return true;
        }
      } catch (e3) {}
    }
    var docs = [document];
    try {
      if (window.parent && window.parent.document) docs.push(window.parent.document);
    } catch (e4) {}
    try {
      if (window.top && window.top.document) docs.push(window.top.document);
    } catch (e5) {}
    for (var d = 0; d < docs.length; d++) {
      var doc = docs[d];
      if (!doc || !doc.querySelector) continue;
      if (doc.querySelector('[data-automation-id="publishButton"]')) return true;
      if (doc.querySelector('[data-automation-id="SaveAndCloseButton"]')) return true;
      if (doc.querySelector('[data-automation-id="pageCommandBar"]')) return true;
      if (doc.querySelector(".CanvasZone--edit")) return true;
    }
    return false;
  }

  function showAuthoringPlaceholder() {
    var root = document.getElementById("sp-pip-ui");
    var boot = document.getElementById("hubBootOverlay");
    var placeholder = document.getElementById("hubAuthoringPlaceholder");
    if (boot) boot.hidden = true;
    if (placeholder) placeholder.hidden = false;
    if (root) root.classList.add("hub-ready", "hub-authoring-mode");
  }

  function fail(msg) {
    var root = document.getElementById("sp-pip-ui");
    if (!root) return;
    root.innerHTML = '<p class="hub-load-err">Training Hub failed to load: ' + String(msg || "unknown error") + "</p>";
  }

  if (isSharePointAuthoringContext()) {
    showAuthoringPlaceholder();
    return;
  }

  fetch(cssUrl, { credentials: "same-origin", cache: "no-cache" })
    .then(function (res) {
      if (!res.ok) throw new Error("Upload " + cssUrl + " to Site Assets (same site as this page).");
      return res.text();
    })
    .then(function (css) {
      var tag = document.createElement("style");
      tag.textContent = css;
      document.head.appendChild(tag);
      return fetch(jsUrl, { credentials: "same-origin", cache: "no-cache" });
    })
    .then(function (res) {
      if (!res.ok) throw new Error("Upload " + jsUrl + " to Site Assets (same site as this page).");
      return res.text();
    })
    .then(function (code) {
      var s = document.createElement("script");
      s.text = code;
      document.body.appendChild(s);
    })
    .catch(function (err) {
      fail(err && err.message ? err.message : err);
    });
})();
'@

$paste = @"
<!-- Training Hub - SMALL paste for Modern Script Editor (keeps the SharePoint page editable). -->
<!-- STEP 1: Upload training-hub-styles.txt and training-hub-script.txt to Site Assets on the SAME site. -->
<!--         Default path: /sites/88thSFS/SiteAssets/ - change SITE_ASSETS_BASE in the loader script if needed. -->
<!-- STEP 2: Paste this entire file into Modern Script Editor, save, publish. -->
<!-- Do NOT paste the old ~130 KB all-in-one file; it can corrupt the page and lock edit mode. -->
<style>
$loaderCss
</style>
$hubHtml
<script>
$loaderJs
</script>
"@

Set-Content -Path (Join-Path $root "sharepoint-script-editor-paste.html") -Value $paste -Encoding UTF8

$pasteBytes = (Get-Item (Join-Path $root "sharepoint-script-editor-paste.html")).Length
Write-Host "Wrote sharepoint-script-editor-paste.html ($pasteBytes bytes) - small loader"
Write-Host "Wrote training-hub-styles.txt and training-hub-script.txt for Site Assets upload"
