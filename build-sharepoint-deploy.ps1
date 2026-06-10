$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "sharepoint-modern-script-editor-ui.html"
$raw = Get-Content $src -Raw

$cssMatch = [regex]::Match($raw, '(?s)<div id="sp-pip-ui">\s*<style>(.*?)</style>')
if (-not $cssMatch.Success) { throw "Could not extract hub CSS" }
$css = $cssMatch.Groups[1].Value.Trim()

$htmlMatch = [regex]::Match($raw, '(?s)<div id="sp-pip-ui">\s*<style>.*?</style>(.*?)\s*</div>\s*\r?\n\s*<script>')
if (-not $htmlMatch.Success) { throw "Could not extract hub HTML" }
$htmlInner = $htmlMatch.Groups[1].Value.Trim()

$jsStart = $raw.IndexOf('<script>')
$jsEnd = $raw.LastIndexOf('</script>')
if ($jsStart -lt 0 -or $jsEnd -le $jsStart) { throw "Could not locate hub script block" }
$scriptBlock = $raw.Substring($jsStart, $jsEnd - $jsStart)
$open = [regex]::Match($scriptBlock, '\(function \(\) \{')
$close = $scriptBlock.LastIndexOf('})();')
if (-not $open.Success -or $close -lt 0) { throw "Could not parse hub script wrapper" }
$jsBody = $scriptBlock.Substring($open.Index, $close - $open.Index + '})();'.Length)

$assetBase = "/sites/88thSFS/SiteAssets"
$cssFile = "training-hub-styles.txt"
$jsFile = "training-hub-script.txt"

Set-Content -Path (Join-Path $root $cssFile) -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $root $jsFile) -Value $jsBody -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub.css") -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub.js") -Value $jsBody -Encoding UTF8

$paste = @"
<!-- STEP 1: Upload these two files to Site Assets (same library, site root /sites/88thSFS): -->
<!--       training-hub-styles.txt  and  training-hub-script.txt -->
<!--       Download the .txt files from GitHub if .js is blocked on your network. -->
<!-- STEP 2: Paste ONLY this block into Modern Script Editor. -->
<div id="sp-pip-ui">
$htmlInner
</div>
<script>
(function () {
  var base = "$assetBase";
  var cssUrl = base + "/$cssFile";
  var jsUrl = base + "/$jsFile";
  function fail(msg) {
    var root = document.getElementById("sp-pip-ui");
    if (root) {
      root.innerHTML =
        '<p style="font-family:Consolas,monospace;color:#ff5a45;padding:16px;">Training Hub failed to load: ' +
        String(msg || "unknown error") +
        "</p>";
    }
  }
  fetch(cssUrl, { credentials: "same-origin", cache: "no-cache" })
    .then(function (res) {
      if (!res.ok) throw new Error("Upload " + cssUrl + " to Site Assets.");
      return res.text();
    })
    .then(function (css) {
      var tag = document.createElement("style");
      tag.textContent = css;
      document.head.appendChild(tag);
      return fetch(jsUrl, { credentials: "same-origin", cache: "no-cache" });
    })
    .then(function (res) {
      if (!res.ok) throw new Error("Upload " + jsUrl + " to Site Assets.");
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
</script>
"@

Set-Content -Path (Join-Path $root "sharepoint-script-editor-paste.html") -Value $paste -Encoding UTF8

Write-Host "Wrote $cssFile ($((Get-Item (Join-Path $root $cssFile)).Length) bytes)"
Write-Host "Wrote $jsFile ($((Get-Item (Join-Path $root $jsFile)).Length) bytes)"
Write-Host "Wrote sharepoint-script-editor-paste.html ($((Get-Item (Join-Path $root 'sharepoint-script-editor-paste.html')).Length) bytes)"
