$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "sharepoint-modern-script-editor-ui.html"
$raw = Get-Content $src -Raw

$cssMatch = [regex]::Match($raw, '(?s)<style>(.*?)</style>')
if (-not $cssMatch.Success) { throw "Could not extract CSS from head" }
$css = $cssMatch.Groups[1].Value.Trim()
# Strip local-preview body shell only (kept in source for opening the .html file in a browser).
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

$paste = @"
<!-- Training Hub - paste this entire file into Modern Script Editor (one copy/paste). -->
<!-- Source: sharepoint-modern-script-editor-ui.html - run build-sharepoint-deploy.ps1 after edits. -->
<style>
$css
</style>
$hubHtml
<script>
$js
</script>
"@

Set-Content -Path (Join-Path $root "sharepoint-script-editor-paste.html") -Value $paste -Encoding UTF8

$pasteBytes = (Get-Item (Join-Path $root "sharepoint-script-editor-paste.html")).Length
Write-Host "Wrote sharepoint-script-editor-paste.html ($pasteBytes bytes)"
Write-Host "Wrote training-hub.css / training-hub.js (local preview helpers)"
