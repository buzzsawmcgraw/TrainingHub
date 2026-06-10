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
$js = $jsBody

$assetBase = "/sites/88thSFS/SiteAssets"
Set-Content -Path (Join-Path $root "training-hub.css") -Value $css -Encoding UTF8
Set-Content -Path (Join-Path $root "training-hub.js") -Value $js -Encoding UTF8

$paste = @"
<!-- STEP 1: Upload training-hub.css and training-hub.js to Site Assets on your SharePoint site. -->
<!-- STEP 2: Paste ONLY this block into Modern Script Editor (about 5 KB — do NOT paste training-hub.js). -->
<link rel="stylesheet" href="$assetBase/training-hub.css">
<div id="sp-pip-ui">
$htmlInner
</div>
<script src="$assetBase/training-hub.js"></script>
"@

Set-Content -Path (Join-Path $root "sharepoint-script-editor-paste.html") -Value $paste -Encoding UTF8

Write-Host "Wrote training-hub.css ($((Get-Item (Join-Path $root 'training-hub.css')).Length) bytes)"
Write-Host "Wrote training-hub.js ($((Get-Item (Join-Path $root 'training-hub.js')).Length) bytes)"
Write-Host "Wrote sharepoint-script-editor-paste.html ($((Get-Item (Join-Path $root 'sharepoint-script-editor-paste.html')).Length) bytes)"
