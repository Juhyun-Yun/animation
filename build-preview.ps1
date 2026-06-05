# Build a single previewable HTML from the apps-script parts.
# This script contains NO Korean text on purpose: PowerShell 5.1 reads .ps1
# files as the system ANSI codepage when there is no BOM, which would corrupt
# Korean. All Korean lives in UTF-8 source files that we read explicitly as UTF-8.
#
# Usage:  in PowerShell, run   .\build-preview.ps1
#         then open the generated preview.html in Chrome.

$root = $PSScriptRoot
$utf8 = [System.Text.Encoding]::UTF8
function ReadUtf8($path) { return [System.IO.File]::ReadAllText($path, $utf8) }

$tpl    = ReadUtf8 (Join-Path $root 'preview-template.html')
$style  = ReadUtf8 (Join-Path $root 'apps-script\Style.html')
$app    = ReadUtf8 (Join-Path $root 'apps-script\App.html')
$studio = ReadUtf8 (Join-Path $root 'apps-script\Studio.html')

$out = $tpl.Replace('@@STYLE@@', $style).Replace('@@APP@@', $app).Replace('@@STUDIO@@', $studio)

# Write with a UTF-8 BOM so every browser detects UTF-8 for sure.
$bomUtf8 = New-Object System.Text.UTF8Encoding($true)
$target  = Join-Path $root 'preview.html'
[System.IO.File]::WriteAllText($target, $out, $bomUtf8)

Write-Host ("Built: " + $target)
Write-Host "Open preview.html in Chrome. Teacher PIN: 1234"
