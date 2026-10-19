param(
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$W = 1600,
    [int]$H = 900
)
# Export every slide to PNG through PowerPoint COM.
#
# HARD RULE: keep this file pure ASCII, with NO BOM.
# Windows PowerShell 5.1 (the only powershell.exe on this machine) reads a
# BOM-less script as ANSI/GBK. A UTF-8 Chinese comment then decodes into
# garbage that swallows the following newline, which silently comments out the
# next code line -- the symptom is "$pres is null" plus "slides: 0".
# Add Chinese only as UTF-8 with BOM, or not at all.
$ErrorActionPreference = 'Stop'
if (Test-Path $Out) { Remove-Item $Out -Recurse -Force }
New-Item -ItemType Directory -Force -Path $Out | Out-Null

$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    # Open(path, ReadOnly, Untitled, WithWindow) -- the 3rd arg must be False,
    # otherwise a headless session hangs waiting for a new window.
    $pres = $pp.Presentations.Open([string]$Src, $true, $false, $false)
    if ($null -eq $pres) { throw "PowerPoint returned no presentation for: $Src" }
    "slides: $($pres.Slides.Count)"
    $pres.Export([string]$Out, 'PNG', $W, $H)
    $pres.Close()
} finally {
    if ($pp) { $pp.Quit() }
}
Start-Sleep -Milliseconds 500
$files = Get-ChildItem $Out -Filter '*.PNG'
"exported: $($files.Count) PNG"
