<#
.SYNOPSIS
  Clean-start test loop for xll-bridge: no Add-ins dialog, no stale registrations.

.DESCRIPTION
  Opens the .xll directly in Excel (File > Open does the same): it is loaded for that session
  only and is NOT added to Excel's Add-ins list, so there is nothing to delete afterwards.
  Each run first removes what an earlier run left behind and copies the files to a fresh
  folder, so Excel and Windows can never hand you a cached or locked copy.

  What it removes (all per-user, HKCU):
    * the COM registration the add-in writes for its ribbon and its stream (RTD) server
    * the "OPEN" entries in Excel's options that make it look for a missing .xll at startup
  Optional: -ResetResiliency re-enables add-ins Excel disabled after a crash;
            -ResetTrust forgets which workbooks you trusted and the extracted DLL cache.

.EXAMPLE
  .\dev-run.ps1                      # fresh copy of xll-bridge.xll + demo.xlsx next to this script
  .\dev-run.ps1 -Xll C:\dl\xll-bridge.xll -Workbook C:\dl\demo.xlsx -ResetTrust
  .\dev-run.ps1 -Clean               # only clean up, do not start Excel

  If PowerShell says the script "is not digitally signed", run it once as
    powershell -ExecutionPolicy Bypass -File .\dev-run.ps1
  or unblock it first:  Unblock-File .\dev-run.ps1
#>
[CmdletBinding()]
param(
    [string]$Xll = (Join-Path $PSScriptRoot 'xll-bridge.xll'),
    [string]$Workbook,
    [switch]$Clean,
    [switch]$ResetTrust,
    [switch]$ResetResiliency,
    [switch]$TrustFolder,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# --- Excel must not be running: it holds the add-in and its registrations open.
$running = Get-Process excel -ErrorAction SilentlyContinue
if ($running) {
    if (-not $Force) {
        throw 'Excel is running. Save your work and close it, or use -Force to end it.'
    }
    $running | Stop-Process -Force
    Start-Sleep -Seconds 1
}

# --- Remove what earlier runs (or the Add-ins dialog) left behind.
$clsids = '{3C1D5B7A-6E24-4F0B-9A52-7B8D1E4C9F63}', '{9A4C2D18-5E7B-4F36-8C01-D2B6A3E5F974}'
$keys = @(
    'HKCU:\Software\Classes\XllBridge.Ribbon',
    'HKCU:\Software\Classes\XllBridge.Rtd',
    'HKCU:\Software\Microsoft\Office\Excel\Addins\XllBridge.Ribbon'
) + ($clsids | ForEach-Object { "HKCU:\Software\Classes\CLSID\$_" })
foreach ($k in $keys) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Host "removed $k" }
}

# Excel remembers add-ins it was asked to open in Options\OPEN, OPEN1, ...
Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue | ForEach-Object {
    $opts = Join-Path $_.PSPath 'Excel\Options'
    if (Test-Path $opts) {
        $props = Get-ItemProperty $opts
        foreach ($p in $props.PSObject.Properties) {
            if ($p.Name -like 'OPEN*' -and "$($p.Value)" -match 'xll-bridge') {
                Remove-ItemProperty $opts -Name $p.Name
                Write-Host "removed Excel option $($p.Name) ($($p.Value))"
            }
        }
    }
    if ($ResetResiliency) {
        $res = Join-Path $_.PSPath 'Excel\Resiliency'
        if (Test-Path $res) { Remove-Item $res -Recurse -Force; Write-Host "removed $res" }
    }
}

if ($ResetTrust) {
    foreach ($p in "$env:APPDATA\xll-bridge\trust.json", "$env:LOCALAPPDATA\xll-bridge\cache") {
        if (Test-Path $p) { Remove-Item $p -Recurse -Force; Write-Host "removed $p" }
    }
}

if ($Clean) { Write-Host 'Clean.'; return }

# --- A fresh folder per run: never a cached or locked copy. Keep the last few for comparison.
if (-not (Test-Path $Xll)) { throw "Cannot find $Xll (use -Xll)" }
if (-not $Workbook) {
    $beside = Join-Path (Split-Path (Resolve-Path $Xll)) 'demo.xlsx'
    if (Test-Path $beside) { $Workbook = $beside }
}
$root = Join-Path $env:LOCALAPPDATA 'xll-bridge-dev'
$stage = Join-Path $root (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item $stage -ItemType Directory | Out-Null
$stagedXll = Join-Path $stage 'xll-bridge.xll'
Copy-Item $Xll $stagedXll
$excelArgs = @("`"$stagedXll`"")
if ($Workbook) {
    # The workbook goes to a fixed folder: Lua trust is tied to the workbook's path, so a new
    # folder every run would ask you to enable the scripts every time.
    $work = Join-Path $root 'work'
    New-Item $work -ItemType Directory -Force | Out-Null
    $stagedBook = Join-Path $work (Split-Path $Workbook -Leaf)
    Copy-Item $Workbook $stagedBook -Force
    Unblock-File $stagedBook
    $excelArgs += "`"$stagedBook`""
}
Get-ChildItem $stage | Unblock-File   # removes the "downloaded from the internet" mark
Get-ChildItem $root -Directory | Where-Object Name -ne 'work' | Sort-Object Name -Descending | Select-Object -Skip 3 |
    ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }

if ($TrustFolder) {
    # Optional: a Trusted Location, so Excel does not ask about the unsigned add-in.
    $ver = Get-ChildItem 'HKCU:\Software\Microsoft\Office' | Where-Object Name -match '\\\d+\.\d+$' |
        Sort-Object Name -Descending | Select-Object -First 1
    if ($ver) {
        $loc = Join-Path $ver.PSPath 'Excel\Security\Trusted Locations\XllBridgeDev'
        New-Item $loc -Force | Out-Null
        Set-ItemProperty $loc -Name Path -Value "$root\"
        Set-ItemProperty $loc -Name AllowSubfolders -Value 1 -Type DWord
        Set-ItemProperty $loc -Name Description -Value 'xll-bridge test copies'
    }
}

Write-Host "Starting Excel with $stagedXll"
Start-Process excel.exe -ArgumentList $excelArgs
Write-Host 'When you are done, just close Excel; the next run (or -Clean) removes the leftovers.'
