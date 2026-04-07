@echo off
setlocal EnableDelayedExpansion
:: ============================================================
::  Office Add-in Setup — install / update / uninstall
::  Auto-detects Excel / Word / PowerPoint from file extension.
::  Double-click to run. No admin rights needed.
:: ============================================================

set "PS_TEMP=%TEMP%\addin-setup-%RANDOM%.ps1"

set "FOUND="
(
    for /f "usebackq delims=" %%L in ("%~f0") do (
        if defined FOUND echo(%%L
        if "%%L"=="::__PS_BEGIN__" set "FOUND=1"
    )
) > "%PS_TEMP%"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_TEMP%"
set "RC=%ERRORLEVEL%"

del /f /q "%PS_TEMP%" >nul 2>&1
exit /b %RC%

::__PS_BEGIN__
# =============================================================
# Office Add-in Setup - PowerShell
# =============================================================

$ErrorActionPreference = 'Stop'

# ═════════════════════════════════════════════════════════════
# CONFIG — change these for each add-in
# ═════════════════════════════════════════════════════════════

$AddinFile    = 'excelEY.xlam'
$AddinName    = 'excelEY'
$DownloadUrl  = 'https://github.com/therepos/excelUI/releases/latest/download/excelEY.xlam'

# ═════════════════════════════════════════════════════════════
# AUTO-DETECT — do not edit below
# ═════════════════════════════════════════════════════════════

$appdata = $env:APPDATA
$ext = [System.IO.Path]::GetExtension($AddinFile).ToLower()

# Determine app type, install path, process name, and registry app key
switch ($ext) {
    { $_ -in '.xlam', '.xla' } {
        $AppType     = 'Excel'
        $DestDir     = "$appdata\Microsoft\AddIns"
        $ProcessName = 'EXCEL'
        $RegAppKey   = 'Excel'
        $NeedReg     = $true
    }
    { $_ -in '.dotm', '.dot' } {
        $AppType     = 'Word'
        $DestDir     = "$appdata\Microsoft\Word\STARTUP"
        $ProcessName = 'WINWORD'
        $RegAppKey   = 'Word'
        $NeedReg     = $false   # Word auto-loads from STARTUP
    }
    { $_ -in '.ppam', '.ppa' } {
        $AppType     = 'PowerPoint'
        $DestDir     = "$appdata\Microsoft\PowerPoint\AddIns"
        $ProcessName = 'POWERPNT'
        $RegAppKey   = 'PowerPoint'
        $NeedReg     = $true
    }
    default {
        Write-Host "  ERROR: Unsupported file type: $ext" -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        exit 1
    }
}

$DestFile = Join-Path $DestDir $AddinFile
$OfficeVersions = @('16.0','15.0','14.0')
$Host.UI.RawUI.WindowTitle = "$AddinName Setup"

# =============================================================
# FUNCTIONS
# =============================================================

function Test-AppRunning {
    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($proc) {
        Write-Host "  $AppType is running. Please close it first." -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"
        exit 1
    }
}

function Invoke-Download {
    if (-not (Test-Path $DestDir)) {
        New-Item -Path $DestDir -ItemType Directory -Force | Out-Null
    }

    Write-Host "  Downloading $AddinFile ..."
    try {
        $ProgressPreference = 'SilentlyContinue'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $DestFile -ErrorAction Stop

        if (-not (Test-Path $DestFile)) { throw 'File not created' }
        Write-Host "  Downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "  Download failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Check your internet connection." -ForegroundColor Yellow
        Read-Host "`n  Press Enter to exit"
        exit 1
    }
}

function Invoke-Register {
    if (-not $NeedReg) {
        Write-Host "  $AddinFile will auto-load from $AppType STARTUP folder." -ForegroundColor Green
        # Still add trusted location
        Add-TrustedLocation
        return
    }

    $ErrorActionPreference = 'SilentlyContinue'

    foreach ($ver in $OfficeVersions) {
        $verBase = "HKCU:\Software\Microsoft\Office\$ver"
        if (-not (Test-Path $verBase)) { continue }

        # Register OPEN value
        $optKey = "$verBase\$RegAppKey\Options"
        if (Test-Path $optKey) {
            $props = Get-ItemProperty -Path $optKey
            $existingValues = $props.PSObject.Properties |
                Where-Object { $_.Name -match '^OPEN\d*$' } |
                Select-Object -ExpandProperty Value

            # Skip if already registered
            if ($existingValues -contains $DestFile) { continue }

            # Find next available OPEN name
            $i = 0
            while ($true) {
                $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
                $cur = (Get-ItemProperty -Path $optKey -Name $name -ErrorAction SilentlyContinue).$name
                if (-not $cur) {
                    New-ItemProperty -Path $optKey -Name $name -Value $DestFile -PropertyType String -Force | Out-Null
                    break
                }
                $i++
            }
        }
    }

    Add-TrustedLocation
    Write-Host "  Registered add-in for auto-load." -ForegroundColor Green
    $ErrorActionPreference = 'Stop'
}

function Add-TrustedLocation {
    $ErrorActionPreference = 'SilentlyContinue'
    $tlPath = $DestDir.TrimEnd('\') + '\'

    foreach ($ver in $OfficeVersions) {
        $tlBase = "HKCU:\Software\Microsoft\Office\$ver\$RegAppKey\Security\Trusted Locations"
        if (-not (Test-Path $tlBase)) { continue }

        # Check if already trusted
        $found = $false
        Get-ChildItem $tlBase | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($p.Path -and ($p.Path.TrimEnd('\') + '\') -ieq $tlPath) { $found = $true }
        }

        if (-not $found) {
            $n = 1
            while (Test-Path "$tlBase\Location$n") { $n++ }
            $k = "$tlBase\Location$n"
            New-Item -Path $k -Force | Out-Null
            New-ItemProperty -Path $k -Name Path -Value $tlPath -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $k -Name AllowSubFolders -Value 1 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $k -Name Description -Value "$AddinName Add-in" -PropertyType String -Force | Out-Null
        }
    }
    $ErrorActionPreference = 'Stop'
}

function Invoke-Deregister {
    $ErrorActionPreference = 'SilentlyContinue'

    foreach ($ver in $OfficeVersions) {
        $verBase = "HKCU:\Software\Microsoft\Office\$ver"
        if (-not (Test-Path $verBase)) { continue }

        # Remove OPEN entries
        if ($NeedReg) {
            $optKey = "$verBase\$RegAppKey\Options"
            if (Test-Path $optKey) {
                $props = Get-ItemProperty -Path $optKey
                foreach ($p in $props.PSObject.Properties) {
                    if ($p.Name -match '^OPEN\d*$' -and $p.Value -eq $DestFile) {
                        Remove-ItemProperty -Path $optKey -Name $p.Name -ErrorAction SilentlyContinue
                    }
                }
            }
        }

        # Remove trusted location
        $tlBase = "$verBase\$RegAppKey\Security\Trusted Locations"
        if (Test-Path $tlBase) {
            $tlPath = $DestDir.TrimEnd('\') + '\'
            Get-ChildItem $tlBase | ForEach-Object {
                $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                if ($p.Path -and ($p.Path.TrimEnd('\') + '\') -ieq $tlPath -and $p.Description -like "$AddinName*") {
                    Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    Write-Host "  Deregistered add-in." -ForegroundColor Green
    $ErrorActionPreference = 'Stop'
}

# =============================================================
# MAIN
# =============================================================

Write-Host ''
Write-Host '  ========================================' -ForegroundColor Cyan
Write-Host "    $AddinName - $AppType Add-in Setup" -ForegroundColor Cyan
Write-Host '  ========================================' -ForegroundColor Cyan
Write-Host ''

if (Test-Path $DestFile) {
    # Already installed
    Write-Host "  $AddinName is already installed at:"
    Write-Host "    $DestFile"
    Write-Host ''
    Write-Host '  [U] Update    - download latest version'
    Write-Host '  [R] Uninstall - remove add-in'
    Write-Host '  [C] Cancel'
    Write-Host ''

    $ans = Read-Host '  Choose (U/R/C)'

    switch ($ans.Trim().ToUpper()) {
        'U' {
            Test-AppRunning
            Invoke-Download
            Invoke-Register
            Write-Host ''
            Write-Host "  Updated. Restart $AppType to load the new version." -ForegroundColor Green
        }
        'R' {
            Test-AppRunning
            Invoke-Deregister
            try {
                Remove-Item -Path $DestFile -Force -ErrorAction Stop
                Write-Host "  $AddinName uninstalled." -ForegroundColor Green
            } catch {
                Write-Host "  ERROR: Could not remove $AddinFile." -ForegroundColor Red
            }
        }
        default {
            Write-Host '  Cancelled.'
        }
    }
} else {
    # Fresh install
    Test-AppRunning
    Invoke-Download
    Invoke-Register
    Write-Host ''
    Write-Host "  Installed. The $AddinName tab will appear next time" -ForegroundColor Green
    Write-Host "  you open $AppType." -ForegroundColor Green
}

Write-Host ''
Read-Host '  Press Enter to exit'