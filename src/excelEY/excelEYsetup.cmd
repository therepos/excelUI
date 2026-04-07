@echo off
:: ============================================================
::  Office Add-in Setup — install / update / uninstall
::  Auto-detects Excel / Word / PowerPoint from file extension.
::  Double-click to run. No admin rights needed.
:: ============================================================

set "PS_TEMP=%TEMP%\addin-setup-%RANDOM%.ps1"

:: Use PowerShell to extract everything after ::__PS_BEGIN__ from this file
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$lines = [IO.File]::ReadAllLines('%~f0');" ^
  "$start = -1;" ^
  "for ($i=0; $i -lt $lines.Count; $i++) { if ($lines[$i] -eq '::__PS_BEGIN__') { $start = $i + 1; break } };" ^
  "if ($start -ge 0) { [IO.File]::WriteAllLines('%PS_TEMP%', $lines[$start..($lines.Count-1)]) }"

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
$GithubRepo   = 'therepos/excelUI'
$TagPrefix    = 'excelEY-v'

# ═════════════════════════════════════════════════════════════
# AUTO-DETECT — do not edit below
# ═════════════════════════════════════════════════════════════

$appdata = $env:APPDATA
$ext = [System.IO.Path]::GetExtension($AddinFile).ToLower()

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
        $NeedReg     = $false
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

function Resolve-DownloadUrl {
    # Query GitHub API to find the latest release matching our tag prefix,
    # then return the download URL for our add-in file.
    Write-Host "  Finding latest $AddinName release..."

    $ProgressPreference = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

    $apiUrl = "https://api.github.com/repos/$GithubRepo/releases"
    try {
        $releases = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'addin-setup' } -ErrorAction Stop
    } catch {
        throw "Cannot reach GitHub API: $($_.Exception.Message)"
    }

    # Find the latest release whose tag starts with our prefix
    $matched = $releases | Where-Object { $_.tag_name -like "$TagPrefix*" } | Select-Object -First 1

    if (-not $matched) {
        throw "No release found with tag prefix '$TagPrefix'"
    }

    # Find our file in the release assets
    $asset = $matched.assets | Where-Object { $_.name -eq $AddinFile } | Select-Object -First 1

    if (-not $asset) {
        throw "$AddinFile not found in release $($matched.tag_name)"
    }

    Write-Host "  Found: $($matched.tag_name)" -ForegroundColor Green
    return $asset.browser_download_url
}

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

    try {
        $url = Resolve-DownloadUrl
    } catch {
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"
        exit 1
    }

    Write-Host "  Downloading $AddinFile ..."
    try {
        $ProgressPreference = 'SilentlyContinue'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
        Invoke-WebRequest -Uri $url -OutFile $DestFile -ErrorAction Stop

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
        Add-TrustedLocation
        return
    }

    $ErrorActionPreference = 'SilentlyContinue'

    foreach ($ver in $OfficeVersions) {
        $verBase = "HKCU:\Software\Microsoft\Office\$ver"
        if (-not (Test-Path $verBase)) { continue }

        $optKey = "$verBase\$RegAppKey\Options"
        if (Test-Path $optKey) {
            $props = Get-ItemProperty -Path $optKey
            $existingValues = $props.PSObject.Properties |
                Where-Object { $_.Name -match '^OPEN\d*$' } |
                Select-Object -ExpandProperty Value

            if ($existingValues -contains $DestFile) { continue }

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
    Test-AppRunning
    Invoke-Download
    Invoke-Register
    Write-Host ''
    Write-Host "  Installed. The $AddinName tab will appear next time" -ForegroundColor Green
    Write-Host "  you open $AppType." -ForegroundColor Green
}

Write-Host ''
Read-Host '  Press Enter to exit'