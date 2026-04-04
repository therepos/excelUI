@echo off
setlocal EnableExtensions

:: excelUI Setup — one-click install/update/uninstall
:: Downloads excelUI.xlam from GitHub and installs to %APPDATA%\Microsoft\AddIns
:: Registers via HKCU registry so Excel loads it automatically.
:: No admin rights needed.

set "ADDIN=excelUI.xlam"
set "ADDIN_NAME=excelUI"
set "DEST=%APPDATA%\Microsoft\AddIns"
set "DST=%DEST%\%ADDIN%"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/therepos/excelUI/refs/heads/main/src/xlam/excelUI.xlam"

echo.
echo  ========================================
echo    excelUI - Excel Add-in Setup
echo  ========================================
echo.

:: Check if already installed
if exist "%DST%" goto :existing

:: ---- FRESH INSTALL ----
:install

call :checkexcel
if %errorlevel%==1 exit /b 1

echo  Downloading %ADDIN% ...
call :download
if %errorlevel%==1 exit /b 1

call :register

echo.
echo  Installed. The ExcelUI tab will appear next time
echo  you open Excel.
echo.
pause
exit /b 0

:: ---- ALREADY INSTALLED ----
:existing

echo  excelUI is already installed at:
echo    %DST%
echo.
echo  [U] Update    - download latest version
echo  [R] Uninstall - remove add-in
echo  [C] Cancel
echo.
set /p "ANS=  Choose (U/R/C): "
if /i "%ANS%"=="U" goto :update
if /i "%ANS%"=="R" goto :uninstall
echo  Cancelled.
echo.
pause
exit /b 0

:: ---- UPDATE ----
:update

call :checkexcel
if %errorlevel%==1 exit /b 1

echo  Downloading latest %ADDIN% ...
call :download
if %errorlevel%==1 exit /b 1

call :register

echo.
echo  Updated. Restart Excel to load the new version.
echo.
pause
exit /b 0

:: ---- UNINSTALL ----
:uninstall

call :checkexcel
if %errorlevel%==1 exit /b 1

:: Remove OPEN entries from Excel Options registry
call :deregister

:: Remove file
del /f "%DST%" >nul 2>&1

if exist "%DST%" (
    echo  ERROR: Could not remove %ADDIN%.
    echo.
    pause
    exit /b 1
)

echo.
echo  excelUI uninstalled.
echo.
pause
exit /b 0

:: ===========================================================
::  SUBROUTINES
:: ===========================================================

:checkexcel
tasklist /fi "imagename eq EXCEL.EXE" 2>nul | find /i "EXCEL.EXE" >nul
if %errorlevel%==0 (
    echo  Excel is running. Please close it first.
    echo.
    pause
    exit /b 1
)
exit /b 0

:download
if not exist "%DEST%" mkdir "%DEST%"
powershell -NoProfile -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13; Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%DST%'; if(!(Test-Path '%DST%')){exit 1}"
if errorlevel 1 (
    echo  Download failed. Check your internet connection.
    echo.
    pause
    exit /b 1
)
echo  Downloaded successfully.
exit /b 0

:register
:: Excel uses OPEN / OPEN1 / OPEN2 ... entries under Options to auto-load add-ins.
:: Also add the AddIns folder as a Trusted Location.
powershell -NoProfile -Command ^
 "$ErrorActionPreference='SilentlyContinue'; "^
 "$versions = @('16.0','15.0','14.0') | Where-Object { Test-Path \"HKCU:\Software\Microsoft\Office\$_\" }; "^
 "foreach ($ver in $versions) { "^
 "  $optKey = \"HKCU:\Software\Microsoft\Office\$ver\Excel\Options\"; "^
 "  if (Test-Path $optKey) { "^
 "    $props = Get-ItemProperty -Path $optKey; "^
 "    $existing = $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | Select-Object -ExpandProperty Value; "^
 "    if ($existing -contains '%DST%') { continue }; "^
 "    $i = 1; $done = $false; "^
 "    while (-not $done) { "^
 "      $name = if ($i -eq 1) { 'OPEN' } else { 'OPEN' + ($i - 1) }; "^
 "      $cur = (Get-ItemProperty -Path $optKey -Name $name -ErrorAction SilentlyContinue).$name; "^
 "      if (-not $cur) { New-ItemProperty -Path $optKey -Name $name -Value '%DST%' -PropertyType String -Force | Out-Null; $done = $true }; "^
 "      $i++ "^
 "    } "^
 "  }; "^
 "  $tlBase = \"HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\"; "^
 "  if (Test-Path $tlBase) { "^
 "    $tlPath = '%DEST%\'.TrimEnd('\') + '\'; "^
 "    $found = $false; "^
 "    Get-ChildItem $tlBase | ForEach-Object { "^
 "      $p = (Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue); "^
 "      if ($p.Path -and ($p.Path.TrimEnd('\') + '\') -ieq $tlPath) { $found = $true } "^
 "    }; "^
 "    if (-not $found) { "^
 "      $n = 1; while (Test-Path \"$tlBase\Location$n\") { $n++ }; "^
 "      $k = \"$tlBase\Location$n\"; "^
 "      New-Item -Path $k -Force | Out-Null; "^
 "      New-ItemProperty -Path $k -Name Path -Value $tlPath -PropertyType String -Force | Out-Null; "^
 "      New-ItemProperty -Path $k -Name AllowSubFolders -Value 1 -PropertyType DWord -Force | Out-Null; "^
 "      New-ItemProperty -Path $k -Name Description -Value 'excelUI Add-in' -PropertyType String -Force | Out-Null "^
 "    } "^
 "  } "^
 "}"
echo  Registered add-in for auto-load.
exit /b 0

:deregister
powershell -NoProfile -Command ^
 "$ErrorActionPreference='SilentlyContinue'; "^
 "$versions = @('16.0','15.0','14.0') | Where-Object { Test-Path \"HKCU:\Software\Microsoft\Office\$_\" }; "^
 "foreach ($ver in $versions) { "^
 "  $optKey = \"HKCU:\Software\Microsoft\Office\$ver\Excel\Options\"; "^
 "  if (Test-Path $optKey) { "^
 "    $props = Get-ItemProperty -Path $optKey; "^
 "    foreach ($p in $props.PSObject.Properties) { "^
 "      if ($p.Name -match '^OPEN\d*$' -and $p.Value -eq '%DST%') { "^
 "        Remove-ItemProperty -Path $optKey -Name $p.Name -ErrorAction SilentlyContinue "^
 "      } "^
 "    } "^
 "  }; "^
 "  $tlBase = \"HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations\"; "^
 "  if (Test-Path $tlBase) { "^
 "    $tlPath = '%DEST%\'.TrimEnd('\') + '\'; "^
 "    Get-ChildItem $tlBase | ForEach-Object { "^
 "      $p = (Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue); "^
 "      if ($p.Path -and ($p.Path.TrimEnd('\') + '\') -ieq $tlPath -and $p.Description -like 'excelUI*') { "^
 "        Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue "^
 "      } "^
 "    } "^
 "  } "^
 "}"
echo  Deregistered add-in.
exit /b 0