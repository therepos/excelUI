@echo off
setlocal EnableExtensions

:: Sanity: required files must be beside this cmd
if not exist "%~dp0excelUI.xlam" (
  echo Missing excelUI.xlam next to setup.cmd
  pause
  exit /b 1
)

:: Extract embedded PowerShell payload
set "PAYTAG=::PAYLOAD"
for /f "delims=:" %%A in ('findstr /n /c:"%PAYTAG%" "%~f0"') do set /a LN=%%A+1
set "TMPPS=%TEMP%\setup_%RANDOM%.ps1"
more +%LN% "%~f0" > "%TMPPS%"

:: Pass our folder via ENV and PS param; run STA for Excel COM
set "SETUP_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%TMPPS%" -SourceDirOverride "%~dp0"
set "rc=%ERRORLEVEL%"
del "%TMPPS%" >nul 2>&1
if not "%rc%"=="0" (
  echo.
  echo Installer reported an error (code %rc%). See messages above.
  pause
)
endlocal
exit /b %rc%

::PAYLOAD
param([string]$SourceDirOverride)

# ================= SETUP =================
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$Host.UI.RawUI.WindowTitle = 'excelUI setup'
# ========================================

<# 
Excel Add-in Installer/Uninstaller (single file)
- Copies to C:\Apps (creates if missing)
- Adds Excel Trusted Location (HKCU for 16.0/15.0/14.0 if present)
- Loads/Unloads .xlam via COM; falls back to Excel OPEN registry if COM fails
- Aligned status; stops on copy failure
#>

# ===== CONFIG =====
$TargetDir      = 'C:\Apps'
$FilesToInstall = @('excelUI.xlam')
$AddInFile      = 'excelUI.xlam'
$TrustedDesc    = 'Installer-Managed: Apps XL Add-in'
# ==================

# -------- Helpers --------
function Get-ScriptDir {
  if ($env:SETUP_DIR)     { return $env:SETUP_DIR }        # from wrapper ENV (most reliable)
  if ($SourceDirOverride) { return $SourceDirOverride }    # also passed as PS arg
  if ($PSScriptRoot)      { return $PSScriptRoot }
  if ($PSCommandPath)     { return (Split-Path -Parent $PSCommandPath) }
  if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
  return (Get-Location).Path
}
$SourceDir = (Get-ScriptDir).TrimEnd('\') + '\'

function Ensure-Dir($p){ if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } }
function Get-OfficeVersions { @('16.0','15.0','14.0') | Where-Object { Test-Path "HKCU:\Software\Microsoft\Office\$_" } }

function Status($msg,[scriptblock]$act,[switch]$Fatal){
  $w=34; Write-Host ($msg.PadRight($w)) -NoNewline
  try{ & $act | Out-Null; Write-Host "Done" -ForegroundColor Green }
  catch{
    Write-Host "Failed" -ForegroundColor Red
    Write-Host ("  " + $_.Exception.Message) -ForegroundColor DarkRed
    if($Fatal){ throw }
  }
}

# Trusted Locations
function Add-TrustedLocation($path,$desc){
  $ok=$false
  foreach($ver in Get-OfficeVersions){
    $base="HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations"
    if(-not(Test-Path $base)){ continue }
    $n=1; while(Test-Path "$base\Location$n"){ $n++ }
    $k="$base\Location$n"
    New-Item -Path $k -Force | Out-Null
    New-ItemProperty -Path $k -Name Path -Value ($path.TrimEnd('\')+'\') -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $k -Name AllowSubFolders -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $k -Name Description -Value $desc -PropertyType String -Force | Out-Null
    $ok=$true
  }
  return $ok
}
function Remove-TrustedLocation($path,$desc){
  $removed=$false
  foreach($ver in Get-OfficeVersions){
    $base="HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations"
    if(-not(Test-Path $base)){ continue }
    Get-ChildItem $base | ForEach-Object{
      $p=(Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue)
      if($p.Path -and ($p.Path.TrimEnd('\')+'\') -ieq ($path.TrimEnd('\')+'\') -and ($p.Description -like "$desc*")){
        Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue; $removed=$true
      }
    }
  }
  return $removed
}

# Excel OPEN registry
function Get-ExcelOptionsKeyPaths{ foreach($v in Get-OfficeVersions){ $k="HKCU:\Software\Microsoft\Office\$v\Excel\Options"; if(Test-Path $k){ $k } } }
function Ensure-OpenEntry($path){
  foreach($k in Get-ExcelOptionsKeyPaths){
    $props=Get-ItemProperty -Path $k -ErrorAction SilentlyContinue
    $existing=$props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | Select-Object -ExpandProperty Value
    if($existing -and $existing -contains $path){ return $true }
    for($i=1;;$i++){
      $name=if($i -eq 1){'OPEN'}else{"OPEN$($i-1)"}
      $cur=(Get-ItemProperty -Path $k -Name $name -ErrorAction SilentlyContinue).$name
      if(-not $cur){ New-ItemProperty -Path $k -Name $name -Value $path -PropertyType String -Force | Out-Null; return $true }
    }
  }
  return $false
}
function Remove-OpenEntry($path){
  foreach($k in Get-ExcelOptionsKeyPaths){
    $props=Get-ItemProperty -Path $k -ErrorAction SilentlyContinue
    foreach($p in $props.PSObject.Properties){
      if($p.Name -match '^OPEN\d*$' -and $p.Value -eq $path){ Remove-ItemProperty -Path $k -Name $p.Name -ErrorAction SilentlyContinue }
    }
  }
}

# Excel COM
function With-ExcelCOM([scriptblock]$action){
  $excel=$null
  try{
    $excel=New-Object -ComObject Excel.Application
    $excel.DisplayAlerts=$false; $excel.Visible=$false
    & $action $excel | Out-Null
  } finally {
    if($excel){ try{ $excel.Quit() | Out-Null }catch{} }
    if($excel){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}
function Try-LoadAddIn-COM($full){
  try{
    With-ExcelCOM {
      param($excel)
      $name=[IO.Path]::GetFileName($full)
      $a=$excel.AddIns | Where-Object { $_.FullName -ieq $full -or $_.Name -ieq $name }
      if(-not $a){ $a=$excel.AddIns.Add($full,$false) }
      $a.Installed=$true
    }; $true
  } catch { $false }
}
function Try-UnloadAddIn-COM($full){
  try{
    With-ExcelCOM {
      param($excel)
      $name=[IO.Path]::GetFileName($full)
      $excel.AddIns | Where-Object { $_.FullName -ieq $full -or $_.Name -ieq $name } | ForEach-Object { $_.Installed=$false }
    }; $true
  } catch { $false }
}

# Detection (file must exist)
function Detect-Installed{
  $present=@()
  $cand=Join-Path $TargetDir $AddInFile
  if(Test-Path $cand){ $present+=$cand }

  foreach($k in Get-ExcelOptionsKeyPaths){
    $props=Get-ItemProperty -Path $k -ErrorAction SilentlyContinue
    foreach($p in $props.PSObject.Properties){
      if($p.Name -match '^OPEN\d*$' -and $p.Value -and (Split-Path $p.Value -Leaf) -ieq $AddInFile){
        if(Test-Path $p.Value){ $present+=$p.Value }
      }
    }
  }

  try{
    With-ExcelCOM {
      param($excel)
      $matches=$excel.AddIns | Where-Object { ($_.Name -ieq $AddInFile -or (Split-Path $_.FullName -Leaf) -ieq $AddInFile) -and $_.Installed }
      foreach($m in $matches){ if($m.FullName -and (Test-Path $m.FullName)){ $script:__paths=($script:__paths+$m.FullName) } }
    }
  } catch {}
  ($present + $script:__paths) | Where-Object { $_ } | Select-Object -Unique
}

# Actions
function Install-Addin{
  Status "Installing files" -Fatal {
    Ensure-Dir $TargetDir
    foreach($f in $FilesToInstall){
      $src=Join-Path $SourceDir $f; $dst=Join-Path $TargetDir $f
      if(-not(Test-Path $src)){ throw "Missing file '$f' in $SourceDir" }
      Copy-Item $src $dst -Force
      if($f -like '*.xla*'){ try{ Unblock-File -Path $dst -ErrorAction SilentlyContinue }catch{} }
    }
  }
  Status "Adding to Trusted Location" { if(-not(Add-TrustedLocation -path $TargetDir -desc $TrustedDesc)){ throw "Could not add location" } }
  $addinPath=Join-Path $TargetDir $AddInFile
  Status "Loading add-in" {
    if(-not(Try-LoadAddIn-COM $addinPath)){ if(-not(Ensure-OpenEntry $addinPath)){ throw "Could not register add-in" } }
  }
}
function Uninstall-Addin($paths){
  foreach($p in $paths){
    Status "Unloading add-in"          { Try-UnloadAddIn-COM $p | Out-Null }
    Status "Removing Excel OPEN entry" { Remove-OpenEntry $p | Out-Null }
  }
  Status "Removing files" {
    foreach($f in $FilesToInstall){
      $dst=Join-Path $TargetDir $f
      if(Test-Path $dst){ Remove-Item $dst -Force -ErrorAction SilentlyContinue }
    }
  }
  Status "Removing Trusted Location" {
    if(-not(Remove-TrustedLocation -path $TargetDir -desc $TrustedDesc)){ throw "No trusted location found" }
  }
}

# Main
$paths=Detect-Installed
if(-not $paths -or $paths.Count -eq 0){
  $ans=Read-Host "exceladdin is NOT installed. Install now? (Y/N)"
  if($ans -match '^[Yy]'){ Install-Addin }
} else {
  $ans=Read-Host ("exceladdin is installed at " + ($paths -join ', ') + ". Uninstall it? (Y/N)")
  if($ans -match '^[Yy]'){ Uninstall-Addin $paths }
}

# ... keep your Main block ...

# success/final message
Write-Host ""
Write-Host "All done. You can close this window now to proceed." -ForegroundColor Yellow

exit 0

