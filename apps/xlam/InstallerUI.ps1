<#
 Installer.XlamOnly.ps1
 - Deploys excelUI.xlam to C:\Apps
 - Adds C:\Apps as a trusted location
 - Registers the add-in from that folder
#>

$ErrorActionPreference = 'Stop'

# Resolve script folder robustly
$srcDir = if ($PSScriptRoot) { $PSScriptRoot }
          elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
          else { (Get-Location).Path }

$xlamSrc = Join-Path $srcDir 'excelUI.xlam'
if (!(Test-Path $xlamSrc)) { throw "Missing excelUI.xlam in $srcDir" }

$appsDir  = 'C:\Apps'
$appsXlam = Join-Path $appsDir 'excelUI.xlam'

# --- Kill Excel if running ---
Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 1

# --- Copy payload to C:\Apps ---
New-Item -ItemType Directory -Force -Path $appsDir | Out-Null
Copy-Item -LiteralPath $xlamSrc -Destination $appsXlam -Force
Unblock-File -Path $xlamSrc,$appsXlam -ErrorAction SilentlyContinue
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Copied -> $appsXlam"

# --- Add C:\Apps to trusted locations (HKCU) ---
foreach ($v in '16.0','15.0','14.0') {
  $k="HKCU:\Software\Microsoft\Office\$v\Excel\Security\Trusted Locations\XLAMDeploy"
  if (!(Test-Path $k)) { New-Item $k -Force | Out-Null }
  Set-ItemProperty $k Path ($appsDir.TrimEnd('\')+'\')
  Set-ItemProperty $k Description 'Trusted XLAM deployment'
  Set-ItemProperty $k AllowSubfolders 0
  Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Trusted location set: $k"
}

# --- Register the add-in from C:\Apps ---
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

try {
  # Prefer exact path match; fall back to adding it
  $ai = $xl.AddIns2 | Where-Object { $_.FullName -ieq $appsXlam }
  if (-not $ai) { $ai = $xl.AddIns2.Add($appsXlam, $true) }
  $ai.Installed = $true
  Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Add-in activated from $appsXlam"
} catch {
  Write-Warning "AddIn activation failed: $($_.Exception.Message)"
} finally {
  $wb.Close($false)
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
  [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ✅ Done. Restart Excel to use the add-in."
