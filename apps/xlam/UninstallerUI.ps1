# Uninstall excelUI.xlam only (no exportedUI)

$ErrorActionPreference = 'Stop'

# Paths
$addinName  = "excelUI.xlam"
$appsXlam   = "C:\Apps\excelUI.xlam"
$addinsXlam = Join-Path $env:APPDATA "Microsoft\AddIns\excelUI.xlam"

# 0) Close Excel
Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 2

# 1) Open Excel to unregister the add-in
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$wb = $xl.Workbooks.Add()

function Remove-Addin($col, $paths, $name){
  $hits = @()
  foreach($ai in $col){
    if(($ai.FullName -and $paths -contains $ai.FullName) -or ($ai.Name -ieq $name)){
      try { $ai.Installed = $false } catch {}
      $hits += $ai.Name
      try { $col.Item($ai.Name).Delete() } catch {}
    }
  }
  return $hits
}

$paths = @($appsXlam,$addinsXlam) | Where-Object { $_ }
[void](Remove-Addin $xl.AddIns  $paths $addinName)
[void](Remove-Addin $xl.AddIns2 $paths $addinName)

$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

# 2) Delete xlam files with retries
function Remove-WithRetry([string]$p){
  if(-not (Test-Path $p)) { return }
  for($i=1;$i -le 5;$i++){
    try {
      Attrib -R $p -ErrorAction SilentlyContinue
      Remove-Item -LiteralPath $p -Force
      if(-not (Test-Path $p)){ return }
    } catch { Start-Sleep -Milliseconds 400 }
  }
}
$toDelete = @($appsXlam,$addinsXlam)
$toDelete | ForEach-Object { Remove-WithRetry $_ }

# 3) Final verification
Write-Host "Leftovers (should be empty if fully removed):"
@($appsXlam,$addinsXlam) | Where-Object { Test-Path $_ } | ForEach-Object { "  $_" } | Write-Host
Write-Host "Done. Reopen Excel."
