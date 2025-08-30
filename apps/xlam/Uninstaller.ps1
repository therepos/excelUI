# Only these:
$addinName = "exceladdin.xlam"
$appsXlam  = "C:\Apps\exceladdin.xlam"
$appsUI    = "C:\Apps\exceladdin.exportedUI"
$addinsXlam= Join-Path $env:APPDATA "Microsoft\AddIns\exceladdin.xlam"
$officeUI  = Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"
$localUI   = Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI"

# 0) Close Excel (so nothing re-loads mid-uninstall)
Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 2

# 1) Open Excel once to remove the add-in registration & uncheck it
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$wb = $xl.Workbooks.Add()

# helper to uninstall by FullName/Name on BOTH collections
function Remove-Addin($col, $paths, $name){
  $hits = @()
  foreach($ai in $col){
    if(($ai.FullName -and $paths -contains $ai.FullName) -or ($ai.Name -ieq $name)){
      try { $ai.Installed = $false } catch {}
      $hits += $ai.Name
      try { $col.Item($ai.Name).Delete() } catch {}   # remove from list
    }
  }
  return $hits
}

$paths = @($appsXlam,$addinsXlam) | Where-Object { $_ }
$hits1 = Remove-Addin $xl.AddIns  $paths $addinName
$hits2 = Remove-Addin $xl.AddIns2 $paths $addinName

$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

# 2) Delete files with retries (handles locks)
function Remove-WithRetry([string]$p){
  if(-not (Test-Path $p)) { return }
  for($i=1;$i -le 5;$i++){
    try { Attrib -R $p -ErrorAction SilentlyContinue; Remove-Item -LiteralPath $p -Force; if(-not (Test-Path $p)){ return } }
    catch { Start-Sleep -Milliseconds 400 }
  }
}
$toDelete = @($appsXlam,$addinsXlam,$appsUI)
$toDelete | ForEach-Object { Remove-WithRetry $_ }

# 3) Remove ribbon definition file(s) (backup then delete) – this kills the custom tab
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
foreach($ui in @($officeUI,$localUI)){
  if(Test-Path $ui){
    Copy-Item $ui "$ui.bak.$stamp" -Force
    Remove-Item $ui -Force -ErrorAction SilentlyContinue
  }
}

# 4) Final verification (prints anything left to help debug)
Write-Host "Leftovers (should be empty if fully removed):"
@($appsXlam,$addinsXlam) | Where-Object { Test-Path $_ } | ForEach-Object { "  $_" } | Write-Host
Write-Host "Done. Reopen Excel."
