@echo off
echo Installing SET Add-in...
mkdir "C:\Apps" >nul 2>&1
copy /Y "SET-Addin.xlam" "C:\Apps\SET-Addin.xlam"
copy /Y "SET-Addin.exportedUI" "%APPDATA%\Microsoft\Office"
cscript //nologo Install.vbs
echo Done. You can now open Excel.
pause
