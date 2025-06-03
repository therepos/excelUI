Set shell = CreateObject("WScript.Shell")

' Add C:\Apps to Trusted Locations
excelVer = GetExcelVersion()
trustKey = "HKCU\Software\Microsoft\Office\" & excelVer & "\Excel\Security\Trusted Locations\Location99\"
shell.RegWrite trustKey & "Path", "C:\Apps", "REG_SZ"
shell.RegWrite trustKey & "AllowSubfolders", 1, "REG_DWORD"
shell.RegWrite trustKey & "Description", "SET Addin Folder", "REG_SZ"

' Register the Add-in
Set xlApp = CreateObject("Excel.Application")
Set addin = xlApp.AddIns.Add("C:\Apps\SET-Addin.xlam", True)
addin.Installed = True
xlApp.Quit

Function GetExcelVersion()
    Dim xl
    Set xl = CreateObject("Excel.Application")
    GetExcelVersion = Int(xl.Version)
    xl.Quit
End Function
