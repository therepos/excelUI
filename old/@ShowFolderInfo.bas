Sub ShowFolderInfo(folderspec)
'   Purpose: Estimate user login timestamp to counter expiry tampering
'   Reference: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folder-object
'   Reference: https://www.automateexcel.com/vba/getfolder-getfile/
'   Note:
'   - Assumes certain folders are accessed by Windows on startup and attempts to use its accessed time

    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    s = f.DateLastModified

    MsgBox s

End Sub