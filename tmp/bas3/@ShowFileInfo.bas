Sub ShowFileInfo(filespec)
'   Purpose: Estimate user login timestamp to counter expiry tampering
'   Reference: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object
'   Reference: https://www.automateexcel.com/vba/getfolder-getfile/
'   Reference: [TextStream Objects only] https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'   Requirement: Microsoft Scripting Runtime

    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.DateLastAccessed
    
    MsgBox s
    
End Sub