Sub GetExtendedFileDetails()
'   Purpose: Get extended file properties by ID
'   Reference: https://www.access-programmers.co.uk/forums/threads/windows-10-retrieving-extended-file-properties.294416/
'   Reference: https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
'   Reference: https://sourcedaddy.com/windows-7/important-startup-files.html
'   Reference: https://www.codeproject.com/Articles/1101956/Check-for-Clock-Tampering-to-Extend-Licence-Durati#:~:text=The%20Basic%20Check&text=Do%20this%20every%20time%20the,clock%20has%20been%20tampered%20with.
'   Reference: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/gg264782(v=office.14)
'   Notes:
'   - Const strFilePath = "C:\Windows\System32\winlogon.exe"
'   - Does not work on ..System32\winlogon.exe
'   Status:
'   - Function not working on some files

    Dim strFilePath As String
    strFilePath = Environ("UserProfile") & "\Google Drive\Ad Astra 001.mp3"
    
    Dim objShell As Shell32.Shell
    Dim objFolder As Shell32.Folder
    Dim objFolderItem As Shell32.FolderItem
    Dim strPath As String, strFileName As String
    Dim aName As String, aValue As String
    Dim I As Integer
    Dim tComments As String, tCategory As String, tTitle As String
    
    I = 1
    strFileName = strFilePath
    'Find the last "\" and get the filename
    Do Until I = 0
        I = InStr(1, strFileName, "\", vbBinaryCompare)
        strFileName = Mid(strFileName, I + 1)
    Loop
    strPath = Left(strFilePath, Len(strFilePath) - Len(strFileName) - 1)
    Set objShell = New Shell
    Set objFolder = objShell.Namespace(strPath)
    Set objFolderItem = objFolder.ParseName(strFileName)
        
    For I = 0 To 10
        aName = objFolder.GetDetailsOf(objFolder.Items, I)
        aValue = objFolder.GetDetailsOf(objFolderItem, I)
        
        'List attributes available in this folder:
        Debug.Print I & " - " & objFolder.GetDetailsOf(objFolder.Items, I) & ": " & _
                  objFolder.GetDetailsOf(objFolderItem, I)
    Next I
    
    Set objFolderItem = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing
     
End Sub