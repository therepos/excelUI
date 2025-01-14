Private Sub GetFileLastAccessedDate()
'   Purpose: Get last accessed date of a file using command through Shell object
'   Reference: https://superuser.com/questions/814298/command-to-get-files-with-last-access-date-in-windows

    Call Shell("cmd.exe /c dir /s /ta ""c:\temp""", vbNormalFocus)

End Sub