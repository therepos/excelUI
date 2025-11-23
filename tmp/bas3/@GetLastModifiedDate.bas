Private Function GetLastModifiedDate(wbname As String)
'   Purpose: Get last modified date of workbook
'   Reference: https://stackoverflow.com/questions/16657170/last-modification-date-of-open-workbook

    Dim rv, wb As Workbook

    rv = "workbook?" 'default return value

    On Error Resume Next
    Set wb = Workbooks(wbname)
    On Error GoTo 0

    If Not wb Is Nothing Then
        rv = Format(FileDateTime(wb.FullName), "m/d/yy h:n ampm")
    End If

    LastModifiedDate = rv

End Function