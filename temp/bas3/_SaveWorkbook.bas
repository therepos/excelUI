Private Sub SaveWorkbook()
'   Purpose: Saves a copy of active workbook in temp folder for recovery
'   Todo: Create a new folder if temp folder is not accessible on system
'       : https://stackoverflow.com/questions/43658276/create-folder-path-if-does-not-exist-saving-from-vba
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler

    Dim strFileExists As String
    
    strFileExists = Dir(ActiveWorkbook.FullName)
    If strFileExists = "" Then
        ActiveWorkbook.SaveCopyAs "c:\temp\" & "copy_" & ActiveWorkbook.Name & ".xlsx"
    Else
        ActiveWorkbook.SaveCopyAs "c:\temp\" & "copy_" & ActiveWorkbook.Name
    End If

ErrorHandler:
    Exit Sub

End Sub