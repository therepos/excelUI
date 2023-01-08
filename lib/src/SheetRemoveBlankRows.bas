Sub SheetRemoveBlankRows(control As IRibbonControl)
'   Purpose: Remove blank rows in sheet
'   Reference: https://www.ablebits.com/office-addins-blog/2018/12/19/delete-blank-lines-excel/
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
        
    Dim SourceRange As Range
    Dim EntireRow As Range
 
    Set SourceRange = Application.ActiveSheet.UsedRange
 
    If Not (SourceRange Is Nothing) Then
        For i = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
    End If
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

