Sub RemoveBlankRows(control As IRibbonControl)
'   Purpose: Remove blank rows in selection
'   Reference: https://www.wallstreetmojo.com/vba-last-row/
'   Reference: https://www.rondebruin.nl/win/s9/win005.htm
'   Notes:
'   - Selection.SpecialCells(xlCellTypeLastCell).Row    Return the last used in the worksheet regardless of selection
'   - Selection.Find Method                             Returns the last used based on selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
        
    Dim SourceRange As Range
    Dim TargetRow As Range
    Dim lastRow As Long
    Dim firstRow As Long
    
    Set SourceRange = Selection
    firstRow = SourceRange.Cells(1).row
    
    If XLASTUSEDROW(SourceRange) > 0 Then
        lastRow = XLASTUSEDROW(SourceRange)
    Else
        lastRow = SourceRange.row + SourceRange.Rows.Count - 1
    End If
    
    For i = lastRow To firstRow Step -1
        Set TargetRow = Cells(i, 1).EntireRow
        If Application.WorksheetFunction.CountA(TargetRow) = 0 Then
            TargetRow.Delete
        End If
    Next
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

