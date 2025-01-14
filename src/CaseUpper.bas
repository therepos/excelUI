Sub CaseUpper(control As IRibbonControl)
'   Purpose: Set upper case on selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    For Each cell In rng
        cell.Value = UCase(cell)
    Next cell
    
ErrorHandler:
    Exit Sub

End Sub

