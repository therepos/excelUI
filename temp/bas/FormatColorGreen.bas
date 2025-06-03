Sub FormatColorGreen(control As IRibbonControl)
'   Purpose: To highlight range for follow-up
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Set rng = Selection
    
    rng.Interior.Color = RGB(204, 285, 204)
    
ErrorHandler:
    Exit Sub

End Sub

