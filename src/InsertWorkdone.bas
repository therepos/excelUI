Sub InsertWorkdone(control As IRibbonControl)
'   Purpose: Insert customised legend for workdone
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    
    rng.Value = "Legend:"
    rng.Font.Bold = True
    rng.Offset(1, 0) = "TB: Agreed to current year trial balance."
    rng.Offset(2, 0) = "PY: Agreed to prior year audited balance."
    rng.Offset(3, 0) = "i: Immaterial (below CTT), suggest to leave."
    rng.Offset(4, 0) = "GL: Agreed to current year general ledger."
    rng.Offset(1, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(2, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(3, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(4, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(1, 0).Characters(1, 2).Font.Bold = True
    rng.Offset(2, 0).Characters(1, 2).Font.Bold = True
    rng.Offset(3, 0).Characters(1, 2).Font.Bold = True
    rng.Offset(4, 0).Characters(1, 2).Font.Bold = True
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

