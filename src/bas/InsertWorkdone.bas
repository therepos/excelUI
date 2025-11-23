Attribute VB_Name = "Module1"
Sub InsertWorkdone()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Value = "Keys to Workdone:"
    Rng.Font.Bold = True
    Rng.Offset(1, 0) = "TB"
    Rng.Offset(1, 1) = ": Agreed to current year trial balance."
    Rng.Offset(2, 0) = "PY"
    Rng.Offset(2, 1) = ": Agreed to prior year audited balance."
    Rng.Offset(3, 0) = "imm"
    Rng.Offset(3, 1) = ": Immaterial (below SUM), suggest to leave."
    Rng.Offset(4, 0) = "^"
    Rng.Offset(4, 1) = ": Casted."
    Rng.Offset(5, 0) = "Cal"
    Rng.Offset(5, 1) = ": Calculation checked."
    Rng.Offset(1, 0).Characters(1, 3).Font.Color = RGB(0, 112, 192)
    Rng.Offset(2, 0).Characters(1, 3).Font.Color = RGB(255, 51, 0)
    Rng.Offset(3, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(4, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(5, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(1, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(2, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(3, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(4, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(5, 0).Characters(1, 3).Font.Bold = True
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub