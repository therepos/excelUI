
Sub FormatCellRed()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 204, 204)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub