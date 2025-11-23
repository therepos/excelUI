Sub CaseUpper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = UCase(cell)
    Next cell
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub