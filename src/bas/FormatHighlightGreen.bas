Attribute VB_Name = "Module1"
Sub FormatHighlightGreen()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(204, 285, 204)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub