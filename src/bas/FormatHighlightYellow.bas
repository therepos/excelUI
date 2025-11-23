Attribute VB_Name = "Module1"
Sub FormatHighlightYellow()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 255, 0)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub