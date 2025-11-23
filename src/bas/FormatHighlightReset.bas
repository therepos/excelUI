Attribute VB_Name = "Module1"
Sub FormatHighlightReset()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = xlNone
    Rng.Font.Color = RGB(0, 0, 0)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub