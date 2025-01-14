Sub FormatHighlightReset()
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Interior.Color = xlNone
    Application.ScreenUpdating = True

End Sub
