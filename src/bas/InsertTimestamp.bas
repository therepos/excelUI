Attribute VB_Name = "Module1"
Sub InsertTimestamp()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Value = Now
    Rng.NumberFormat = "dd-mmm-yy"
    Rng.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub