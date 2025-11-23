Attribute VB_Name = "Module1"
Sub InsertColumnWidth()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In Rng
        c.Formula = "=" & "XCOLUMNWIDTH(" & c.Address & ")"
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0.0_);_((#,##0.0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub