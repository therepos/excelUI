Sub InsertColumnWidth(control As IRibbonControl)
'   Purpose: Insert column width counter
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    For Each c In rng
        c.Formula = "=" & "XCOLUMNWIDTH(" & c.Address & ")"
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0.0_);_((#,##0.0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

