Sub InsertCrossReference(control As IRibbonControl)
'   Purpose: Create hyperlink based on targeted cell
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Call XHYPERACTIVE(Selection)
    
ErrorHandler:
    Exit Sub

End Sub

