Sub RemoveNamedRanges(control As IRibbonControl)
'   Purpose: Delete all named ranges
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim i As Long
    
    Application.Calculation = xlCalculationManual
    For i = ThisWorkbook.Names.Count To 1 Step -1
        ThisWorkbook.Names(i).Delete
    Next
    Application.Calculation = xlCalculationAutomatic

ErrorHandler:
    Exit Sub

End Sub

