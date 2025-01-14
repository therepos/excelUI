Sub FormulaToValue(control As IRibbonControl)
'   Purpose: Convert selected formulas to values
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    
ErrorHandler:
    Exit Sub

End Sub

