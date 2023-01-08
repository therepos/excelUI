Sub SheetFormulaToValue(control As IRibbonControl)
'   Purpose: Convert all worksheet formulas to values (most efficient way)
'   Updated: 2022MAR12

'   Saves worksheet before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
     
    ActiveSheet.Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

