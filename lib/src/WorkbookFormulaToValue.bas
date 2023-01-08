Sub WorkbookFormulaToValue(control As IRibbonControl)
'   Purpose: Convert all workbook formulas to values (most efficient way)
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim sh As Worksheet, HidShts As New Collection

    For Each sh In ActiveWorkbook.Worksheets
        If Not sh.visible Then
            HidShts.Add sh
            sh.visible = xlSheetVisible
        End If
    Next sh
     
    Worksheets.Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
     
    For Each sh In HidShts
        sh.visible = xlSheetHidden
    Next sh
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

