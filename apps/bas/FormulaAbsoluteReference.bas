Sub FormulaAbsoluteReference(control As IRibbonControl)
'   Purpose: Absolute reference selected cells
'   Reference: http://www.excelforum.com/excel-general/372383-making-multiple-cells-absolute-at-once.html
'   Reference: http://www.contextures.com/xlvba01.html#videoreg
'   Todo: Check if cell formula is already referenced.
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim cell As Range
    
    For Each cell In Selection
        If cell.HasFormula Then
            cell.Formula = _
            Application.ConvertFormula(cell.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next
 
ErrorHandler:
    Exit Sub

End Sub

