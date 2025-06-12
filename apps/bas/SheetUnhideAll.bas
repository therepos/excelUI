Sub SheetUnhideAll(control As IRibbonControl)
'   Purpose: Unhide all rows and columns
'   Reference: https://www.extendoffice.com/documents/excel/1173-excel-break-all-links.html
'   Updated: 2022MAR12

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    ActiveSheet.Columns.EntireColumn.Hidden = False
    ActiveSheet.Rows.EntireRow.Hidden = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

