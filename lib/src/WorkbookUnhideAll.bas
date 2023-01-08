Sub WorkbookUnhideAll(control As IRibbonControl)
'   Purpose: Unhide all rows and columns
'   Reference: https://www.extendoffice.com/documents/excel/1173-excel-break-all-links.html
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Columns.EntireColumn.Hidden = False
        ws.Rows.EntireRow.Hidden = False
    Next ws
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

