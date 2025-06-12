Sub SheetColumnsWP(control As IRibbonControl)
'   Purpose: Standardise workbook columns width
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    For Each ws In Worksheets
        Columns.ColumnWidth = 14
        Columns("A").ColumnWidth = 1
        Columns("B").ColumnWidth = 3
        Columns("C").ColumnWidth = 5
    Next ws
    
ErrorHandler:
    Exit Sub

End Sub

