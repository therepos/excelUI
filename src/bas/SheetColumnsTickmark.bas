Sub SheetColumnsTickmark(control As IRibbonControl)
'   Purpose: Standardise columns width for specific worksheet: Tickmark tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.ColumnWidth = 15
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 1
    Columns("C").ColumnWidth = 3
    Columns("D").ColumnWidth = 15
    Columns("E").ColumnWidth = 15
    Columns("F").ColumnWidth = 15
    Columns("G").ColumnWidth = 15
    Columns("H").ColumnWidth = 15
    Columns("I").ColumnWidth = 15
    Columns("J").ColumnWidth = 15
    Columns("K").ColumnWidth = 15
    Columns("L").ColumnWidth = 15
    Columns("M").ColumnWidth = 1
    Columns("N").ColumnWidth = 5
    
ErrorHandler:
    Exit Sub

End Sub

