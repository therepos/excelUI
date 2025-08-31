Sub SortRight(control As IRibbonControl)
'   Purpose: Sort a series of numbers from left to right
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim row As Range
    
    For Each row In Selection.Rows
        row.Sort Key1:=row, Order1:=xlAscending, Orientation:=xlSortRows
    Next row

ErrorHandler:
    Exit Sub

End Sub

