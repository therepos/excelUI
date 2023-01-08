Function XCOLUMNWIDTH(target As Range) As Double
'   Purpose: Get column width
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    
    XCOLUMNWIDTH = target.ColumnWidth
    Application.ScreenUpdating = True
    
End Function

