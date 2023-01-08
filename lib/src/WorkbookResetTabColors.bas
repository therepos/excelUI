Sub WorkbookResetTabColors(control As IRibbonControl)
'   Purpose: Reset all tab colors
'   Reference: https://www.extendoffice.com/documents/excel/5179-excel-remove-tab-color.html
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim xSheet As Worksheet
    
    For Each xSheet In ActiveWorkbook.Worksheets
        xSheet.Tab.ColorIndex = xlColorIndexNone
    Next xSheet
    
ErrorHandler:
    Exit Sub

End Sub

