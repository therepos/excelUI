Sub ColorBordersAll(control As IRibbonControl)
'   Purpose: Change Border Colors without affecting thickness/styles
'   Reference: www.TheSpreadsheetGuru.com/the-code-vault
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim cell As Range
    Dim DesiredColor As Long

'   Color To Change Borders To
    DesiredColor = RGB(198, 198, 198)

'   Ensure Range is Selected
    If TypeName(Selection) <> "Range" Then Exit Sub

'   Loop Through each cell in selection and change border color (if applicable)
    For Each cell In Selection.Cells
        cell.Borders(xlEdgeTop).Color = DesiredColor
        cell.Borders(xlEdgeBottom).Color = DesiredColor
        cell.Borders(xlEdgeLeft).Color = DesiredColor
        cell.Borders(xlEdgeRight).Color = DesiredColor
    Next cell

ErrorHandler:
    Exit Sub

End Sub

