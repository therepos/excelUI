Sub ColorBordersOuter(control As IRibbonControl)
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

    Selection.BorderAround , Color:=DesiredColor, Weight:=xlThin

ErrorHandler:
    Exit Sub

End Sub

