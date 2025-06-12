Sub FormatTextToValue(control As IRibbonControl)
'   Purpose: Convert text format to number format on selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlRight
            c.NumberFormat = "General"
            c.Value = c.Value
'       End If
    Next c

ErrorHandler:
    Exit Sub

End Sub

