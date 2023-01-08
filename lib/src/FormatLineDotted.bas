Sub FormatLineDotted(control As IRibbonControl)
'   Purpose: Insert dotted line
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    
ErrorHandler:
    Exit Sub

End Sub

