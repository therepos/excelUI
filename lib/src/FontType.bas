Sub FontType(control As IRibbonControl)
'   Purpose: Set selected range to Arial
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
'   ===================================
'   Customised use-case
'   ===================================
    Dim userFont As String
    Dim userFontSize As Long
    
    Select Case MySelectedFont
        Case "ddSelectionFont01": userFont = "Arial"
        Case "ddSelectionFont02": userFont = "Verdana"
        Case "ddSelectionFont03": userFont = "Times New Roman"
        Case "": userFont = "Arial"
    End Select
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ===================================

    Dim rng As Range
    Set rng = Selection
    rng.Font.Name = userFont
    rng.Font.Size = userFontSize
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

