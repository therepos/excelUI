Sub WorkbookFont(control As IRibbonControl)
'   Purpose: Standardise workbook font type and size
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
'   ==================================
    
    Dim ws As Worksheet
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = userFont
            .Cells.Font.Size = userFontSize
         End With
    Next ws
    For Each ws In Worksheets
        ws.Activate
        ActiveWindow.Zoom = 100
    Next

ErrorHandler:
    Exit Sub
    
    Application.ScreenUpdating = True
    
End Sub

