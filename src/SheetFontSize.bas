Sub SheetFontSize(control As IRibbonControl)
'   Purpose: Standardise workbook font size
'   Updated: 2022MAR12

'   Saves worksheet before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

'   ===================================
'   Customised use-case
'   ===================================
    Dim userFontSize As Long
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ==================================

    With ActiveSheet
       .Cells.Font.Size = userFontSize
    End With
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

