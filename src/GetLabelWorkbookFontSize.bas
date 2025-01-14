Sub GetLabelWorkbookFontSize(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": returnedVal = 8
        Case "ddSelectionFontSize02": returnedVal = 9
        Case "ddSelectionFontSize03": returnedVal = 10
        Case "ddSelectionFontSize04": returnedVal = 11
    End Select

End Sub

