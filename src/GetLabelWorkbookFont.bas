Sub GetLabelWorkbookFont(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFont
        Case "ddSelectionFont01": returnedVal = "Arial"
        Case "ddSelectionFont02": returnedVal = "Verdana"
        Case "ddSelectionFont03": returnedVal = "Times"
    End Select

End Sub

