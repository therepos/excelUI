Function XREMOVESYMBOLS(text As String, opType As String, Optional charCount)
'   Purpose: Substitute a leading symbol
'   Feature: By default leading 3 characters are considered if not specified
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim varSymbols() As Variant
    Dim oldText As String
    Dim n As Integer

    varSymbols = Array(Chr(34), "-")
    
    Select Case opType
        Case "Full"
            For n = 0 To 1
                oldText = varSymbols(n)
                text = WorksheetFunction.Substitute(text, oldText, "")
            Next n
        Case "LEAD"
            For n = 0 To 1
                oldText = varSymbols(n)
                text = SUBSTITUTELEADING(text, oldText, "", charCount)
            Next n
        Case "TRAIL"
             For n = 0 To 1
                oldText = varSymbols(n)
                text = SUBSTITUTETRAILING(text, oldText, "", charCount)
            Next n
    End Select

    XREMOVESYMBOLS = Trim(text)

End Function

