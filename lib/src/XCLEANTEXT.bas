Function XCLEANTEXT(text As String)
'   Purpose: Removes excess non-alphanumeric characters
'   Usage: =LEN(B3)-LEN(SUBSTITUTE(B3,C3,))
'   Feature: To count number of delimiter
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save

'   Remove leading and training symbols
    text = REMOVESYMBOLS(text, LEAD)
    text = REMOVESYMBOLS(text, TRAIL)

'   Replace hanging double quotation marks
    Dim m As Integer

    For m = 1 To Len(text)
        If Mid(text, m, 1) = Chr(34) Then
            If m = 1 Then
                If Not IsNumeric(Mid(text, m - 1, 1)) Then
                    text = Left(text, m - 1) & Right(text, Len(text) - m)
                    m = m - 1
                End If
            Else
                text = REMOVESYMBOLS(text, LEAD)
            End If
        End If
    Next m
'   Double spacing
    text = WorksheetFunction.Substitute(text, "  ", "")
'   Comma
    text = WorksheetFunction.Substitute(text, ",", "")
    XCLEANTEXT = Trim(text)

End Function

