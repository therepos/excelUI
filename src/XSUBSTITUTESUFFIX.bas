Function XSUBSTITUTESUFFIX(text As String, oldText As String, newText As String, Optional charCount)
'   Purpose: Substitute a trailing symbol
'   Feature: By default leading 3 characters are considered if not specified
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim i As Integer
    Dim leadText As String
    Dim trailText As String
    
    If IsMissing(charCount) Then
        i = 3
    Else
        i = charCount
    End If

    If Len(text) = i Then
        i = Len(text)
    End If

    trailText = Right(text, i)
    leadText = Left(text, Len(text) - i)

    trailText = WorksheetFunction.Substitute(trailText, oldText, newText)
    text = leadText & trailText

    XSUBSTITUTESUFFIX = Trim(text)

End Function

