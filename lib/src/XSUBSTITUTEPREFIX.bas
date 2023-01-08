Function XSUBSTITUTEPREFIX(text As String, oldText As String, newText As String, Optional charCount)
'   Purpose: Substitute a leading symbol
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

    leadText = Left(text, i)
    trailText = Right(text, Len(text) - i)

    leadText = WorksheetFunction.Substitute(leadText, oldText, newText)
    text = leadText & trailText

    XSUBSTITUTEPREFIX = Trim(text)

End Function

