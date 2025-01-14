Private Function XGETHUNDREDS(ByVal MyNumber)
'   Purpose: Converts a number from 100-999 into text

    Dim result As String
    
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    
'   Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        result = XGETDIGIT(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    
'   Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        result = result & XGETTENS(Mid(MyNumber, 2))
    Else
        result = result & XGETDIGIT(Mid(MyNumber, 3))
    End If
    
    XGETHUNDREDS = result

End Function

