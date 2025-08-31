Private Function XLASTUSEDROW(rng As Range) As Long
'   Purpose: Find last used row
'   Reference: https://www.rondebruin.nl/win/s9/win005.htm
'   Note: Returns 0 if not found

    Dim result As Long

    On Error Resume Next
    result = rng.Find(What:="*", _
               After:=rng.Cells(1), _
               Lookat:=xlPart, _
               LookIn:=xlFormulas, _
               SearchOrder:=xlByRows, _
               SearchDirection:=xlPrevious, _
               MatchCase:=False).row
                
    XLASTUSEDROW = result
    If Err.Number <> 0 Then
        XLASTUSEDROW = 0
    End If
         
End Function

