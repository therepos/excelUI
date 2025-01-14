Private Function XFIRSTUSEDROW(rng As Range) As Long

    Dim result As Long

    On Error Resume Next
    If IsEmpty(rng.Cells(1)) Then
        result = rng.Find(What:="*", _
               After:=rng.Cells(1), _
               Lookat:=xlPart, _
               LookIn:=xlFormulas, _
               SearchOrder:=xlByRows, _
               SearchDirection:=xlNext, _
               MatchCase:=False).Row
    Else: result = rng.Cells(1).Row
    End If
    XFIRSTUSEDROW = result
    If Err.Number <> 0 Then
        XFIRSTUSEDROW = 0
    End If
         
End Function
