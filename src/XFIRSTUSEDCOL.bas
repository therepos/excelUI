Private Function XFIRSTUSEDCOL(rng As Range) As Long

    Dim result As Long
          
    On Error Resume Next
    result = rng.Find(What:="*", _
                After:=rng.Cells(1), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, _
                MatchCase:=False).Column
                
    XFIRSTUSEDCOL = result
    If Err.Number <> 0 Then
        XFIRSTUSEDCOL = rng.Column + rng.Columns.Count - 1
    End If
         
End Function
