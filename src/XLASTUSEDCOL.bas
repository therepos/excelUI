Private Function XLASTUSEDCOL(rng As Range) As Long
'   Purpose: Find last column
'   Reference: https://www.rondebruin.nl/win/s9/win005.htm

    Dim result As Long
          
    On Error Resume Next
    result = rng.Find(What:="*", _
                After:=rng.Cells(1), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Column
                
    XLASTUSEDCOL = result
    If Err.Number <> 0 Then
        XLASTUSEDCOL = rng.Column + rng.Columns.Count - 1
    End If
         
End Function

