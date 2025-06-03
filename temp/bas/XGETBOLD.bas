Function XGETBOLD(pWorkRng As Range)

    If pWorkRng.Font.Bold Then
        XGETBOLD = True
    Else
        XGETBOLD = False
    End If
    
End Function
