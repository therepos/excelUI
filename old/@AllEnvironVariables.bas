Sub AllEnvironVariables()
'   Purpose: Get all Environ variables
'   Reference: https://wellsr.com/vba/2019/excel/list-all-environment-variables-with-vba-environ/
'   Reference: https://www.codevba.com/office/environ.htm#.YiNGjuhByMp
'   Reference: [SystemInfo] https://vbaoverall.com/find-complete-system-information-of-your-pc-vba-code-example/

    Dim strEnviron As String
    Dim VarSplit As Variant
    Dim I As Long
    For I = 1 To 255
        strEnviron = Environ$(I)
        If LenB(strEnviron) = 0& Then GoTo TryNext:
        VarSplit = Split(strEnviron, "=")
        If UBound(VarSplit) > 1 Then Stop
        Range("A" & Range("A" & Rows.Count).End(xlUp).Row + 1).Value = I
        Range("B" & Range("B" & Rows.Count).End(xlUp).Row + 1).Value = VarSplit(0)
        Range("C" & Range("C" & Rows.Count).End(xlUp).Row + 1).Value = VarSplit(1)
TryNext:
    Next
    
End Sub