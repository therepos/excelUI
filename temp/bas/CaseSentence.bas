Sub CaseSentence(control As IRibbonControl)
'   Purpose: Set sentence case on selection
'   Reference: KuTools
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    Set WorkRng = Application.Selection
    For Each rng In WorkRng
        xValue = rng.Value
        xStart = True
        For i = 1 To VBA.Len(xValue)
            ch = Mid(xValue, i, 1)
            Select Case ch
                Case "."
                xStart = True
                Case "?"
                xStart = True
                Case "a" To "z"
                If xStart Then
                    ch = UCase(ch)
                    xStart = False
                End If
                Case "A" To "Z"
                If xStart Then
                    xStart = False
                Else
                    ch = LCase(ch)
                End If
            End Select
            Mid(xValue, i, 1) = ch
        Next
        rng.Value = xValue
    Next
    
ErrorHandler:
    Exit Sub

End Sub

