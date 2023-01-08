Function XEXTRACTBEFORE(rngWord As Range, strWord As String) As String
'   Purpose: Extract the leading text before a specific word
'   Usage: =XETRACTBEFORE(cellA,"word")
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    On Error GoTo ExtractBefore_Error

    Application.Volatile

    Dim lngStart        As Long
    Dim lngEnd          As Long
    Dim tempResult      As String

    lngEnd = InStr(1, rngWord, strWord)
    If lngEnd = 0 Then
        XEXTRACTBEFORE = "Not found"
        Exit Function
    End If
    lngStart = 1

    tempResult = Left(rngWord, lngEnd - 1)
    XEXTRACTBEFORE = Trim(tempResult)

    On Error GoTo 0
    Exit Function

ExtractBefore_Error:

    XEXTRACTBEFORE = Err.Description

End Function

