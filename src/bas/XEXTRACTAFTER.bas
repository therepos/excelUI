Function XEXTRACTAFTER(rngWord As Range, strWord As String) As String
'   Purpose: Extract the trailing text after a specific word
'   Usage: =XETRACTAFTER(cellA,"word")
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    On Error GoTo ExtractAfter_Error

    Application.Volatile

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim tempResult As String
    
    lngStart = InStr(1, rngWord, strWord)
    If lngStart = 0 Then
        XEXTRACTAFTER = "Not found"
        Exit Function
    End If
    lngEnd = InStr(lngStart + Len(strWord), rngWord, Len(rngWord))

    If lngEnd = 0 Then lngEnd = Len(rngWord)

    tempResult = Mid(rngWord, lngStart + Len(strWord), lngEnd - lngStart)
    XEXTRACTAFTER = Trim(tempResult)
        
    On Error GoTo 0
    Exit Function

ExtractAfter_Error:

    XEXTRACTAFTER = Err.Description

End Function

