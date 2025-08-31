Function XREMOVEBETWEEN(ByVal str As String, oldStart As String, oldEnd As String) As String
'   Purpose:  Remove text between delimiter
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
'   Check syntax
    While InStr(str, oldStart) = 0 And InStr(str, oldEnd) > InStr(str, oldStart)
        str = Left(str, InStr(str, oldStart) - 1) & Mid(str, InStr(str, oldEnd) + 1)
    Wend
  
    XREMOVEBETWEEN = Trim(str)
  
End Function

