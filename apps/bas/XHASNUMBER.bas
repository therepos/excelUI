Function XHASNUMBER(target As Range)
'   Purpose: Check if there are numbers in a text
'   Usage: =XHASNUMBER(target cell)
'   Alternative: =COUNT(FIND({0,1,2,3,4,5,6,7,8,9},A1))
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim targetStr As String, length As Long, i As Long
    XHASNUMBER = False
    targetStr = target.text
    length = Len(targetStr)

    For i = 1 To length
        If IsNumeric(Mid(targetStr, i, 1)) Then
            XHASNUMBER = True
        End If
    Next i
    
End Function

