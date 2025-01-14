Function XREPLACEWORDS(strSource As String, strFind As Range, strReplace As Range)
'   Purpose: Replace strictly words in a text with boundary based on wordlists
'   Usage: =XREPLACEWORDS(targetcell, searchlist, replacelist)
'   Example: =XREPLACEWORDS(A1, B1:B5, C1:C5)
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim strPattern As String
    Dim regEx As New RegExp
    Dim result As String
    Dim i As Single
    
    For i = 1 To strFind.Cells.Count
        strPattern = b & strFind.Cells(i).Value & b
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        strSource = regEx.Replace(strSource, strReplace.Cells(i).Value)

    Next i
    XREPLACEWORDS = Trim(strSource)
    
End Function

