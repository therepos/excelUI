Function XFIND(text As Range, wordList As Range)
'   Purpose: Return the matched words from text description based on a word list
'   Requirement: Microsoft VBScript Regular Expressions 5.5
'   Usage: =XFIND(text, wordlist)
'   Examples: =XFIND(cellA, cellB1:B5)
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim strPattern As String
    Dim regEx As New RegExp
    Dim result As String
    Dim m As Integer
    Dim i As Single
    
    m = 0

    For i = 1 To wordList.Cells.Count
        strPattern = "\b" & wordList.Cells(i).Value & "\b"
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If regEx.Test(text) Then
            If m = 0 Then
                result = wordList.Cells(i).Value
                m = m + 1
            Else
                result = result & " " & wordList.Cells(i).Value
                m = m + 1
            End If
        End If
    Next
    
'   Starts of polymorphic test
'   Purpose: To use the same function for string search as well
'   ==========================
'
'    strPattern = "\b" & wordList & "\b"
'    With regEx
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = strPattern
'    End With
'
'    If regEx.Test(text) Then
'        If m = 0 Then
'            result = wordList
'            m = m + 1
'        Else
'            result = result & " " & wordList
'            m = m + 1
'        End If
'    End If

'   Ends of polymorphic test
    
    If m = 0 Then
        result = "Not found"
    End If
    
    XFIND = result

End Function

