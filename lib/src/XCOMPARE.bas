Function XCOMPARE(target As Range, reference As Range) As String
'   Purpose: Return the difference between two cells by words
'   Usage: =XCOMPARE(target cell, reference cell)
'   Examples: =XCOMPARE(cellA, cellB)
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim WordsA As Variant, WordsB As Variant
    Dim ndxA As Long, ndxB As Long, strTemp As String
    
    WordsA = Split(target.text, " ")
    WordsB = Split(reference.text, " ")
    
    For ndxB = LBound(WordsB) To UBound(WordsB)
        For ndxA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(ndxA), WordsB(ndxB), vbTextCompare) = 0 Then
                WordsA(ndxA) = vbNullString
                Exit For
            End If
        Next ndxA
    Next ndxB
    
'   Generates the difference found in range A compared to range B
    For ndxA = LBound(WordsA) To UBound(WordsA)
        If WordsA(ndxA) <> vbNullString Then strTemp = strTemp & WordsA(ndxA) & " "
    Next ndxA
    
    XCOMPARE = Trim(strTemp)

End Function

