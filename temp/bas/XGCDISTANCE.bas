Function XGCDISTANCE(textQuery As Range, varTarget As Range, varDictionary As Range)
'   Purpose: Fill array by criteria
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim result As Variant
    Dim trackResult As Variant
    Dim cLat As Double
    Dim cLon As Double
    Dim i As Single
    
    result = 0
    
    For i = 1 To varDictionary.Cells.Count
        If varDictionary.Cells(i, 1).Value = textQuery Then
            cLat = varDictionary.Cells(i, 2)
            cLon = varDictionary.Cells(i, 3)
            trackResult = XHARVERSINE(varTarget.Cells(1, 1), varTarget.Cells(1, 2), cLat, cLon)
            If result = 0 Then result = trackResult
            If trackResult < result Then result = trackResult
        End If
    Next
    
    XGCDISTANCE = result

End Function

