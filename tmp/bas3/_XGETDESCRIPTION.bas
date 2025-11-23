Public Function XGETDESCRIPTION(leadText As Range, trailText As Range, Optional brandText As String)
'	Purpose: Print full item description from parts
'	Format: (a) Item Type Measurement
'	Format: (b) Item Type Measurement, Brand
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim result As String
    
    If IsMissing(brandText) Then
        result = leadText & " " & trailText
    Else
        result = leadText & " " & trailText & "," & brandText
    End If

    XGETDESCRIPTION = Trim(UCase(result))
    
End Function