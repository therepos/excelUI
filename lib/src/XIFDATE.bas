Function XIFDATE(rCell As Range) As String
'   Purpose: Check if it is date
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    XIFDATE = IsDate(rCell)

End Function

