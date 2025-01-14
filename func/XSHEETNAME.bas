Function XSHEETNAME(rCell As Range, Optional UseAsRef As Boolean) As String
'   Purpose: Return sheet name of a cell
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Application.Volatile
    If UseAsRef = True Then
        XSHEETNAME = "'" & rCell.Parent.Name & "'!"
    Else
        XSHEETNAME = rCell.Parent.Name
    End If

End Function

