Function XSUBSTITUTEMULTIPLE(text As String, old_text As Range, new_text As Variant)
'   Purpose: Substitute multiple values, including symbols, in a text without boundary
'   Usage: =XSUBSTITUTEMULTIPLE(A1, B1:B10, C1:C10)
'   Usage: =XSUBSTITUTEMULTIPLE(A1, B1:B10, C1)
'   Feature: Faster than REPLACEWORDS
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim i As Single

    For i = 1 To old_text.Cells.Count
        If TypeName(new_text) = Range Then
            text = WorksheetFunction.Substitute(text, old_text.Cells(i).Value, new_text.Cells(i).Value)
        Else
            text = WorksheetFunction.Substitute(text, old_text.Cells(i).Value, new_text)
        End If
    Next i

    XSUBSTITUTEMULTIPLE = Trim(text)

End Function

