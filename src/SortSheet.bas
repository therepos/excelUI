Sub SortSheet(control As IRibbonControl)
'   Purpose: Sort worksheets alphabetically
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
        
    Dim i As Integer
    Dim j As Integer
   
    For i = 1 To Sheets.Count
        For j = 1 To Sheets.Count - 1
            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
                Sheets(j).Move After:=Sheets(j + 1)
            End If
        Next j
    Next i
    
ErrorHandler:
    Exit Sub

End Sub

