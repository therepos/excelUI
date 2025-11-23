Sub TestUsedRange()
'	Experimental: To use UsedRange instead of looping through all in Selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    With ActiveSheet.UsedRange.Columns(3).Offset(1)
        .Formula = "=IF(ISERROR(MATCH(B2,A:A,0)),"""",B2)"
        .Value = .Value
    End With

ErrorHandler:
    Exit Sub

End Sub