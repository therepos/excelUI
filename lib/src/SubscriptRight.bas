Sub SubscriptRight(control As IRibbonControl)
'   Purpose: Subscripts the last character of a text in the selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim cell As Object
    Dim charCount As Variant
    charCount = InputBox("Enter the number of trailing characters to subscript:")
    
    For Each cell In Selection
        cell.Characters(Start:=(Len(cell) - (charCount - 1)), length:=(charCount + 1)).Font.Subscript = True
    Next cell

ErrorHandler:
    Exit Sub

End Sub

