Sub FormatHyperlink(control As IRibbonControl)
'   Purpose: Converts a range of text hyperlink selected into a working hyperlink
'   Note: Uses built-in hyperlink() function
'   Reference: https://superuser.com/questions/580387/how-to-turn-plain-text-links-into-hyperlinks-in-excel
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim xCell As Range
        
    For Each xCell In Selection
        ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=xCell.Formula
    Next xCell
    
ErrorHandler:
    Exit Sub

End Sub

