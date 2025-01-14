Sub WorkbookBreakLinks(control As IRibbonControl)
'   Purpose: Break all external links
'   Reference: https://www.extendoffice.com/documents/excel/1173-excel-break-all-links.html
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
        For Each link In wb.LinkSources(xlExcelLinks)
            wb.BreakLink link, xlLinkTypeExcelLinks
        Next link
    End If

'   Alternative approach
'   Purpose: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
'   Source: www.TheSpreadsheetGuru.com/The-Code-Vault

    'Dim ExternalLinks As Variant
    'Dim wb As Workbook
    'Dim x As Long
    '
    'Set wb = ActiveWorkbook
    'ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    'For x = 1 To UBound(ExternalLinks)
    '    wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
    'Next x
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

