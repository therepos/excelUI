Sub InsertPageHeader()
'	Purpose: Customise headers for audit purposes
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    For Each ws In Worksheets
        ws.PageSetup.DifferentFirstPageHeaderFooter = True
        ws.PageSetup.FirstPage.LeftHeader.text = _
        "&""Arial,Bold""&12" & " " & UCase(wb.BuiltinDocumentProperties("Company").Value) & vbCr & _
        "&""Arial,Bold""&11" & " " & UCase(wb.BuiltinDocumentProperties("Title").Value) & vbCr & _
        "&""Arial,Bold""&11" & " " & UCase(wb.BuiltinDocumentProperties("Subject").Value)
        ws.PageSetup.FirstPage.RightHeader.text = _
        "&""Arial,Bold""&12" & UCase("&A") & vbCr & _
        "&""Arial,Regular""&11" & "Preparer:" & "&K00+000AAAAAAAAAAA&K000000" & vbCr & _
        "&""Arial,Regular""&11" & "Reviewer:" & "&K00+000AAAAAAAAAAA&K000000"
        ws.PageSetup.RightHeader = _
        "&""Arial,Bold""&12" & UCase("&A/&P-1 ")
    Next ws

ErrorHandler:
    Exit Sub

End Sub