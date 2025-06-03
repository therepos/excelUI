Sub SetPrintMargin(control As IRibbonControl)
'   Purpose: Set print margins
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Worksheets(i).Activate
        With Worksheets(i).PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .FirstPageNumber = 0
            .PrintGridlines = True
            .CenterHorizontally = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .DifferentFirstPageHeaderFooter = True
            .LeftMargin = Application.CentimetersToPoints(1.5)
            .RightMargin = Application.CentimetersToPoints(0.5)
            .TopMargin = Application.CentimetersToPoints(1)
            .BottomMargin = Application.CentimetersToPoints(1)
            .HeaderMargin = Application.CentimetersToPoints(0.7)
            .FooterMargin = Application.CentimetersToPoints(0.7)
            .FirstPage.RightFooter.text = "&A"
            .RightFooter = "&A" & " - " & "&P"
        End With
    Next i
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

