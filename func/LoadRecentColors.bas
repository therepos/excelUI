Sub LoadRecentColors(control As IRibbonControl)
'   Purpose: Use A List Of RGB Codes To Load Colors Into Recent Colors Section of Color Palette
'   Reference: www.TheSpreadsheetGuru.com/the-code-vault
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ColorList As Variant
    Dim CurrentFill As Variant

    ' Array List of RGB Color Codes to Add To Recent Colors Section (Max 10)
    ColorList = Array("248,248,248", "134,188,037", "098,181,229", "000,151,169")

    ' Store ActiveCell's Fill Color (if applicable)
    If ActiveCell.Interior.ColorIndex <> xlNone Then CurrentFill = ActiveCell.Interior.Color

    ' Optimize Code
    Application.ScreenUpdating = False

    ' Loop Through List Of RGB Codes And Add To Recent Colors
    For x = LBound(ColorList) To UBound(ColorList)
        ActiveCell.Interior.Color = RGB(Left(ColorList(x), 3), Mid(ColorList(x), 5, 3), Right(ColorList(x), 3))
        DoEvents
        SendKeys "%hhm~"
    DoEvents
    Next x

    ' Return ActiveCell Original Fill Color
    If CurrentFill = Empty Then
        ActiveCell.Interior.ColorIndex = xlNone
    Else
        ActiveCell.Interior.Color = currentColor
    End If

ErrorHandler:
    Exit Sub

End Sub

