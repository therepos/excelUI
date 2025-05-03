Attribute VB_Name = "Controls"
Sub CaseProper(control As IRibbonControl)
'   Purpose: Set upper case on selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    For Each cell In rng
        cell.Value = StrConv(cell, vbProperCase)
    Next cell
    
ErrorHandler:
    Exit Sub

End Sub

Sub CaseSentence(control As IRibbonControl)
'   Purpose: Set sentence case on selection
'   Reference: KuTools
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    Set WorkRng = Application.Selection
    For Each rng In WorkRng
        xValue = rng.Value
        xStart = True
        For i = 1 To VBA.Len(xValue)
            ch = Mid(xValue, i, 1)
            Select Case ch
                Case "."
                xStart = True
                Case "?"
                xStart = True
                Case "a" To "z"
                If xStart Then
                    ch = UCase(ch)
                    xStart = False
                End If
                Case "A" To "Z"
                If xStart Then
                    xStart = False
                Else
                    ch = LCase(ch)
                End If
            End Select
            Mid(xValue, i, 1) = ch
        Next
        rng.Value = xValue
    Next
    
ErrorHandler:
    Exit Sub

End Sub

Sub CaseUpper(control As IRibbonControl)
'   Purpose: Set upper case on selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    For Each cell In rng
        cell.Value = UCase(cell)
    Next cell
    
ErrorHandler:
    Exit Sub

End Sub

Sub CellTrim(control As IRibbonControl)
'   Purpose: Trim spaces in cell
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
        For Each cell In rng
    cell.Value = Trim(cell)
    Next cell
    
ErrorHandler:
    Exit Sub

End Sub

Sub ColorBordersAll(control As IRibbonControl)
'   Purpose: Change Border Colors without affecting thickness/styles
'   Reference: www.TheSpreadsheetGuru.com/the-code-vault
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim cell As Range
    Dim DesiredColor As Long

'   Color To Change Borders To
    DesiredColor = RGB(198, 198, 198)

'   Ensure Range is Selected
    If TypeName(Selection) <> "Range" Then Exit Sub

'   Loop Through each cell in selection and change border color (if applicable)
    For Each cell In Selection.Cells
        cell.Borders(xlEdgeTop).Color = DesiredColor
        cell.Borders(xlEdgeBottom).Color = DesiredColor
        cell.Borders(xlEdgeLeft).Color = DesiredColor
        cell.Borders(xlEdgeRight).Color = DesiredColor
    Next cell

ErrorHandler:
    Exit Sub

End Sub

Sub ColorBordersOuter(control As IRibbonControl)
'   Purpose: Change Border Colors without affecting thickness/styles
'   Reference: www.TheSpreadsheetGuru.com/the-code-vault
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim cell As Range
    Dim DesiredColor As Long

'   Color To Change Borders To
    DesiredColor = RGB(198, 198, 198)

    Selection.BorderAround , Color:=DesiredColor, Weight:=xlThin

ErrorHandler:
    Exit Sub

End Sub

Sub FontType(control As IRibbonControl)
'   Purpose: Set selected range to Arial
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
'   ===================================
'   Customised use-case
'   ===================================
    Dim userFont As String
    Dim userFontSize As Long
    
    Select Case MySelectedFont
        Case "ddSelectionFont01": userFont = "Arial"
        Case "ddSelectionFont02": userFont = "Verdana"
        Case "ddSelectionFont03": userFont = "Times New Roman"
        Case "": userFont = "Arial"
    End Select
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ===================================

    Dim rng As Range
    Set rng = Selection
    rng.Font.Name = userFont
    rng.Font.Size = userFontSize
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormatAccounting(control As IRibbonControl)
'   Purpose: Set accounting number format on selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlRight
            c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
'       End If
    Next c
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormatColorGreen(control As IRibbonControl)
'   Purpose: To highlight range for follow-up
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Set rng = Selection
    
    rng.Interior.Color = RGB(204, 285, 204)
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormatColorRed(control As IRibbonControl)
'   Purpose: To highlight range for follow-up
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Set rng = Selection
    
    rng.Interior.Color = RGB(255, 204, 204)
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormatDateFull(control As IRibbonControl)
'   Purpose: Set date format on selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlCenter
            c.NumberFormat = "DD MMMM YYYY"
'       End If
    Next c
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormatDateShort(control As IRibbonControl)
'   Purpose: Set date format on selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlCenter
            c.NumberFormat = "dd-mmm-yy"
'       End If
    Next c
    
ErrorHandler:
    Exit Sub

End Sub

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

Sub FormatLineDotted(control As IRibbonControl)
'   Purpose: Insert dotted line
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormatTextToValue(control As IRibbonControl)
'   Purpose: Convert text format to number format on selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlRight
            c.NumberFormat = "General"
            c.Value = c.Value
'       End If
    Next c

ErrorHandler:
    Exit Sub

End Sub

Sub FormulaAbsolute(control As IRibbonControl)
'   Purpose: Convert selected values to absolute
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set rng = Selection

    Dim regexTargetText As Object
    Set regexTargetText = New RegExp
    With regexTargetText
    .Pattern = "ABS"
    .Global = False
    End With

    Dim c As Range
    For Each c In rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            If regexTargetText.Test(myFormula) Then
                myFormula = regexTargetText.Replace(myFormula, "")
                myFormula = Replace(myFormula, "(", "")
                myFormula = Replace(myFormula, ")", "")
                c.Formula = "=" & myFormula
            Else
                c.Formula = "=ABS(" & myFormula & ")"
            End If
        Else
            cellValue = c.Value
            c.Formula = "=ABS(" & cellValue & ")"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormulaAbsoluteReference(control As IRibbonControl)
'   Purpose: Absolute reference selected cells
'   Reference: http://www.excelforum.com/excel-general/372383-making-multiple-cells-absolute-at-once.html
'   Reference: http://www.contextures.com/xlvba01.html#videoreg
'   Todo: Check if cell formula is already referenced.
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim cell As Range
    
    For Each cell In Selection
        If cell.HasFormula Then
            cell.Formula = _
            Application.ConvertFormula(cell.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next
 
ErrorHandler:
    Exit Sub

End Sub

Sub FormulaReverseSign(control As IRibbonControl)
'   Purpose: Reverse the sign of selected range
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    For Each c In rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            If Left(myFormula, 1) = "-" Then
                c.Formula = "=" & Right(myFormula, Len(myFormula) - 1)
            Else
                c.Formula = "=-" & myFormula
            End If
        Else
                c.Formula = "=-" & c.Value
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c

ErrorHandler:
    Exit Sub

End Sub

Sub FormulaRound(control As IRibbonControl)
'   Purpose: Convert selected values to absolute
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    For Each c In rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            c.Formula = "=ROUND(" & myFormula & ",0)"
        Else
            c.Formula = "=ROUND(" & c.Value & ",0)"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c

ErrorHandler:
    Exit Sub

End Sub

Sub FormulaToValue(control As IRibbonControl)
'   Purpose: Convert selected formulas to values
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    
ErrorHandler:
    Exit Sub

End Sub

Sub InsertArrowDown(control As IRibbonControl)
'   Purpose: Draw line down
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim X1 As Long
    Dim X2 As Long
    Dim Y1 As Long
    Dim Y2 As Long
    
    Dim Line1 As Shape
    
    Dim mX1 As Long
    Dim mY1 As Long
    Dim mX2 As Long
    Dim mY2 As Long
    
    Dim Line2 As Shape
    
    Dim lCell As Range
    
    Set lCell = Selection.Cells(Selection.Rows.Count, Selection.Columns.Count) 'Last Cell
        
    With Selection 'First cell
'   Original code places the arrow in the middle of the selection
'   X1 = .Left + .Width / 2
        X1 = .Left + 10
        Y1 = .Top
    End With
        
    With lCell
'   Original code places the arrow in the middle of the selection
'   X2 = .Left + .Width / 2
        X2 = .Left + 10
        Y2 = .Top + .Height - 1.5
    End With
        
    With ActiveSheet.Shapes
'   Get the return value and create the line.
        Set Line1 = .AddLine(X1, Y1, X2, Y2)
        Line1.Line.Weight = 1
        Line1.Line.BeginArrowheadStyle = msoArrowheadNone
        Line1.Line.EndArrowheadStyle = msoArrowheadTriangle
        Line1.Line.EndArrowheadWidth = msoArrowheadWidthMedium
        Line1.Line.EndArrowheadLength = msoArrowheadLengthMedium
        Line1.Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    With lCell
        mX1 = .Left + .Width / 2 - 4
        mX2 = .Left + .Width / 2 + 4
        mY1 = .Top + .Height - 1
        mY2 = .Top + .Height - 1
    End With
    
'    With ActiveSheet.Shapes
'        Set Line2 = .AddLine(mX1, mY1, mX2, mY2)
'        Line2.Line.Weight = 1
'        Line2.Line.ForeColor.RGB = RGB(0, 0, 255)
'    End With

ErrorHandler:
    Exit Sub

End Sub

Sub InsertColumnWidth(control As IRibbonControl)
'   Purpose: Insert column width counter
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    For Each c In rng
        c.Formula = "=" & "XCOLUMNWIDTH(" & c.Address & ")"
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0.0_);_((#,##0.0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub InsertCrossReference(control As IRibbonControl)
'   Purpose: Create hyperlink based on targeted cell
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Call XHYPERACTIVE(Selection)
    
ErrorHandler:
    Exit Sub

End Sub

Private Function XHYPERACTIVE(ByRef rng As Range)
'   Purpose: To create hyperlink based on selected cell
'   Note: Passive function to be activated by InsertCrossReference()

    Dim strAddress As String
    Dim target As Range

    On Error Resume Next
        Set target = Application.InputBox( _
          Title:="Create Hyperlink", _
          Prompt:="Select a cell to create hyperlink", _
          Type:=8)
    On Error GoTo 0
  
'   Ensure User did not cancel
    If target Is Nothing Then Exit Function
  
'   Set Variable to first cell in user's input (ensuring only 1 cell)
    Set target = target.Cells(1, 1)
    
'   Get the text value of the address to display as hyperlink TextToDisplay
    strAddress = target.Parent.Name & "!" & target.Address(External:=False)

'   Generate hyperlink
    With ActiveSheet.Hyperlinks
    .Add Anchor:=rng, _
         Address:="", _
         SubAddress:="=" & strAddress, _
         TextToDisplay:=strAddress
    End With
    
End Function

Sub InsertHeadingAudit(control As IRibbonControl)
'   Purpose: Insert customised headings for audit workpapers
'   Note: Utilises CCH Engagement functions
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Dim myClient As String
    Dim myYear As String
    Set rng = Selection
    
    myClient = "=UPPER(PJNAME())"

    myYear = "=" & Chr(34) & "FINANCIAL YEAR ENDED " & Chr(34) & "&"
    myYear = myYear & "UPPER(TEXT(" & "CYEDATE()" & "," & Chr(34) & "dd mmmm yyyy" & Chr(34)
    myYear = myYear & "))"

    If rng.HasFormula = True Then
        rng.Formula = Replace(rng.Formula, rng.Formula, myClient)
    Else
        rng.Formula = "=1"
        rng.Formula = Replace(rng.Formula, rng.Formula, myClient)
    End If
    
    If rng.Offset(1, 0).HasFormula = True Then
        rng.Offset(1, 0).Formula = Replace(rng.Offset(1, 0).Formula, rng.Offset(1, 0).Formula, myYear)
    Else
        rng.Offset(1, 0).Formula = "=1"
        rng.Offset(1, 0).Formula = Replace(rng.Offset(1, 0).Formula, rng.Offset(1, 0).Formula, myYear)
    End If
    
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    rng.Offset(1, 0).Copy
    rng.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    
    rng.Font.Bold = True
    rng.Offset(1, 0).Font.Bold = True
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub InsertHeadingTax(control As IRibbonControl)
'   Purpose: Insert customised headings for tax workpapers
'   Note: Utilises CCH Engagement functions
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim rng As Range
    Dim myClient As String
    Dim myYear As String
    Set rng = Selection
    
    myClient = "=UPPER(PJNAME())"

    myYear = "=" & Chr(34) & "YEAR OF ASSESSMENT " & Chr(34) & "&"
    myYear = myYear & "UPPER(TEXT(" & "CYBDATE()+365*2" & "," & Chr(34) & "yyyy" & Chr(34)
    myYear = myYear & "))"

    If rng.HasFormula = True Then
        rng.Formula = Replace(rng.Formula, rng.Formula, myClient)
    Else
        rng.Formula = "=1"
        rng.Formula = Replace(rng.Formula, rng.Formula, myClient)
    End If
    
    If rng.Offset(1, 0).HasFormula = True Then
        rng.Offset(1, 0).Formula = Replace(rng.Offset(1, 0).Formula, rng.Offset(1, 0).Formula, myYear)
    Else
        rng.Offset(1, 0).Formula = "=1"
        rng.Offset(1, 0).Formula = Replace(rng.Offset(1, 0).Formula, rng.Offset(1, 0).Formula, myYear)
    End If
    
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    rng.Offset(1, 0).Copy
    rng.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    
    rng.Font.Bold = True
    rng.Offset(1, 0).Font.Bold = True
    
    Application.ScreenUpdating = True
        
ErrorHandler:
    Exit Sub

End Sub

Sub InsertWorkdone(control As IRibbonControl)
'   Purpose: Insert customised legend for workdone
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    
    rng.Value = "Legend:"
    rng.Font.Bold = True
    rng.Offset(1, 0) = "TB: Agreed to current year trial balance."
    rng.Offset(2, 0) = "PY: Agreed to prior year audited balance."
    rng.Offset(3, 0) = "i: Immaterial (below CTT), suggest to leave."
    rng.Offset(4, 0) = "GL: Agreed to current year general ledger."
    rng.Offset(1, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(2, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(3, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(4, 0).Characters(1, 2).Font.Color = RGB(255, 51, 0)
    rng.Offset(1, 0).Characters(1, 2).Font.Bold = True
    rng.Offset(2, 0).Characters(1, 2).Font.Bold = True
    rng.Offset(3, 0).Characters(1, 2).Font.Bold = True
    rng.Offset(4, 0).Characters(1, 2).Font.Bold = True
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

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

Sub RemoveBlankCells(control As IRibbonControl)
'   Purpose: Remove blank cells in selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
        
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
    Else
        If Application.WorksheetFunction.CountA(Selection) = 0 Then
            Selection.Delete Shift:=xlUp
        End If
    End If
    
    Selection.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub RemoveBlankRows(control As IRibbonControl)
'   Purpose: Remove blank rows in selection
'   Reference: https://www.wallstreetmojo.com/vba-last-row/
'   Reference: https://www.rondebruin.nl/win/s9/win005.htm
'   Notes:
'   - Selection.SpecialCells(xlCellTypeLastCell).Row    Return the last used in the worksheet regardless of selection
'   - Selection.Find Method                             Returns the last used based on selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
        
    Dim SourceRange As Range
    Dim TargetRow As Range
    Dim lastRow As Long
    Dim firstRow As Long
    
    Set SourceRange = Selection
    firstRow = SourceRange.Cells(1).row
    
    If XLASTUSEDROW(SourceRange) > 0 Then
        lastRow = XLASTUSEDROW(SourceRange)
    Else
        lastRow = SourceRange.row + SourceRange.Rows.Count - 1
    End If
    
    For i = lastRow To firstRow Step -1
        Set TargetRow = Cells(i, 1).EntireRow
        If Application.WorksheetFunction.CountA(TargetRow) = 0 Then
            TargetRow.Delete
        End If
    Next
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Private Function XLASTUSEDROW(rng As Range) As Long
'   Purpose: Find last used row
'   Reference: https://www.rondebruin.nl/win/s9/win005.htm
'   Note: Returns 0 if not found

    Dim result As Long

    On Error Resume Next
    result = rng.Find(What:="*", _
               After:=rng.Cells(1), _
               Lookat:=xlPart, _
               LookIn:=xlFormulas, _
               SearchOrder:=xlByRows, _
               SearchDirection:=xlPrevious, _
               MatchCase:=False).row
                
    XLASTUSEDROW = result
    If Err.Number <> 0 Then
        XLASTUSEDROW = 0
    End If
         
End Function

Private Function XLASTUSEDCOL(rng As Range) As Long
'   Purpose: Find last column
'   Reference: https://www.rondebruin.nl/win/s9/win005.htm

    Dim result As Long
          
    On Error Resume Next
    result = rng.Find(What:="*", _
                After:=rng.Cells(1), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Column
                
    XLASTUSEDCOL = result
    If Err.Number <> 0 Then
        XLASTUSEDCOL = rng.Column + rng.Columns.Count - 1
    End If
         
End Function

Sub RemoveNamedRanges(control As IRibbonControl)
'   Purpose: Delete all named ranges
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim i As Long
    
    Application.Calculation = xlCalculationManual
    For i = ThisWorkbook.Names.Count To 1 Step -1
        ThisWorkbook.Names(i).Delete
    Next
    Application.Calculation = xlCalculationAutomatic

ErrorHandler:
    Exit Sub

End Sub

Sub RevertFile(control As IRibbonControl)
'   Purpose: Revert macro changes
'   Reference: https://www.excelforum.com/excel-programming-vba-macros/491103-undoing-a-macro.html
'   Reference: https://stackoverflow.com/questions/33813806/is-it-possible-to-undo-a-macro-action#:~:text=1)%20Have%20the%20macro%20save,did%20whatever%20the%20macro%20does.
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler

    wkname = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    ActiveWorkbook.Close SaveChanges:=False
    
    Workbooks.Open FileName:=wkname

ErrorHandler:
    Exit Sub

End Sub

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

Sub SheetColumnsFS(control As IRibbonControl)
'   Purpose: Standardise columns width for specific worksheet: BS/PL tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.ColumnWidth = 14
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 1
    Columns("C").ColumnWidth = 28
    Columns("D").ColumnWidth = 4
    Columns("E").ColumnWidth = 11
    Columns("F").ColumnWidth = 1
    Columns("G").ColumnWidth = 11
    Columns("H").ColumnWidth = 1
    Columns("I").ColumnWidth = 11
    Columns("J").ColumnWidth = 1
    Columns("K").ColumnWidth = 11
    Columns("L").ColumnWidth = 1
    Columns("M").ColumnWidth = 1
    Columns("N").ColumnWidth = 1
    Columns("M").Interior.Color = RGB(217, 217, 217)
    Columns("A:L").Font.Name = "Times New Roman"
    Columns("A:L").Font.Size = 10
    Range("B1").Formula = "=XCOLUMNWIDTH(B1)"
    Range("C1").Formula = "=XCOLUMNWIDTH(C1)"
    Range("D1").Formula = "=XCOLUMNWIDTH(D1)"
    Range("E1").Formula = "=XCOLUMNWIDTH(E1)"
    Range("F1").Formula = "=XCOLUMNWIDTH(F1)"
    Range("G1").Formula = "=XCOLUMNWIDTH(G1)"
    Range("H1").Formula = "=XCOLUMNWIDTH(H1)"
    Range("I1").Formula = "=XCOLUMNWIDTH(I1)"
    Range("J1").Formula = "=XCOLUMNWIDTH(J1)"
    Range("K1").Formula = "=XCOLUMNWIDTH(K1)"
    Range("O1").Formula = "=SUM(B1:K1)"
    Range("B1").HorizontalAlignment = xlCenter
    Range("C1").HorizontalAlignment = xlCenter
    Range("D1").HorizontalAlignment = xlCenter
    Range("E1").HorizontalAlignment = xlCenter
    Range("F1").HorizontalAlignment = xlCenter
    Range("G1").HorizontalAlignment = xlCenter
    Range("H1").HorizontalAlignment = xlCenter
    Range("I1").HorizontalAlignment = xlCenter
    Range("J1").HorizontalAlignment = xlCenter
    Range("K1").HorizontalAlignment = xlCenter
    Range("O1").HorizontalAlignment = xlLeft
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetColumnsNTA4X(control As IRibbonControl)
'   Purpose: Standardise columns width for specific worksheet: NTA 4-columns tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.ColumnWidth = 14
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 1
    Columns("C").ColumnWidth = 14
    Columns("D").ColumnWidth = 1
    Columns("E").ColumnWidth = 13
    Columns("F").ColumnWidth = 1
    Columns("G").ColumnWidth = 10
    Columns("H").ColumnWidth = 1
    Columns("I").ColumnWidth = 10
    Columns("J").ColumnWidth = 1
    Columns("K").ColumnWidth = 10
    Columns("L").ColumnWidth = 1
    Columns("M").ColumnWidth = 10
    Columns("N").ColumnWidth = 1
    Columns("O").ColumnWidth = 1
    Columns("P").ColumnWidth = 1
    Columns("O").Interior.Color = RGB(217, 217, 217)
    Columns("A:N").Font.Name = "Times New Roman"
    Columns("A:N").Font.Size = 10
    Range("B1").Formula = "=XCOLUMNWIDTH(B1)"
    Range("C1").Formula = "=XCOLUMNWIDTH(C1)"
    Range("D1").Formula = "=XCOLUMNWIDTH(D1)"
    Range("E1").Formula = "=XCOLUMNWIDTH(E1)"
    Range("F1").Formula = "=XCOLUMNWIDTH(F1)"
    Range("G1").Formula = "=XCOLUMNWIDTH(G1)"
    Range("H1").Formula = "=XCOLUMNWIDTH(H1)"
    Range("I1").Formula = "=XCOLUMNWIDTH(I1)"
    Range("J1").Formula = "=XCOLUMNWIDTH(J1)"
    Range("K1").Formula = "=XCOLUMNWIDTH(K1)"
    Range("L1").Formula = "=XCOLUMNWIDTH(L1)"
    Range("M1").Formula = "=XCOLUMNWIDTH(M1)"
    Range("Q1").Formula = "=SUM(B1:M1)"
    Range("B1").HorizontalAlignment = xlCenter
    Range("C1").HorizontalAlignment = xlCenter
    Range("D1").HorizontalAlignment = xlCenter
    Range("E1").HorizontalAlignment = xlCenter
    Range("F1").HorizontalAlignment = xlCenter
    Range("G1").HorizontalAlignment = xlCenter
    Range("H1").HorizontalAlignment = xlCenter
    Range("I1").HorizontalAlignment = xlCenter
    Range("J1").HorizontalAlignment = xlCenter
    Range("K1").HorizontalAlignment = xlCenter
    Range("L1").HorizontalAlignment = xlCenter
    Range("M1").HorizontalAlignment = xlCenter
    Range("Q1").HorizontalAlignment = xlLeft

ErrorHandler:
    Exit Sub

End Sub

Sub SheetColumnsNTA6X(control As IRibbonControl)
'   Purpose: Standardise columns width for specific worksheet: NTA 6-columns tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.ColumnWidth = 14
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 3
    Columns("C").ColumnWidth = 10
    Columns("D").ColumnWidth = 15
    Columns("E").ColumnWidth = 11
    Columns("F").ColumnWidth = 1
    Columns("G").ColumnWidth = 5
    Columns("H").ColumnWidth = 1
    Columns("I").ColumnWidth = 5
    Columns("J").ColumnWidth = 1
    Columns("K").ColumnWidth = 10
    Columns("L").ColumnWidth = 1
    Columns("M").ColumnWidth = 10
    Columns("N").ColumnWidth = 1
    Columns("O").ColumnWidth = 1
    Columns("P").ColumnWidth = 1
    Columns("O").Interior.Color = RGB(217, 217, 217)
    Columns("A:N").Font.Name = "Times New Roman"
    Columns("A:N").Font.Size = 10
    Range("B1").Formula = "=XCOLUMNWIDTH(B1)"
    Range("C1").Formula = "=XCOLUMNWIDTH(C1)"
    Range("D1").Formula = "=XCOLUMNWIDTH(D1)"
    Range("E1").Formula = "=XCOLUMNWIDTH(E1)"
    Range("F1").Formula = "=XCOLUMNWIDTH(F1)"
    Range("G1").Formula = "=XCOLUMNWIDTH(G1)"
    Range("H1").Formula = "=XCOLUMNWIDTH(H1)"
    Range("I1").Formula = "=XCOLUMNWIDTH(I1)"
    Range("J1").Formula = "=XCOLUMNWIDTH(J1)"
    Range("K1").Formula = "=XCOLUMNWIDTH(K1)"
    Range("L1").Formula = "=XCOLUMNWIDTH(L1)"
    Range("M1").Formula = "=XCOLUMNWIDTH(M1)"
    Range("Q1").Formula = "=SUM(B1:M1)"
    Range("B1").HorizontalAlignment = xlCenter
    Range("C1").HorizontalAlignment = xlCenter
    Range("D1").HorizontalAlignment = xlCenter
    Range("E1").HorizontalAlignment = xlCenter
    Range("F1").HorizontalAlignment = xlCenter
    Range("G1").HorizontalAlignment = xlCenter
    Range("H1").HorizontalAlignment = xlCenter
    Range("I1").HorizontalAlignment = xlCenter
    Range("J1").HorizontalAlignment = xlCenter
    Range("K1").HorizontalAlignment = xlCenter
    Range("L1").HorizontalAlignment = xlCenter
    Range("M1").HorizontalAlignment = xlCenter
    Range("Q1").HorizontalAlignment = xlLeft

ErrorHandler:
    Exit Sub

End Sub

Sub SheetColumnsTickmark(control As IRibbonControl)
'   Purpose: Standardise columns width for specific worksheet: Tickmark tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.ColumnWidth = 15
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 1
    Columns("C").ColumnWidth = 3
    Columns("D").ColumnWidth = 15
    Columns("E").ColumnWidth = 15
    Columns("F").ColumnWidth = 15
    Columns("G").ColumnWidth = 15
    Columns("H").ColumnWidth = 15
    Columns("I").ColumnWidth = 15
    Columns("J").ColumnWidth = 15
    Columns("K").ColumnWidth = 15
    Columns("L").ColumnWidth = 15
    Columns("M").ColumnWidth = 1
    Columns("N").ColumnWidth = 5
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetColumnsWP(control As IRibbonControl)
'   Purpose: Standardise workbook columns width
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    For Each ws In Worksheets
        Columns.ColumnWidth = 14
        Columns("A").ColumnWidth = 1
        Columns("B").ColumnWidth = 3
        Columns("C").ColumnWidth = 5
    Next ws
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetRemoveBlankRows(control As IRibbonControl)
'   Purpose: Remove blank rows in sheet
'   Reference: https://www.ablebits.com/office-addins-blog/2018/12/19/delete-blank-lines-excel/
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
        
    Dim SourceRange As Range
    Dim EntireRow As Range
 
    Set SourceRange = Application.ActiveSheet.UsedRange
 
    If Not (SourceRange Is Nothing) Then
        For i = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
    End If
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetResetComments(control As IRibbonControl)
'   Purpose: Reset position of comments
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
    
    Dim pComment As Comment
    For Each pComment In Application.ActiveSheet.Comments
       pComment.Shape.Top = pComment.Parent.Top + 5
       pComment.Shape.Left = pComment.Parent.Offset(0, 1).Left + 5
    Next
    
ErrorHandler:
    Exit Sub

End Sub

Sub SortRight(control As IRibbonControl)
'   Purpose: Sort a series of numbers from left to right
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim row As Range
    
    For Each row In Selection.Rows
        row.Sort Key1:=row, Order1:=xlAscending, Orientation:=xlSortRows
    Next row

ErrorHandler:
    Exit Sub

End Sub

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

Sub SheetFont(control As IRibbonControl)
'   Purpose: Standardise worksheet font type and size
'   Updated: 2022MAR12

'   Saves worksheet before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
       
'   ===================================
'   Customised use-case
'   ===================================
    Dim userFont As String
    Dim userFontSize As Long
    
    Select Case MySelectedFont
        Case "ddSelectionFont01": userFont = "Arial"
        Case "ddSelectionFont02": userFont = "Verdana"
        Case "ddSelectionFont03": userFont = "Times New Roman"
        Case "": userFont = "Arial"
    End Select
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ==================================
    
    With ActiveSheet
       .Cells.Font.Name = userFont
       .Cells.Font.Size = userFontSize
    End With
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetFontSize(control As IRibbonControl)
'   Purpose: Standardise workbook font size
'   Updated: 2022MAR12

'   Saves worksheet before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

'   ===================================
'   Customised use-case
'   ===================================
    Dim userFontSize As Long
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ==================================

    With ActiveSheet
       .Cells.Font.Size = userFontSize
    End With
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetFormulaToValue(control As IRibbonControl)
'   Purpose: Convert all worksheet formulas to values (most efficient way)
'   Updated: 2022MAR12

'   Saves worksheet before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
     
    ActiveSheet.Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub WorkbookFont(control As IRibbonControl)
'   Purpose: Standardise workbook font type and size
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
       
'   ===================================
'   Customised use-case
'   ===================================
    Dim userFont As String
    Dim userFontSize As Long
    
    Select Case MySelectedFont
        Case "ddSelectionFont01": userFont = "Arial"
        Case "ddSelectionFont02": userFont = "Verdana"
        Case "ddSelectionFont03": userFont = "Times New Roman"
        Case "": userFont = "Arial"
    End Select
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ==================================
    
    Dim ws As Worksheet
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = userFont
            .Cells.Font.Size = userFontSize
         End With
    Next ws
    For Each ws In Worksheets
        ws.Activate
        ActiveWindow.Zoom = 100
    Next

ErrorHandler:
    Exit Sub
    
    Application.ScreenUpdating = True
    
End Sub

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

Sub WorkbookFontSize(control As IRibbonControl)
'   Purpose: Standardise workbook font size
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

'   ===================================
'   Customised use-case
'   ===================================
    Dim userFontSize As Long
        
    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": userFontSize = 8
        Case "ddSelectionFontSize02": userFontSize = 9
        Case "ddSelectionFontSize03": userFontSize = 10
        Case "ddSelectionFontSize04": userFontSize = 11
        Case "": userFontSize = 10
    End Select
'   ==================================

    Dim ws As Worksheet
    For Each ws In Worksheets
         With ws
            .Cells.Font.Size = userFontSize
         End With
    Next ws

    For Each ws In Worksheets
        ws.Activate
        ActiveWindow.Zoom = 100
    Next
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub WorkbookFormulaToValue(control As IRibbonControl)
'   Purpose: Convert all workbook formulas to values (most efficient way)
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim sh As Worksheet, HidShts As New Collection

    For Each sh In ActiveWorkbook.Worksheets
        If Not sh.visible Then
            HidShts.Add sh
            sh.visible = xlSheetVisible
        End If
    Next sh
     
    Worksheets.Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
     
    For Each sh In HidShts
        sh.visible = xlSheetHidden
    Next sh
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetUnhideAll(control As IRibbonControl)
'   Purpose: Unhide all rows and columns
'   Reference: https://www.extendoffice.com/documents/excel/1173-excel-break-all-links.html
'   Updated: 2022MAR12

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    ActiveSheet.Columns.EntireColumn.Hidden = False
    ActiveSheet.Rows.EntireRow.Hidden = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetPageBreakOff(control As IRibbonControl)
'   Purpose: This removes all page breaks for worksheet
'   Reference: www.DedicatedExcel.com
'   Updated: 2022MAR12

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.Activate
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Sub WorkbookPageBreakOff(control As IRibbonControl)
'   Purpose: This removes all page breaks for all worksheets in the workbook
'   Reference: www.DedicatedExcel.com
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
 
    For Each ws In Sheets
        ws.DisplayPageBreaks = False
    Next ws
 
     For Each ws In Sheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws
          
ErrorHandler:
    Exit Sub

End Sub

Sub WorkbookResetStyles(control As IRibbonControl)
'   Purpose: Removes all styles in the active workbook and rebuild the default styles
'   Reference: https://support.microsoft.com/en-us/topic/how-to-programmatically-reset-a-workbook-to-default-styles-36e94af7-d185-68fb-3962-0882a5260132
'   Note: Rebuilds the default styles by merging them from a new workbook
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

'   Dimension variables
    Dim MyBook As Workbook
    Dim tempBook As Workbook
    Dim CurStyle As Style

'   Set MyBook to the active workbook
    Set MyBook = ActiveWorkbook
    On Error Resume Next
'   Delete all the styles in the workbook
    For Each CurStyle In MyBook.Styles
        'If CurStyle.Name <> "Normal" Then CurStyle.Delete
        Select Case CurStyle.Name
            Case "20% - Accent1", "20% - Accent2", _
            "20% - Accent3", "20% - Accent4", "20% - Accent5", "20% - Accent6", _
            "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4", _
            "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", _
            "60% - Accent3", "60% - Accent4", "60% - Accent5", "60% - Accent6", _
            "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", _
            "Bad", "Calculation", "Check Cell", "Comma", "Comma [0]", "Currency", _
            "Currency [0]", "Explanatory Text", "Good", "Heading 1", "Heading 2", _
            "Heading 3", "Heading 4", "Input", "Linked Cell", "Neutral", "Normal", _
            "Note", "Output", "Percent", "Title", "Total", "Warning Text"
            'Do nothing, these are the default styles
            Case Else
            CurStyle.Delete
        End Select
    Next CurStyle

'   Open a new workbook.
'   Disable alerts so you may merge changes to the Normal style from the new workbook.
'   Merge styles from the new workbook into the existing workbook.
'   Enable alerts.
'   Close the new workbook.

    Set tempBook = Workbooks.Add
    Application.DisplayAlerts = False
    MyBook.Styles.Merge Workbook:=tempBook
    Application.DisplayAlerts = True
    tempBook.Close

ErrorHandler:
    Exit Sub

End Sub

Sub WorkbookResetTabColors(control As IRibbonControl)
'   Purpose: Reset all tab colors
'   Reference: https://www.extendoffice.com/documents/excel/5179-excel-remove-tab-color.html
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim xSheet As Worksheet
    
    For Each xSheet In ActiveWorkbook.Worksheets
        xSheet.Tab.ColorIndex = xlColorIndexNone
    Next xSheet
    
ErrorHandler:
    Exit Sub

End Sub

Sub WorkbookUnhideAll(control As IRibbonControl)
'   Purpose: Unhide all rows and columns
'   Reference: https://www.extendoffice.com/documents/excel/1173-excel-break-all-links.html
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Columns.EntireColumn.Hidden = False
        ws.Rows.EntireRow.Hidden = False
    Next ws
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

Function XCELLFORMULA(rCell As Range) As String
'   Purpose: Return cell formula
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save

    XCELLFORMULA = rCell.Formula

End Function

Function XCLEANTEXT(text As String)
'   Purpose: Removes excess non-alphanumeric characters
'   Usage: =LEN(B3)-LEN(SUBSTITUTE(B3,C3,))
'   Feature: To count number of delimiter
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save

'   Remove leading and training symbols
    text = REMOVESYMBOLS(text, LEAD)
    text = REMOVESYMBOLS(text, TRAIL)

'   Replace hanging double quotation marks
    Dim m As Integer

    For m = 1 To Len(text)
        If Mid(text, m, 1) = Chr(34) Then
            If m = 1 Then
                If Not IsNumeric(Mid(text, m - 1, 1)) Then
                    text = Left(text, m - 1) & Right(text, Len(text) - m)
                    m = m - 1
                End If
            Else
                text = REMOVESYMBOLS(text, LEAD)
            End If
        End If
    Next m
'   Double spacing
    text = WorksheetFunction.Substitute(text, "  ", "")
'   Comma
    text = WorksheetFunction.Substitute(text, ",", "")
    XCLEANTEXT = Trim(text)

End Function

Function XCOLUMNWIDTH(target As Range) As Double
'   Purpose: Get column width
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    
    XCOLUMNWIDTH = target.ColumnWidth
    Application.ScreenUpdating = True
    
End Function

Function XCOMPARE(target As Range, reference As Range) As String
'   Purpose: Return the difference between two cells by words
'   Usage: =XCOMPARE(target cell, reference cell)
'   Examples: =XCOMPARE(cellA, cellB)
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim WordsA As Variant, WordsB As Variant
    Dim ndxA As Long, ndxB As Long, strTemp As String
    
    WordsA = Split(target.text, " ")
    WordsB = Split(reference.text, " ")
    
    For ndxB = LBound(WordsB) To UBound(WordsB)
        For ndxA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(ndxA), WordsB(ndxB), vbTextCompare) = 0 Then
                WordsA(ndxA) = vbNullString
                Exit For
            End If
        Next ndxA
    Next ndxB
    
'   Generates the difference found in range A compared to range B
    For ndxA = LBound(WordsA) To UBound(WordsA)
        If WordsA(ndxA) <> vbNullString Then strTemp = strTemp & WordsA(ndxA) & " "
    Next ndxA
    
    XCOMPARE = Trim(strTemp)

End Function

Function XEXTRACTAFTER(rngWord As Range, strWord As String) As String
'   Purpose: Extract the trailing text after a specific word
'   Usage: =XETRACTAFTER(cellA,"word")
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    On Error GoTo ExtractAfter_Error

    Application.Volatile

    Dim lngStart As Long
    Dim lngEnd As Long
    Dim tempResult As String
    
    lngStart = InStr(1, rngWord, strWord)
    If lngStart = 0 Then
        XEXTRACTAFTER = "Not found"
        Exit Function
    End If
    lngEnd = InStr(lngStart + Len(strWord), rngWord, Len(rngWord))

    If lngEnd = 0 Then lngEnd = Len(rngWord)

    tempResult = Mid(rngWord, lngStart + Len(strWord), lngEnd - lngStart)
    XEXTRACTAFTER = Trim(tempResult)
        
    On Error GoTo 0
    Exit Function

ExtractAfter_Error:

    XEXTRACTAFTER = Err.Description

End Function

Function XEXTRACTBEFORE(rngWord As Range, strWord As String) As String
'   Purpose: Extract the leading text before a specific word
'   Usage: =XETRACTBEFORE(cellA,"word")
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    On Error GoTo ExtractBefore_Error

    Application.Volatile

    Dim lngStart        As Long
    Dim lngEnd          As Long
    Dim tempResult      As String

    lngEnd = InStr(1, rngWord, strWord)
    If lngEnd = 0 Then
        XEXTRACTBEFORE = "Not found"
        Exit Function
    End If
    lngStart = 1

    tempResult = Left(rngWord, lngEnd - 1)
    XEXTRACTBEFORE = Trim(tempResult)

    On Error GoTo 0
    Exit Function

ExtractBefore_Error:

    XEXTRACTBEFORE = Err.Description

End Function

Function XFIND(text As Range, wordList As Range)
'   Purpose: Return the matched words from text description based on a word list
'   Requirement: Microsoft VBScript Regular Expressions 5.5
'   Usage: =XFIND(text, wordlist)
'   Examples: =XFIND(cellA, cellB1:B5)
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim strPattern As String
    Dim regEx As New RegExp
    Dim result As String
    Dim m As Integer
    Dim i As Single
    
    m = 0

    For i = 1 To wordList.Cells.Count
        strPattern = "\b" & wordList.Cells(i).Value & "\b"
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If regEx.Test(text) Then
            If m = 0 Then
                result = wordList.Cells(i).Value
                m = m + 1
            Else
                result = result & " " & wordList.Cells(i).Value
                m = m + 1
            End If
        End If
    Next
    
'   Starts of polymorphic test
'   Purpose: To use the same function for string search as well
'   ==========================
'
'    strPattern = "\b" & wordList & "\b"
'    With regEx
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = strPattern
'    End With
'
'    If regEx.Test(text) Then
'        If m = 0 Then
'            result = wordList
'            m = m + 1
'        Else
'            result = result & " " & wordList
'            m = m + 1
'        End If
'    End If

'   Ends of polymorphic test
    
    If m = 0 Then
        result = "Not found"
    End If
    
    XFIND = result

End Function

Function XGCDISTANCE(textQuery As Range, varTarget As Range, varDictionary As Range)
'   Purpose: Fill array by criteria
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim result As Variant
    Dim trackResult As Variant
    Dim cLat As Double
    Dim cLon As Double
    Dim i As Single
    
    result = 0
    
    For i = 1 To varDictionary.Cells.Count
        If varDictionary.Cells(i, 1).Value = textQuery Then
            cLat = varDictionary.Cells(i, 2)
            cLon = varDictionary.Cells(i, 3)
            trackResult = XHARVERSINE(varTarget.Cells(1, 1), varTarget.Cells(1, 2), cLat, cLon)
            If result = 0 Then result = trackResult
            If trackResult < result Then result = trackResult
        End If
    Next
    
    XGCDISTANCE = result

End Function

Private Function XHARVERSINE(Lat1 As Variant, Lon1 As Variant, Lat2 As Variant, Lon2 As Variant)
 '  Purpose: Great Circle Distance calculation
 '  Note: Returns results in kilometers

    Dim R As Integer, dlon As Variant, dlat As Variant, Rad1 As Variant
    Dim A As Variant, c As Variant, d As Variant, Rad2 As Variant

    R = 6371
    dlon = Excel.WorksheetFunction.Radians(Lon2 - Lon1)
    dlat = Excel.WorksheetFunction.Radians(Lat2 - Lat1)
    Rad1 = Excel.WorksheetFunction.Radians(Lat1)
    Rad2 = Excel.WorksheetFunction.Radians(Lat2)
    A = Sin(dlat / 2) * Sin(dlat / 2) + Cos(Rad1) * Cos(Rad2) * Sin(dlon / 2) * Sin(dlon / 2)
    c = 2 * Excel.WorksheetFunction.Atan2(Sqr(1 - A), Sqr(A))
    d = R * c
    XHARVERSINE = d
    
End Function

Function XGETPAGENUMBER(CurrentCell As Range) As String
'   Purpose: Return page number of a cell
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim VPC As Integer, HPC As Integer
    Dim VerticalPageBreak As VPageBreak, HorizontalPageBreak As HPageBreak
    Dim NumPage As Integer

    If ActiveSheet.PageSetup.Order = xlDownThenOver Then
        HPC = ActiveSheet.HPageBreaks.Count + 1
        VPC = 1
    Else
        VPC = ActiveSheet.VPageBreaks.Count + 1
        HPC = 1
    End If

    NumPage = 1
    For Each VerticalPageBreak In ActiveSheet.VPageBreaks
      If VerticalPageBreak.Location.Column > CurrentCell.Cells.Column Then Exit For
      NumPage = NumPage + HPC
    Next VerticalPageBreak
    
    For Each HorizontalPageBreak In ActiveSheet.HPageBreaks
      If HorizontalPageBreak.Location.row > CurrentCell.Cells.row Then Exit For
      NumPage = NumPage + VPC
    Next HorizontalPageBreak

    XGETPAGENUMBER = NumPage

End Function

Function XHASNUMBER(target As Range)
'   Purpose: Check if there are numbers in a text
'   Usage: =XHASNUMBER(target cell)
'   Alternative: =COUNT(FIND({0,1,2,3,4,5,6,7,8,9},A1))
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim targetStr As String, length As Long, i As Long
    XHASNUMBER = False
    targetStr = target.text
    length = Len(targetStr)

    For i = 1 To length
        If IsNumeric(Mid(targetStr, i, 1)) Then
            XHASNUMBER = True
        End If
    Next i
    
End Function

Function XIFDATE(rCell As Range) As String
'   Purpose: Check if it is date
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    XIFDATE = IsDate(rCell)

End Function

Function XLOOKUP(text As Variant, targetList As Range, resultList As Variant, Optional errResult As Variant)
'   Purpose: Custom XLOOKUP
'   Usage 01: =XLOOKUP(A1, A1:A10, B1:B10)
'   Usage 02: =XLOOKUP(A1, A1:A10, "True", "False")
'   Reference: https://stackoverflow.com/questions/44638867/vba-excel-try-catch
'   Reference: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
'   Reference: https://stackoverflow.com/questions/32008841/best-way-to-return-error-in-udf-vba-function
'   Reference: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement
'   Todo: resultList unable to accept cell value
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = False
    On Error GoTo XLOOKUP_Error
    
    If TypeName(resultList) = "Range" Then
        XLOOKUP = WorksheetFunction.index(resultList, WorksheetFunction.Match(text, targetList, 0))
    Else
        If IsError(WorksheetFunction.Match(text, targetList, 0)) Then
            GoTo XLOOKUP_Error
        Else
            XLOOKUP = resultList
        End If
    End If
    
    Application.ScreenUpdating = True
    Exit Function
    
XLOOKUP_Error:

    If IsMissing(errResult) Then
'   Substituted xlErrValue to xlErrName
        XLOOKUP = CVErr(xlErrName)
    Else
        XLOOKUP = errResult
    End If
    Resume Next
    
End Function

Function XREMOVEBETWEEN(ByVal str As String, oldStart As String, oldEnd As String) As String
'   Purpose:  Remove text between delimiter
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
'   Check syntax
    While InStr(str, oldStart) = 0 And InStr(str, oldEnd) > InStr(str, oldStart)
        str = Left(str, InStr(str, oldStart) - 1) & Mid(str, InStr(str, oldEnd) + 1)
    Wend
  
    XREMOVEBETWEEN = Trim(str)
  
End Function

Function XREMOVESYMBOLS(text As String, opType As String, Optional charCount)
'   Purpose: Substitute a leading symbol
'   Feature: By default leading 3 characters are considered if not specified
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim varSymbols() As Variant
    Dim oldText As String
    Dim n As Integer

    varSymbols = Array(Chr(34), "-")
    
    Select Case opType
        Case "Full"
            For n = 0 To 1
                oldText = varSymbols(n)
                text = WorksheetFunction.Substitute(text, oldText, "")
            Next n
        Case "LEAD"
            For n = 0 To 1
                oldText = varSymbols(n)
                text = SUBSTITUTELEADING(text, oldText, "", charCount)
            Next n
        Case "TRAIL"
             For n = 0 To 1
                oldText = varSymbols(n)
                text = SUBSTITUTETRAILING(text, oldText, "", charCount)
            Next n
    End Select

    XREMOVESYMBOLS = Trim(text)

End Function

Function XREPLACEWORDS(strSource As String, strFind As Range, strReplace As Range)
'   Purpose: Replace strictly words in a text with boundary based on wordlists
'   Usage: =XREPLACEWORDS(targetcell, searchlist, replacelist)
'   Example: =XREPLACEWORDS(A1, B1:B5, C1:C5)
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim strPattern As String
    Dim regEx As New RegExp
    Dim result As String
    Dim i As Single
    
    For i = 1 To strFind.Cells.Count
        strPattern = b & strFind.Cells(i).Value & b
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        strSource = regEx.Replace(strSource, strReplace.Cells(i).Value)

    Next i
    XREPLACEWORDS = Trim(strSource)
    
End Function

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

Function XSPELLNUMBER(ByVal MyNumber)
'   Purpose: Spell numbers as dollars
'   Source: https://support.microsoft.com/en-us/office/convert-numbers-into-words-a0d166fb-e1ea-4090-95c8-69442cd55d98
'   Reference: https://stackoverflow.com/questions/11155912/how-to-make-vba-function-vba-only-and-disable-it-as-udf/41130822
'   Modification:
'   - Made private XGETHUNDREDS, XGETTENS. Users do not need to access these.
'   - Added "Only" to the end of result.
'   - Fixed additional spacing between words in the original code.
'   - Fixed 0.XX appearing as "No Dollar and XX Cents" to "XX Cents Only"
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "

'   String representation of amount
    MyNumber = Trim(str(MyNumber))
    
'   Position of decimal place 0 if none
    DecimalPlace = InStr(MyNumber, ".")
    
'   Convert cents and set MyNumber to dollar amount
    If DecimalPlace > 0 Then
        Cents = XGETTENS(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    Count = 1
    Do While MyNumber <> ""
        Temp = XGETHUNDREDS(Right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    
    Select Case Dollars
        Case "": Dollars = "" ' Originally "No Dollar"
        Case "One": Dollars = "One Dollar"
        Case Else
        Dollars = Dollars & " Dollars"
    End Select
    
    Select Case Cents
        Case "": Cents = " Only"
        Case "One": Cents = " and One Cent Only"
        Case Else
        Cents = " and " & Cents & " Cents Only"
    End Select
    
    If Dollars = "" Then
        XSPELLNUMBER = Trim(Replace(Replace(Dollars & Cents, "  ", " "), "and", ""))
    Else
        XSPELLNUMBER = Trim(Replace(Dollars & Cents, "  ", " "))
    End If
    Application.ScreenUpdating = True

End Function
       
Private Function XGETHUNDREDS(ByVal MyNumber)
'   Purpose: Converts a number from 100-999 into text

    Dim result As String
    
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    
'   Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        result = XGETDIGIT(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    
'   Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        result = result & XGETTENS(Mid(MyNumber, 2))
    Else
        result = result & XGETDIGIT(Mid(MyNumber, 3))
    End If
    
    XGETHUNDREDS = result

End Function
    
Private Function XGETTENS(TensText)
'   Purpose: Converts a number from 10 to 99 into text.
 
    Dim result As String
    
'   Null out the temporary function value.
    result = ""
    
'   If value between 10-19...
    If Val(Left(TensText, 1)) = 1 Then
        Select Case Val(TensText)
            Case 10: result = "Ten"
            Case 11: result = "Eleven"
            Case 12: result = "Twelve"
            Case 13: result = "Thirteen"
            Case 14: result = "Fourteen"
            Case 15: result = "Fifteen"
            Case 16: result = "Sixteen"
            Case 17: result = "Seventeen"
            Case 18: result = "Eighteen"
            Case 19: result = "Nineteen"
            Case Else
        End Select
    Else
'   If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: result = "Twenty"
            Case 3: result = "Thirty"
            Case 4: result = "Forty"
            Case 5: result = "Fifty"
            Case 6: result = "Sixty"
            Case 7: result = "Seventy"
            Case 8: result = "Eighty"
            Case 9: result = "Ninety"
            Case Else
        End Select
'   Retrieve ones place.
        result = result & " " & XGETDIGIT(Right(TensText, 1))
    End If
  
    XGETTENS = result
 
End Function
  
Private Function XGETDIGIT(Digit)
'   Purpose: Converts a number from 1 to 9 into text.

    Select Case Val(Digit)
        Case 1: XGETDIGIT = "One"
        Case 2: XGETDIGIT = "Two"
        Case 3: XGETDIGIT = "Three"
        Case 4: XGETDIGIT = "Four"
        Case 5: XGETDIGIT = "Five"
        Case 6: XGETDIGIT = "Six"
        Case 7: XGETDIGIT = "Seven"
        Case 8: XGETDIGIT = "Eight"
        Case 9: XGETDIGIT = "Nine"
        Case Else: XGETDIGIT = ""
    End Select
    
End Function

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

Function XSUBSTITUTEPREFIX(text As String, oldText As String, newText As String, Optional charCount)
'   Purpose: Substitute a leading symbol
'   Feature: By default leading 3 characters are considered if not specified
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim i As Integer
    Dim leadText As String
    Dim trailText As String
    
    If IsMissing(charCount) Then
        i = 3
    Else
        i = charCount
    End If

    If Len(text) = i Then
        i = Len(text)
    End If

    leadText = Left(text, i)
    trailText = Right(text, Len(text) - i)

    leadText = WorksheetFunction.Substitute(leadText, oldText, newText)
    text = leadText & trailText

    XSUBSTITUTEPREFIX = Trim(text)

End Function

Function XSUBSTITUTESUFFIX(text As String, oldText As String, newText As String, Optional charCount)
'   Purpose: Substitute a trailing symbol
'   Feature: By default leading 3 characters are considered if not specified
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
        
    Dim i As Integer
    Dim leadText As String
    Dim trailText As String
    
    If IsMissing(charCount) Then
        i = 3
    Else
        i = charCount
    End If

    If Len(text) = i Then
        i = Len(text)
    End If

    trailText = Right(text, i)
    leadText = Left(text, Len(text) - i)

    trailText = WorksheetFunction.Substitute(trailText, oldText, newText)
    text = leadText & trailText

    XSUBSTITUTESUFFIX = Trim(text)

End Function

Function XTRANSLATE(strInput As String, strFromSourceLanguage As String, strToTargetLanguage As String) As String
'   Purpose: Translate with Google Translate
'   Reference: https://www.youtube.com/watch?v=RsyqqzholVk&ab_channel=DineshKumarTakyar
'   Usage: = Translate(Range("A2"), "en", "es")
'   Todo: Create optional source language input. Default to auto detect language
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save

    Dim strURL As String
    Dim objHTTP As Object
    Dim objHTML As Object
    Dim objDivs As Object, objDiv As Object
    Dim strTranslated As String

    ' send query to web page (google translate mobile)
    strURL = "https://translate.google.com/m?hl=" & strFromSourceLanguage & _
        "&sl=" & strFromSourceLanguage & _
        "&tl=" & strToTargetLanguage & _
        "&ie=UTF-8&prev=_m&q=" & strInput

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP") 'late binding
    objHTTP.Open "GET", strURL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ""

    ' create an html document
    Set objHTML = CreateObject("htmlfile")
    With objHTML
        .Open
        .Write objHTTP.responsetext
        .Close
    End With

    Set objDivs = objHTML.getElementsByTagName("div")

    For Each objDiv In objDivs

        If objDiv.className = "result-container" Then
            strTranslated = objDiv.innerText
            Translate = strTranslated
        End If

    Next objDiv

    Set objHTML = Nothing
    Set objHTTP = Nothing
    
End Function

