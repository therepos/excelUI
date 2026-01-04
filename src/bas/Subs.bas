Attribute VB_Name = "Subs"
Sub CaseProper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = StrConv(cell, vbProperCase)
    Next cell
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
        
End Sub

Sub CaseSentence()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    
    Application.ScreenUpdating = False
    For Each Rng In XRELEVANTAREA(WorkRng)
        xValue = Rng.Value
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
        Rng.Value = xValue
    Next
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub CaseUpper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = UCase(cell)
    Next cell
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub CellTrim()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = Trim(cell)
    Next cell
    Application.ScreenUpdating = True

ErrorHandler:
    Exit Sub
    
End Sub

Sub ETLSAM()
'   Purpose: Format Identified Misstatements

    Dim rngA, rngB, rngTitle As Range
    Dim lastRow, lastCol As Long
    Dim strEGA As String
    
    lastRow = ActiveSheet.UsedRange.Rows.Count
    lastCol = ActiveSheet.UsedRange.Columns.Count
    
    Set rngTitle = ActiveSheet.Range("A2")
    Set rngA = ActiveSheet.UsedRange
    Set rngB = ActiveSheet.Range(Cells(8, 1), Cells(lastRow, lastCol))
    Set rngAmount = ActiveSheet.Range(Cells(9, lastCol), Cells(lastRow, lastCol))
    
    If Not rngTitle = "Summary of Corrected and Uncorrected Misstatements" Then
        Exit Sub
    End If
    
    'Remove blanks from range
    With rngB
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    'Copy down blank cells
    rngB.Select
    With rngB
        On Error GoTo skiperror
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    End With
    
skiperror:

    'Copy paste as values
    rngB.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Format amount to accounting
    rngAmount.Select
    Call FormatAccounting
    
    'Format workbook to Arial
    Call WorkbookArial
    Call WorkbookPageBreakOff
    
End Sub

Sub ETLTB()
'   Purpose: Format Aura TB Export

    Dim rngA, rngB, rngTitle, rngAccount As Range
    Dim lastRow, lastCol As Long
    Dim strEGA As String
    
    lastRow = ActiveSheet.UsedRange.Rows.Count
    lastCol = ActiveSheet.UsedRange.Columns.Count
    
    Set rngTitle = ActiveSheet.Range("A1")
    Set rngA = ActiveSheet.UsedRange
    Set rngB = ActiveSheet.Range(Cells(2, 1), Cells(lastRow, 3))
    Set rngAccount = ActiveSheet.Range(Cells(2, 5), Cells(lastRow, 5))
    
    'Check if EGA is TB
    If Not rngTitle = "FSLI No." Then
        Exit Sub
    End If
       
    'Remove blank rows
    Call SheetRemoveBlankRows
    
    'Copy down blank cells
    With rngB
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    rngB.Select
    With rngB
        On Error GoTo skiperror
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    End With

skiperror:

    'Copy paste as values
    rngB.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Format amount to accounting
'    rngAmount.Select
'    Call FormatAccounting
    
'    Format workbook to Arial
    Call WorkbookArial
    Call WorkbookPageBreakOff
    
    rngAccount.Select
    Call RemoveBlankRows
    
End Sub

Sub FormatAccounting()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range, rngB As Range
    Set rngSelection = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
        
End Sub

Sub FormatCellRed()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(122, 24, 24)
    Rng.Font.Color = RGB(255, 255, 255)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatDateDDMMM()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlCenter
        c.NumberFormat = "dd-mmm-yy"
    Next c
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatDateDDMMYY()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlCenter
        c.NumberFormat = "dd/mm/yy"
    Next c
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatFontBlue()
    
    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Font.Color = RGB(0, 112, 192)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatFontGreen()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Font.Color = RGB(0, 176, 80)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatFontOrange()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Font.Color = RGB(237, 125, 49)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightGreen()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(204, 285, 204)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightRed()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 204, 204)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightReset()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = xlNone
    Rng.Font.Color = RGB(0, 0, 0)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightYellow()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 255, 0)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatTableBordersGrey()

    On Error GoTo ErrorHandler
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    
ErrorHandler:
    Exit Sub
        
End Sub

Sub FormatTextToValue()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range
    Set rngSelection = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlLeft
        c.NumberFormat = "General"
        c.Value = c.Value
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormulaAbsolute()
'   Purpose: Convert selected values to absolute

    On Error GoTo ErrorHandler

    Dim Rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set Rng = Selection

    Dim regexTargetText As Object
    Set regexTargetText = New RegExp
    With regexTargetText
    .Pattern = "ABS"
    .Global = False
    End With

    Dim c As Range
    For Each c In Rng
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

Sub FormulaReverseSign()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(Rng)
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
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
        
End Sub

Sub FormulaRound()

    On Error GoTo ErrorHandler

    Dim Rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set Rng = Selection

    Dim regexTargetText As Object
    Set regexTargetText = New RegExp
    With regexTargetText
    .Pattern = "ROUNDDOWN"
    .Global = False
    End With

    Dim c As Range
    For Each c In Rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            If regexTargetText.Test(myFormula) Then
                myFormula = regexTargetText.Replace(myFormula, "")
                myFormula = Replace(myFormula, "(", "")
                myFormula = Replace(myFormula, ",0)", "")
                c.Formula = "=" & myFormula
            Else
                c.Formula = "=ROUNDDOWN(" & myFormula & ",0)"
            End If
        Else
            cellValue = c.Value
            c.Formula = "=ROUNDDOWN(" & cellValue & ",0)"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    
ErrorHandler:
    Exit Sub

End Sub

Sub FormulaThousands()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(Rng)
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            c.Formula = "=ROUND(" & myFormula & "/1000,0)"
        Else
            c.Formula = "=ROUND(" & c.Value & "/1000,0)"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
        
End Sub

Sub FormulaToValue()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Copy
    Rng.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertArrowDown()

    On Error GoTo ErrorHandler
    
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
        
    Application.ScreenUpdating = False
    
    With Selection
        X1 = .Left + 10
        Y1 = .Top
    End With
        
    With lCell
        X2 = .Left + 10
        Y2 = .Top + .Height - 1.5
    End With
        
    With ActiveSheet.Shapes
        Set Line1 = .AddLine(X1, Y1, X2, Y2)
        Line1.Line.Weight = 0.5
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
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertBlankSheet()

    ActiveSheet.Select
    Sheets.Add.Name = "SourceData >>>"
    ActiveSheet.Tab.ColorIndex = 1
    
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Intentionally left blank"
    Range("B2").Select
    Selection.Font.Italic = True
    ActiveSheet.Select
    
End Sub

Sub InsertColumnWidth()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In Rng
        c.Formula = "=" & "XCOLUMNWIDTH(" & c.Address & ")"
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0.0_);_((#,##0.0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertCrossReference()

    Call XHYPERACTIVE(Selection)

End Sub

Sub InsertTimestamp()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Value = Now
    Rng.NumberFormat = "dd-mmm-yy"
    Rng.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertWorkdone()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Value = "Keys to Workdone:"
    Rng.Font.Bold = True
    Rng.Offset(1, 0) = "TB"
    Rng.Offset(1, 1) = ": Agreed to current year trial balance."
    Rng.Offset(2, 0) = "PY"
    Rng.Offset(2, 1) = ": Agreed to prior year audited balance."
    Rng.Offset(3, 0) = "imm"
    Rng.Offset(3, 1) = ": Immaterial (below SUM), suggest to leave."
    Rng.Offset(4, 0) = "^"
    Rng.Offset(4, 1) = ": Casted."
    Rng.Offset(5, 0) = "Cal"
    Rng.Offset(5, 1) = ": Calculation checked."
    Rng.Offset(1, 0).Characters(1, 3).Font.Color = RGB(0, 112, 192)
    Rng.Offset(2, 0).Characters(1, 3).Font.Color = RGB(255, 51, 0)
    Rng.Offset(3, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(4, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(5, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(1, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(2, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(3, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(4, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(5, 0).Characters(1, 3).Font.Bold = True
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub RemoveBlankCells()
'   Purpose: Remove blank cells in selection

    On Error GoTo ErrorHandler
    
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

Sub RemoveBlankRows()
'   Purpose: Remove blank rows in selection

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
        
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.EntireRow.Delete
    Else
        If Application.WorksheetFunction.CountA(Selection) = 0 Then
            Selection.EntireRow.Delete
        End If
    End If
    
    Selection.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetColumnsFS()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Columns.ColumnWidth = 11
    Columns("A").ColumnWidth = 1
    Columns("B").ColumnWidth = 45
    Columns("C").ColumnWidth = 5
    Columns("D").ColumnWidth = 11
    Columns("E").ColumnWidth = 11
    Columns("F").ColumnWidth = 11
    Columns("G").ColumnWidth = 11
    Columns("H").ColumnWidth = 11
    Columns("I").ColumnWidth = 11
    Columns("J").ColumnWidth = 11
    Columns("K").ColumnWidth = 11
    Application.ScreenUpdating = True

ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetColumnsNTA4X()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Columns.ColumnWidth = 11
    Columns("A").ColumnWidth = 1
    Columns("B").ColumnWidth = 5
    Columns("C").ColumnWidth = 45
    Columns("D").ColumnWidth = 11
    Columns("E").ColumnWidth = 11
    Columns("F").ColumnWidth = 11
    Columns("G").ColumnWidth = 11
    Columns("H").ColumnWidth = 11
    Columns("I").ColumnWidth = 11
    Columns("J").ColumnWidth = 11
    Columns("K").ColumnWidth = 11
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetColumnsWP()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Columns.ColumnWidth = 12
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 5
    Columns("C").ColumnWidth = 12
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetFontArial()

    On Error GoTo ErrorHandler
    
    If Not ActiveSheet.ProtectContents Then
        ActiveSheet.Cells.Font.Name = "Arial"
        ActiveSheet.Cells.Font.Size = 8
        ActiveSheet.Activate
        ActiveWindow.Zoom = 100
    Else: Exit Sub
    End If
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetFontEY()

    On Error GoTo ErrorHandler
    
    If Not ActiveSheet.ProtectContents Then
        ActiveSheet.Cells.Font.Name = "Georgia"
        ActiveSheet.Cells.Font.Size = 8
        ActiveSheet.Activate
        ActiveWindow.Zoom = 100
    Else: Exit Sub
    End If
        
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetFontSize10()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 10
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetFontSize8()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 8
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100

ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetFormulaToValue()

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    ActiveSheet.Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetPageBreakOff()

    On Error GoTo ErrorHandler
    
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.Activate
    ActiveWindow.DisplayGridlines = False
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetRemoveBlankRows()

    On Error GoTo ErrorHandler
    
    Dim SourceRange As Range
    Dim EntireRow As Range
    Set SourceRange = Application.ActiveSheet.UsedRange
    
    Application.ScreenUpdating = False
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

Sub SheetResetComments()
    
    On Error GoTo ErrorHandler
    
    Dim pComment As Comment
    For Each pComment In Application.ActiveSheet.Comments
       pComment.Shape.Top = pComment.Parent.Top + 5
       pComment.Shape.Left = pComment.Parent.Offset(0, 1).Left + 5
    Next
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetTabBlack()
' Reference for ColorIndex: http://dmcritchie.mvps.org/excel/colors.htm
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 1
    Next ws
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetTabGreen()
' Reference for ColorIndex: http://dmcritchie.mvps.org/excel/colors.htm

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 35
    Next ws
    
ErrorHandler:
    Exit Sub

End Sub

Sub SheetTabRed()
' Reference for ColorIndex: http://dmcritchie.mvps.org/excel/colors.htm
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 38
    Next ws
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetTabReset()
' Reference for ColorIndex: http://dmcritchie.mvps.org/excel/colors.htm
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = xlColorIndexNone
    Next ws
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub SheetTabYellow()
' Reference for ColorIndex: http://dmcritchie.mvps.org/excel/colors.htm
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 6
    Next ws
    
ErrorHandler:
    Exit Sub

End Sub

Sub UnhideSheets()

    On Error GoTo ErrorHandler
    
    For Each sh In Worksheets: sh.Visible = True: Next sh
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub WorkbookArial()
  
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                .Cells.Font.Name = "Arial"
                .Cells.Font.Size = 8
            End If
         End With
    Next ws
    
    For Each ws In Worksheets
        If Not ws.ProtectContents Then
            ws.Activate
            ActiveWindow.Zoom = 100
        End If
    Next
    Application.ScreenUpdating = True
    
    Call sourceSheet.Activate
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub WorkbookEY()
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                .Cells.Font.Name = "EYInterstate Light"
                .Cells.Font.Size = 8
            End If
         End With
    Next ws
    For Each ws In Worksheets
        If Not ws.ProtectContents Then
            ws.Activate
            ActiveWindow.Zoom = 100
        End If
    Next
    Application.ScreenUpdating = True
    
    Call sourceSheet.Activate

ErrorHandler:
    Exit Sub
    
End Sub

Sub WorkbookPageBreakOff()
  
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                ws.DisplayPageBreaks = False
                ws.Activate
                ActiveWindow.DisplayGridlines = False
            End If
         End With
    Next ws
    Application.ScreenUpdating = True
    
    Call sourceSheet.Activate

ErrorHandler:
    Exit Sub
           
End Sub

Sub WorkbookFontSize10()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 10
    Next ws

ErrorHandler:
    Exit Sub

End Sub


