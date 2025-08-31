Attribute VB_Name = "Module1"
' VSCode extension: VBScript symbols by Andreas Lenzen

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

Sub WorkbookGeorgia()
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                .Cells.Font.Name = "Georgia"
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

Sub InsertWorkdone()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Value = "Keys to Workdone:"
    rng.Font.Bold = True
    rng.Offset(1, 0) = "TB"
    rng.Offset(1, 1) = ": Agreed to current year trial balance."
    rng.Offset(2, 0) = "PY"
    rng.Offset(2, 1) = ": Agreed to prior year audited balance."
    rng.Offset(3, 0) = "imm"
    rng.Offset(3, 1) = ": Immaterial (below SUM), suggest to leave."
    rng.Offset(4, 0) = "^"
    rng.Offset(4, 1) = ": Casted."
    rng.Offset(5, 0) = "Cal"
    rng.Offset(5, 1) = ": Calculation checked."
    rng.Offset(1, 0).Characters(1, 3).Font.Color = RGB(0, 112, 192)
    rng.Offset(2, 0).Characters(1, 3).Font.Color = RGB(255, 51, 0)
    rng.Offset(3, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    rng.Offset(4, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    rng.Offset(5, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    rng.Offset(1, 0).Characters(1, 3).Font.Bold = True
    rng.Offset(2, 0).Characters(1, 3).Font.Bold = True
    rng.Offset(3, 0).Characters(1, 3).Font.Bold = True
    rng.Offset(4, 0).Characters(1, 3).Font.Bold = True
    rng.Offset(5, 0).Characters(1, 3).Font.Bold = True
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertTimestamp()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Value = Now
    rng.NumberFormat = "dd-mmm-yy"
    rng.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertColumnWidth()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    Application.ScreenUpdating = False
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
    
    Set lCell = Selection.Cells(Selection.Rows.count, Selection.Columns.count) 'Last Cell
        
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

Sub CellTrim()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    For Each Cell In XRELEVANTAREA(rng)
        Cell.Value = Trim(Cell)
    Next Cell
    Application.ScreenUpdating = True

ErrorHandler:
    Exit Sub
    
End Sub

Sub CaseUpper()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    For Each Cell In XRELEVANTAREA(rng)
        Cell.Value = UCase(Cell)
    Next Cell
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub CaseProper()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    For Each Cell In XRELEVANTAREA(rng)
        Cell.Value = StrConv(Cell, vbProperCase)
    Next Cell
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
        
End Sub

Sub CaseSentence()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    
    Application.ScreenUpdating = False
    For Each rng In XRELEVANTAREA(WorkRng)
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
    Application.ScreenUpdating = True
    
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
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Font.Color = RGB(0, 112, 192)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatFontGreen()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Font.Color = RGB(0, 176, 80)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatFontOrange()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Font.Color = RGB(237, 125, 49)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatCellRed()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Interior.Color = RGB(122, 24, 24)
    rng.Font.Color = RGB(255, 255, 255)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightRed()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Interior.Color = RGB(255, 204, 204)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightGreen()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Interior.Color = RGB(204, 285, 204)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightYellow()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Interior.Color = RGB(255, 255, 0)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormatHighlightReset()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Interior.Color = xlNone
    rng.Font.Color = RGB(0, 0, 0)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormulaRound()

    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set rng = Selection

    Dim regexTargetText As Object
    Set regexTargetText = New RegExp
    With regexTargetText
    .Pattern = "ROUNDDOWN"
    .Global = False
    End With

    Dim c As Range
    For Each c In rng
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
    
    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(rng)
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

Sub FormulaAbsolute()
'   Purpose: Convert selected values to absolute

    On Error GoTo ErrorHandler

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

Sub FormulaToValue()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub FormulaReverseSign()

    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(rng)
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

Sub RemoveBlankRows()
'   Purpose: Remove blank rows in selection

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
        
    If Selection.Cells.count > 1 Then
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

Sub RemoveBlankCells()
'   Purpose: Remove blank cells in selection

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
        
    If Selection.Cells.count > 1 Then
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

Sub SheetFontGeorgia()

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

Sub SheetFontSize8()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 8
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100

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

Sub SheetRemoveBlankRows()

    On Error GoTo ErrorHandler
    
    Dim SourceRange As Range
    Dim EntireRow As Range
    Set SourceRange = Application.ActiveSheet.UsedRange
    
    Application.ScreenUpdating = False
    If Not (SourceRange Is Nothing) Then
        For i = SourceRange.Rows.count To 1 Step -1
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

Sub SheetPageBreakOff()

    On Error GoTo ErrorHandler
    
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.Activate
    ActiveWindow.DisplayGridlines = False
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub InsertCrossReference()

    Call XHYPERACTIVE(Selection)

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

Sub UnhideSheets()

    On Error GoTo ErrorHandler
    
    For Each sh In Worksheets: sh.Visible = True: Next sh
    
ErrorHandler:
    Exit Sub
    
End Sub

Sub ETLSAM()
'   Purpose: Format Identified Misstatements

    Dim rngA, rngB, rngTitle As Range
    Dim lastRow, lastCol As Long
    Dim strEGA As String
    
    lastRow = ActiveSheet.UsedRange.Rows.count
    lastCol = ActiveSheet.UsedRange.Columns.count
    
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
    
    lastRow = ActiveSheet.UsedRange.Rows.count
    lastCol = ActiveSheet.UsedRange.Columns.count
    
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

Function XCOLUMNWIDTH(target As Range) As Double

    Application.ScreenUpdating = False
    XCOLUMNWIDTH = target.ColumnWidth
    Application.ScreenUpdating = True
        
End Function

Function XGETBOLD(pWorkRng As Range)

    If pWorkRng.Font.Bold Then
        XGETBOLD = True
    Else
        XGETBOLD = False
    End If
    
End Function

Function XGETINDENTLEVEL(targetCell As Range)

    XGETINDENTLEVEL = targetCell.IndentLevel

End Function

Private Function XHYPERACTIVE(ByRef rng As Range)

    Dim strAddress, strTextDisplay As String
    Dim target As Range

    Application.DisplayAlerts = False
    On Error Resume Next
    Set target = Application.InputBox( _
      Title:="Create Hyperlink", _
      Prompt:="Select a cell to create hyperlink", _
      Type:=8)
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    If rng Is Nothing Then
        Exit Function
    Else
        strAddress = Chr(39) & target.Parent.Name & Chr(39) & "!" & target.Address
        If WorksheetFunction.CountA(rng) = 0 Then
            strTextDisplay = target.Parent.Name
        Else
            strTextDisplay = rng.Value
        End If
        
        With ActiveSheet.Hyperlinks
        .Add Anchor:=rng, _
             Address:="", _
             SubAddress:=strAddress, _
             TextToDisplay:=strTextDisplay
        End With
    End If
    
ErrorHandler:
    Exit Function
    
End Function

Private Function XFIRSTUSEDROW(rng As Range) As Long

    Dim result As Long

    On Error Resume Next
    If IsEmpty(rng.Cells(1)) Then
        result = rng.Find(What:="*", _
               After:=rng.Cells(1), _
               Lookat:=xlPart, _
               LookIn:=xlFormulas, _
               SearchOrder:=xlByRows, _
               SearchDirection:=xlNext, _
               MatchCase:=False).Row
    Else: result = rng.Cells(1).Row
    End If
    XFIRSTUSEDROW = result
    If Err.Number <> 0 Then
        XFIRSTUSEDROW = 0
    End If
         
End Function

Private Function XLASTUSEDROW(rng As Range) As Long

    Dim result As Long

    On Error Resume Next
    result = rng.Find(What:="*", _
               After:=rng.Cells(1), _
               Lookat:=xlPart, _
               LookIn:=xlFormulas, _
               SearchOrder:=xlByRows, _
               SearchDirection:=xlPrevious, _
               MatchCase:=False).Row
                
    XLASTUSEDROW = result
    If Err.Number <> 0 Then
        XLASTUSEDROW = 0
    End If
         
End Function

Private Function XFIRSTUSEDCOL(rng As Range) As Long

    Dim result As Long
          
    On Error Resume Next
    result = rng.Find(What:="*", _
                After:=rng.Cells(1), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, _
                MatchCase:=False).Column
                
    XFIRSTUSEDCOL = result
    If Err.Number <> 0 Then
        XFIRSTUSEDCOL = rng.Column + rng.Columns.count - 1
    End If
         
End Function

Private Function XLASTUSEDCOL(rng As Range) As Long

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
        XLASTUSEDCOL = rng.Column + rng.Columns.count - 1
    End If
         
End Function

Private Function XRELEVANTAREA(rngTarget As Range) As Range

    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    
    firstRow = XFIRSTUSEDROW(rngTarget)
    firstCol = XFIRSTUSEDCOL(rngTarget)
    lastRow = XLASTUSEDROW(rngTarget)
    lastCol = XLASTUSEDCOL(rngTarget)
    Set XRELEVANTAREA = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))
    
    On Error GoTo ErrorHandler
ErrorHandler:
    Exit Function
        
End Function

Function XEXTRACTAFTER(rngWord As Range, strWord As String) As String
'   Purpose: Extract the trailing text after a specific word
'   Usage: =XETRACTAFTER(cellA,"word")
    
    On Error GoTo ErrorHandler

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
    
ErrorHandler:

    XEXTRACTAFTER = Err.Description

End Function

Function XEXTRACTBEFORE(rngWord As Range, strWord As String) As String
'   Purpose: Extract the leading text before a specific word
'   Usage: =XETRACTBEFORE(cellA,"word")
    
    On Error GoTo ErrorHandler

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

ErrorHandler:

    XEXTRACTBEFORE = Err.Description

End Function

Function XCOUNTCOLOR(CountRange As Range, CountColor As Range)

    Dim CountColorValue As Integer
    Dim TotalCount As Integer
    CountColorValue = CountColor.Interior.ColorIndex
    Set rCell = CountRange
    For Each rCell In CountRange
      If rCell.Interior.ColorIndex = CountColorValue Then
        TotalCount = TotalCount + 1
      End If
    Next rCell
    XCOUNTCOLOR = TotalCount
    
End Function

' ======================================
' Experimental
' ======================================

Function GenerateLotteryNumbers(excludeNumbers As Variant, sumLower As Integer, sumUpper As Integer, avgLower As Integer, avgUpper As Integer) As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim arr(1 To 7) As Integer
    Dim i As Integer, j As Integer
    Dim temp As Integer
    Dim sum As Integer
    Dim avg As Double
    
    Randomize
    
    Do
        ' Generate 7 distinct digits from 1 to 49 excluding specified numbers
        For i = 1 To 7
            temp = Int(Rnd() * 49) + 1
            While IsDuplicate(temp, arr) Or IsExcluded(temp, excludeNumbers)
                temp = Int(Rnd() * 49) + 1
            Wend
            arr(i) = temp
        Next i
        
        ' Sort all 7 digits in ascending order
        For i = 1 To 6
            For j = i + 1 To 7
                If arr(j) < arr(i) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next j
        Next i
        
        ' Check if the sum of all 7 digits is between 190 and 250
        sum = WorksheetFunction.sum(arr)
        
        ' Check if the average of all 7 digits is between 20 and 40
        avg = WorksheetFunction.Average(arr)
        
    Loop Until sum >= sumLower And sum <= sumUpper And avg >= avgLower And avg <= avgUpper
    
    ' Return the result as an array of 7 numbers
    GenerateLotteryNumbers = arr
    
End Function

' Function to check if a number already exists in an array
Function IsDuplicate(ByVal n As Integer, arr() As Integer) As Boolean
    Dim i As Integer
    For i = 1 To UBound(arr)
        If arr(i) = n Then
            IsDuplicate = True
            Exit Function
        End If
    Next i
    IsDuplicate = False
End Function

' Function to check if a number is in the list of excluded numbers
Function IsExcluded(ByVal n As Integer, excludeNumbers As Variant) As Boolean
    Dim i As Integer
    For i = LBound(excludeNumbers) To UBound(excludeNumbers)
        If n = excludeNumbers(i) Then
            IsExcluded = True
            Exit Function
        End If
    Next i
    IsExcluded = False
End Function

Function GetStyle(rng As Range)
    GetStyle = rng.Style
End Function
