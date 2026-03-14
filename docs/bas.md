# VBA modules

_This file is generated automatically from `.bas` files in `src/bas`._

## Module `Functions`

### `XFIRSTUSEDCOL`

```vbnet
Public Function XFIRSTUSEDCOL(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    result = Rng.Find(What:="*", _
                      After:=Rng.Cells(1), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlNext, _
                      MatchCase:=False).Column
    XFIRSTUSEDCOL = result
    If Err.Number <> 0 Then XFIRSTUSEDCOL = Rng.Column + Rng.Columns.Count - 1
End Function
```

### `XLASTUSEDCOL`

```vbnet
Public Function XLASTUSEDCOL(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    result = Rng.Find(What:="*", _
                      After:=Rng.Cells(1), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Column
    XLASTUSEDCOL = result
    If Err.Number <> 0 Then XLASTUSEDCOL = Rng.Column + Rng.Columns.Count - 1
End Function
```

### `XFIRSTUSEDROW`

```vbnet
Public Function XFIRSTUSEDROW(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    If IsEmpty(Rng.Cells(1)) Then
        result = Rng.Find(What:="*", _
                          After:=Rng.Cells(1), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False).Row
    Else
        result = Rng.Cells(1).Row
    End If
    XFIRSTUSEDROW = result
    If Err.Number <> 0 Then XFIRSTUSEDROW = 0
End Function
```

### `XLASTUSEDROW`

```vbnet
Public Function XLASTUSEDROW(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    result = Rng.Find(What:="*", _
                      After:=Rng.Cells(1), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Row
    XLASTUSEDROW = result
    If Err.Number <> 0 Then XLASTUSEDROW = 0
End Function
```

### `XRELEVANTAREA`

```vbnet
Public Function XRELEVANTAREA(rngTarget As Range) As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    firstRow = XFIRSTUSEDROW(rngTarget)
    firstCol = XFIRSTUSEDCOL(rngTarget)
    lastRow = XLASTUSEDROW(rngTarget)
    lastCol = XLASTUSEDCOL(rngTarget)
    Set XRELEVANTAREA = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))
End Function
```

### `XCOLUMNWIDTH`

```vbnet
Public Function XCOLUMNWIDTH(target As Range) As Double
    Application.ScreenUpdating = False
    XCOLUMNWIDTH = target.ColumnWidth
    Application.ScreenUpdating = True
End Function
```

### `XGETBOLD`

```vbnet
Public Function XGETBOLD(pWorkRng As Range)
    XGETBOLD = pWorkRng.Font.Bold
End Function
```

### `XGETINDENTLEVEL`

```vbnet
Public Function XGETINDENTLEVEL(targetCell As Range)
    XGETINDENTLEVEL = targetCell.IndentLevel
End Function
```

### `XCOUNTCOLOR`

```vbnet
Public Function XCOUNTCOLOR(CountRange As Range, CountColor As Range)
    Dim CountColorValue As Long
    Dim TotalCount As Long
    Dim rCell As Range
    CountColorValue = CountColor.Interior.ColorIndex
    For Each rCell In CountRange
        If rCell.Interior.ColorIndex = CountColorValue Then
            TotalCount = TotalCount + 1
        End If
    Next rCell
    XCOUNTCOLOR = TotalCount
End Function
```

### `XEXTRACTAFTER`

```vbnet
Public Function XEXTRACTAFTER(rngWord As Range, strWord As String) As String
    On Error GoTo ErrorHandler
    Application.Volatile
    Dim lngStart As Long, lngEnd As Long, tempResult As String
    lngStart = InStr(1, rngWord, strWord)
    If lngStart = 0 Then
        XEXTRACTAFTER = "Not found": Exit Function
    End If
    lngEnd = InStr(lngStart + Len(strWord), rngWord, Len(rngWord))
    If lngEnd = 0 Then lngEnd = Len(rngWord)
    tempResult = Mid(rngWord, lngStart + Len(strWord), lngEnd - lngStart)
    XEXTRACTAFTER = Trim(tempResult)
    Exit Function
ErrorHandler:
    XEXTRACTAFTER = Err.Description
End Function
```

### `XEXTRACTBEFORE`

```vbnet
Public Function XEXTRACTBEFORE(rngWord As Range, strWord As String) As String
    On Error GoTo ErrorHandler
    Application.Volatile
    Dim lngEnd As Long, tempResult As String
    lngEnd = InStr(1, rngWord, strWord)
    If lngEnd = 0 Then
        XEXTRACTBEFORE = "Not found": Exit Function
    End If
    tempResult = Left(rngWord, lngEnd - 1)
    XEXTRACTBEFORE = Trim(tempResult)
    Exit Function
ErrorHandler:
    XEXTRACTBEFORE = Err.Description
End Function
```

### `XHYPERACTIVE`

```vbnet
Public Function XHYPERACTIVE(ByRef Rng As Range)
    Dim strAddress As String, strTextDisplay As String
    Dim target As Range
    Application.DisplayAlerts = False
    On Error Resume Next
    Set target = Application.InputBox(Title:="Create Hyperlink", _
                                      Prompt:="Select a cell to create hyperlink", _
                                      Type:=8)
    On Error GoTo 0
    Application.DisplayAlerts = True
    If Rng Is Nothing Or target Is Nothing Then Exit Function
    strAddress = Chr(39) & target.Parent.Name & Chr(39) & "!" & target.Address
    If WorksheetFunction.CountA(Rng) = 0 Then
        strTextDisplay = target.Parent.Name
    Else
        strTextDisplay = Rng.Value
    End If
    With ActiveSheet.Hyperlinks
        .Add Anchor:=Rng, Address:="", SubAddress:=strAddress, TextToDisplay:=strTextDisplay
    End With
End Function
```

## Module `Ribbon`

### `RibbonOnLoad`

```vbnet
Public Sub RibbonOnLoad(r As IRibbonUI)
    Set RibbonUI = r
End Sub
```

### `RunByName`

```vbnet
Public Sub RunByName(control As IRibbonControl)
    Dim macro As String
    macro = control.Tag
    If Len(macro) = 0 Then macro = control.ID
    On Error GoTo errh
    Application.Run macro
    Exit Sub
errh:
    MsgBox "Macro not found: " & macro, vbExclamation
End Sub
```

### `GetHighlightLabel`

```vbnet
Public Sub GetHighlightLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastHighlight", "Green")
End Sub
```

### `GetWorkbookFontLabel`

```vbnet
Public Sub GetWorkbookFontLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastWbFont", "Arial")
End Sub
```

### `GetWorkbookFontSizeLabel`

```vbnet
Public Sub GetWorkbookFontSizeLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastWbFontSize", "10")
End Sub
```

### `GetSheetFontLabel`

```vbnet
Public Sub GetSheetFontLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastShFont", "Arial")
End Sub
```

### `GetSheetTabLabel`

```vbnet
Public Sub GetSheetTabLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastShTab", "Green")
End Sub
```

### `GetSelNumberLabel`

```vbnet
Public Sub GetSelNumberLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelNumber", "Accounting")
End Sub
```

### `GetSelCaseLabel`

```vbnet
Public Sub GetSelCaseLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelCase", "Proper")
End Sub
```

## Module `Subs`

### `CaseProper`

```vbnet
Sub CaseProper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = StrConv(cell, vbProperCase)
    Next cell
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelCase", "Proper"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `CaseSentence`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastSelCase", "Sentence"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `CaseUpper`

```vbnet
Sub CaseUpper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = UCase(cell)
    Next cell
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelCase", "Upper"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `CellTrim`

```vbnet
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
```

### `ETLSAM`

```vbnet
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
```

### `ETLTB`

```vbnet
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
```

### `FormatAccounting`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Accounting"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormatCellRed`

```vbnet
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
```

### `FormatDateDDMMM`

```vbnet
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
```

### `FormatDateDDMMYY`

```vbnet
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
```

### `FormatFontBlue`

```vbnet
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
```

### `FormatFontGreen`

```vbnet
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
```

### `FormatFontOrange`

```vbnet
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
```

### `FormatHighlightGreen`

```vbnet
Sub FormatHighlightGreen()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(204, 255, 204)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Green"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate

ErrorHandler:
    Exit Sub
End Sub
```

### `FormatHighlightRed`

```vbnet
Sub FormatHighlightRed()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 204, 204)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Red"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate

ErrorHandler:
    Exit Sub
End Sub
```

### `FormatHighlightYellow`

```vbnet
Sub FormatHighlightYellow()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 255, 0)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Yellow"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
End Sub
```

### `FormatHighlightReset`

```vbnet
Sub FormatHighlightReset()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = xlNone
    Rng.Font.Color = RGB(0, 0, 0)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Clear"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
End Sub
```

### `FormatTableBordersGrey`

```vbnet
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
```

### `FormatTextToValue`

```vbnet
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
```

### `FormulaAbsolute`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Absolute"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `FormulaReverseSign`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Reverse Sign"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormulaRound`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Round"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `FormulaThousands`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Thousands"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormulaToValue`

```vbnet
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
```

### `InsertArrowDown`

```vbnet
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
```

### `InsertBlankSheet`

```vbnet
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
```

### `InsertColumnWidth`

```vbnet
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
```

### `InsertCrossReference`

```vbnet
Sub InsertCrossReference()

    Call XHYPERACTIVE(Selection)

End Sub
```

### `InsertTimestamp`

```vbnet
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
```

### `InsertWorkdone`

```vbnet
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
```

### `RemoveBlankCells`

```vbnet
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
```

### `RemoveBlankRows`

```vbnet
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
```

### `SheetColumnsFS`

```vbnet
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
```

### `SheetColumnsNTA4X`

```vbnet
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
```

### `SheetColumnsWP`

```vbnet
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
```

### `SheetFontArial`

```vbnet
Sub SheetFontArial()

    On Error GoTo ErrorHandler
    
    If Not ActiveSheet.ProtectContents Then
        ActiveSheet.Cells.Font.Name = "Arial"
        ActiveSheet.Cells.Font.Size = 8
        ActiveSheet.Activate
        ActiveWindow.Zoom = 100
    Else: Exit Sub
    End If
    
    SaveSetting "ExcelUI", "Preferences", "LastShFont", "Arial"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontEY`

```vbnet
Sub SheetFontEY()

    On Error GoTo ErrorHandler
    
    If Not ActiveSheet.ProtectContents Then
        ActiveSheet.Cells.Font.Name = "EYInterstate Light"
        ActiveSheet.Cells.Font.Size = 8
        ActiveSheet.Activate
        ActiveWindow.Zoom = 100
    Else: Exit Sub
    End If
        
    SaveSetting "ExcelUI", "Preferences", "LastShFont", "EY"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontSize10`

```vbnet
Sub SheetFontSize10()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 10
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontSize8`

```vbnet
Sub SheetFontSize8()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 8
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100

ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFormulaToValue`

```vbnet
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
```

### `SheetPageBreakOff`

```vbnet
Sub SheetPageBreakOff()

    On Error GoTo ErrorHandler
    
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.Activate
    ActiveWindow.DisplayGridlines = False
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetRemoveBlankRows`

```vbnet
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
```

### `SheetResetComments`

```vbnet
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
```

### `SheetTabBlack`

```vbnet
Sub SheetTabBlack()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 1
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Black"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabGreen`

```vbnet
Sub SheetTabGreen()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 35
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Green"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `SheetTabRed`

```vbnet
Sub SheetTabRed()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 38
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Red"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabReset`

```vbnet
Sub SheetTabReset()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = xlColorIndexNone
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Reset"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabYellow`

```vbnet
Sub SheetTabYellow()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 6
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Yellow"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
        
ErrorHandler:
    Exit Sub

End Sub
```

### `UnhideSheets`

```vbnet
Sub UnhideSheets()

    On Error GoTo ErrorHandler
    
    For Each sh In Worksheets: sh.Visible = True: Next sh
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `WorkbookArial`

```vbnet
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
    
    SaveSetting "ExcelUI", "Preferences", "LastWbFont", "Arial"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `WorkbookEY`

```vbnet
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

    SaveSetting "ExcelUI", "Preferences", "LastWbFont", "EY"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `WorkbookPageBreakOff`

```vbnet
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
```

### `WorkbookFontSize8`

```vbnet
Sub WorkbookFontSize8()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 8
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "8"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `WorkbookFontSize9`

```vbnet
Sub WorkbookFontSize9()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 9
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "9"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `WorkbookFontSize10`

```vbnet
Sub WorkbookFontSize10()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 10
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "10"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `WorkbookFontSize11`

```vbnet
Sub WorkbookFontSize11()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 11
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "11"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `RunFormatHighlightRepeat`

```vbnet
Sub RunFormatHighlightRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastHighlight", "Green")
    Select Case last
        Case "Green":  FormatHighlightGreen
        Case "Red":    FormatHighlightRed
        Case "Yellow": FormatHighlightYellow
        Case Else:     FormatHighlightReset
    End Select
End Sub
```

### `RunWorkbookFontRepeat`

```vbnet
Sub RunWorkbookFontRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastWbFont", "Arial")
    Select Case last
        Case "Arial": WorkbookArial
        Case "EY":    WorkbookEY
    End Select
End Sub
```

### `RunWorkbookFontSizeRepeat`

```vbnet
Sub RunWorkbookFontSizeRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastWbFontSize", "10")
    Select Case last
        Case "8":  WorkbookFontSize8
        Case "9":  WorkbookFontSize9
        Case "10": WorkbookFontSize10
        Case "11": WorkbookFontSize11
    End Select
End Sub
```

### `RunSheetFontRepeat`

```vbnet
Sub RunSheetFontRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastShFont", "Arial")
    Select Case last
        Case "Arial": SheetFontArial
        Case "EY":    SheetFontEY
    End Select
End Sub
```

### `RunSheetTabRepeat`

```vbnet
Sub RunSheetTabRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastShTab", "Green")
    Select Case last
        Case "Green":  SheetTabGreen
        Case "Red":    SheetTabRed
        Case "Yellow": SheetTabYellow
        Case Else:     SheetTabReset
    End Select
End Sub
```

### `RunSelNumberRepeat`

```vbnet
Sub RunSelNumberRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastSelNumber", "Accounting")
    Select Case last
        Case "Accounting":   FormatAccounting
        Case "Round":        FormulaRound
        Case "Absolute":     FormulaAbsolute
        Case "Reverse Sign": FormulaReverseSign
    End Select
End Sub
```

### `RunSelCaseRepeat`

```vbnet
Sub RunSelCaseRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastSelCase", "Proper")
    Select Case last
        Case "Proper":   CaseProper
        Case "Upper":    CaseUpper
        Case "Sentence": CaseSentence
    End Select
End Sub
```
