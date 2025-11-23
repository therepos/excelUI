Attribute VB_Name = "Functions"
Option Private Module

'===== Column/Row utilities =====
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

Public Function XRELEVANTAREA(rngTarget As Range) As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    firstRow = XFIRSTUSEDROW(rngTarget)
    firstCol = XFIRSTUSEDCOL(rngTarget)
    lastRow = XLASTUSEDROW(rngTarget)
    lastCol = XLASTUSEDCOL(rngTarget)
    Set XRELEVANTAREA = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))
End Function

'===== Formatting / inspection helpers =====
Public Function XCOLUMNWIDTH(target As Range) As Double
    Application.ScreenUpdating = False
    XCOLUMNWIDTH = target.ColumnWidth
    Application.ScreenUpdating = True
End Function

Public Function XGETBOLD(pWorkRng As Range)
    XGETBOLD = pWorkRng.Font.Bold
End Function

Public Function XGETINDENTLEVEL(targetCell As Range)
    XGETINDENTLEVEL = targetCell.IndentLevel
End Function

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

'===== Text extractors =====
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

'===== Hyperlink helper (called by InsertCrossReference) =====
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


