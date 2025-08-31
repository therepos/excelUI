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

