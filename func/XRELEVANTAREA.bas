Private Function XRELEVANTAREA(rngTarget As Range) As Range

    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    
    firstRow = XFIRSTUSEDROW(rngTarget)
    firstCol = XFIRSTUSEDCOL(rngTarget)
    lastRow = XLASTUSEDROW(rngTarget)
    lastCol = XLASTUSEDCOL(rngTarget)
    Set XRELEVANTAREA = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))
    
End Function
