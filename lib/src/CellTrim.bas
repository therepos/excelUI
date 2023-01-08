Sub CellTrim()

    Dim rng As Range
    Set rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(rng)
        cell.Value = Trim(cell)
    Next cell
    Application.ScreenUpdating = True

End Sub

