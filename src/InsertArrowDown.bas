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

