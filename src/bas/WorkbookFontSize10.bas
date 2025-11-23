Attribute VB_Name = "Module1"
Sub WorkbookFontSize10()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 10
    Next ws

ErrorHandler:
    Exit Sub

End Sub