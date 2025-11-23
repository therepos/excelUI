Attribute VB_Name = "Module1"
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