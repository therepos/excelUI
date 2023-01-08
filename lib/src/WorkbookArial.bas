Sub WorkbookArial()
  
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
    
End Sub
