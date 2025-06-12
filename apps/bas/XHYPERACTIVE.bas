Private Function XHYPERACTIVE(ByRef Rng As Range)

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

    If Rng Is Nothing Then
        Exit Function
    Else
        strAddress = Chr(39) & target.Parent.Name & Chr(39) & "!" & target.Address
        If WorksheetFunction.CountA(Rng) = 0 Then
            strTextDisplay = target.Parent.Name
        Else
            strTextDisplay = Rng.Value
        End If
        
        With ActiveSheet.Hyperlinks
        .Add Anchor:=Rng, _
             Address:="", _
             SubAddress:=strAddress, _
             TextToDisplay:=strTextDisplay
        End With
    End If
    
ErrorHandler:
    Exit Function
    
End Function
