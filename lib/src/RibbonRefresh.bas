Sub RibbonRefresh(Tag As String, Optional TabID As String)
'   Purpose: Refresh the ribbon and activate the custom tab
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Updated: 2022FEB25

    Application.ScreenUpdating = False
    MyTag = Tag
    If Ribbon Is Nothing Then
        MsgBox "Error, Save/Restart your workbook"
    Else
        Ribbon.Invalidate
        If TabID <> "" Then Ribbon.ActivateTab TabID
    End If
    Application.ScreenUpdating = True
    
End Sub

