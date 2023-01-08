Sub RibbonOnLoad(Rib As IRibbonUI)
'   Purpose: Initialise the ribbon
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Modification:
'   - Initialise tab, group and dropdown items to illustrate all 3 features in one example
'   Updated: 2022FEB28

    Set Ribbon = Rib
    MySelectedTabTag = "tb1"
    MySelectedGroupTag = "tb1gp2"
    MySelectedItemID = "tb2gp2dd1_01"

End Sub

Sub RibbonRefresh(TabTag As String, Optional TabID As String, Optional GroupTag As String)
'   Purpose: Refresh the ribbon
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Modification:
'   - Included change of group
'   - Included ScreenUpdating to reduce 'loading' scene
'   Updated: 2022FEB28

    Application.ScreenUpdating = False
    MySelectedTabTag = TabTag
    
    If GroupTag <> "" Then
        MySelectedGroupTag = GroupTag
    End If
    
    If Ribbon Is Nothing Then
        MsgBox "Error, Save/Restart your workbook"
    Else
        Ribbon.Invalidate
        If TabID <> "" Then Ribbon.ActivateTab TabID
    End If
    Application.ScreenUpdating = True
    
End Sub

Sub ShowTab(control As IRibbonControl, ByRef visible)
'   Purpose: Tells which tab to be visible based on the MySelectedTabTag
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Updated: 2022FEB28

    If control.Tag Like MySelectedTabTag Then
        visible = True
    Else
        visible = False
    End If

End Sub

Sub ShowGroup(control As IRibbonControl, ByRef visible)
'   Purpose: Tells which group to be visible based on the MySelectedGroupTabTag
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Updated: 2022FEB28

    If control.Tag Like MySelectedGroupTag Then
        visible = True
    Else
        visible = False
    End If

End Sub

Sub ChangeTab(control As IRibbonControl)
'   Purpose: Display ribbon tab on demand
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Modification:
'   - Included Select statement to cater for different trigger
'   =====================================
'   Notes:
'       Tag:="testTab"     Show/Hide only the Tab, Group or Control with Tag "testTab"
'       Tag:="My*"         Show/Hide every Tab, Group or Control with Tag that starts with "My"
'       Tag:="*"           Show/Hide every Tab, Group or Control
'       Tag:=""            Hide every Tab, Group or Control
'   ======================================
'   Updated: 2022FEB28
    
    Select Case MySelectedTabTag
        Case "tb1": Call RibbonRefresh(TabTag:="tb2", TabID:="tb2")
        Case "tb2": Call RibbonRefresh(TabTag:="tb1", TabID:="tb1")
    End Select

End Sub

Sub ChangeGroup(control As IRibbonControl)
'   Purpose: Display tab group on demand
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Modification:
'   - Included Select statement to cater for different trigger
'   Updated: 2022FEB28

    Select Case MySelectedGroupTag
        Case "tb1gp2": Call RibbonRefresh(TabTag:=MySelectedTabTag, GroupTag:="tb1gp3")
        Case "tb1gp3": Call RibbonRefresh(TabTag:=MySelectedTabTag, GroupTag:="tb1gp2")
    End Select
    
End Sub

Sub GetDefaultItemID(ByRef control As IRibbonControl, ByRef returnedVal As Variant)
'   Purpose: Get default item to display by ID
'   Updated: 2022FEB28

    returnedVal = MySelectedItemID
    
End Sub

Sub GetSelectedItemID(control As IRibbonControl, ID As String, index As Integer)
'   Purpose: Get user selected item by ID
'   Updated: 2022FEB28

    MySelectedItemID = ID

End Sub

Sub LabelNextGroup(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate button label based on the opposite state to MySelectedGroupTag
'   Updated: 2022FEB28

    Select Case MySelectedGroupTag
        Case "tb1gp2": returnedVal = "Group 3"
        Case "tb1gp3": returnedVal = "Group 2"
    End Select

End Sub
