# CustomUI Ribbon Template

CustomUI Ribbon Template provides a simple starting point to implement customised Excel ribbon. 

- Download [CustomUI Ribbon Template v1.0.0](../dist/customuiribbon-t100.xlam). 
- Read/Write embedded XML file with [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).
- Microsoft Excel add-ins [documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview).

## Documentation

### OnLoad

- Initialise state of controls on load.

    ```xml
    <!-- XML -->
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" 
              onLoad="RibbonOnLoad">
    ```
    ```vb
    ' VBA
    ' Declare global variables and object
    Public Ribbon As IRibbonUI
    Public MySelectedTabTag As String
    Public MySelectedGroupTag As String
    Public MySelectedItemID As String
    ```
    ```vb
    ' VBA
    ' Initialise the ribbon
    Sub RibbonOnLoad(Rib As IRibbonUI)

        Set Ribbon = Rib
        MySelectedTabTag = "tb1"
        MySelectedGroupTag = "tb1gp2"
        MySelectedItemID = "tb2gp2dd1_01"
        
    End Sub
    ```
### GetVisible

- Generate visible attribute of tab control, dynamically.

    ```xml
    <!-- XML -->
    <tab id="tb1" 
        tag="tb1" 
        label="Tab 1" 
        getVisible="ShowTab">
    ```
    ```vb
    ' VBA
    ' Set visible attribute of tab control based on the MySelectedTabTag
    Sub ShowTab(control As IRibbonControl, ByRef visible)

        If control.Tag Like MySelectedTabTag Then
            visible = True
        Else
            visible = False
        End If

    End Sub
    ```

- Generate visible attribute of group control, dynamically.

    ```xml
    <!-- XML -->
    <group id="tb1gp3" 
        tag="tb1gp3" 
        label="Group 3" 
        getVisible="ShowGroup">
    ```
    ```vb
    ' VBA
    ' Set visible attribute of group control based on the MySelectedGroupTag
    Sub ShowGroup(control As IRibbonControl, ByRef visible)

        If control.Tag Like MySelectedGroupTag Then
            visible = True
        Else
            visible = False
        End If

    End Sub
    ```

### OnAction

- Execute change of tab.

    ```xml
    <!-- XML -->
    <button id="tb1gp1mn1_01" 
            label="Show Tab 2" 
            onAction="ChangeTab" 
            imageMso="ControlTabControl" />
    ```
    ```vb
    ' VBA
    ' Display ribbon tab on demand
    ' =====================================
    ' Tag:="testTab"     Show/Hide only the Tab, Group or Control with Tag "testTab"
    ' Tag:="My*"         Show/Hide every Tab, Group or Control with Tag that starts with "My"
    ' Tag:="*"           Show/Hide every Tab, Group or Control
    ' Tag:=""            Hide every Tab, Group or Control
    ' ======================================
    Sub ChangeTab(control As IRibbonControl)
        
        Select Case MySelectedTabTag
            Case "tb1": Call RibbonRefresh(TabTag:="tb2", TabID:="tb2")
            Case "tb2": Call RibbonRefresh(TabTag:="tb1", TabID:="tb1")
        End Select

    End Sub
    ```

- Execute change of group.

    ```xml
    <!-- XML -->
    <button id="tb1gp1mn1_02" 
            getLabel="LabelNextGroup"
            onAction="ChangeGroup"
            imageMso="FormControlGroupBox" />
    ```
    ```vb
    ' VBA
    ' Display tab group on demand
    Sub ChangeGroup(control As IRibbonControl)

        Select Case MySelectedGroupTag
            Case "tb1gp2": Call RibbonRefresh(TabTag:=MySelectedTabTag, GroupTag:="tb1gp3")
            Case "tb1gp3": Call RibbonRefresh(TabTag:=MySelectedTabTag, GroupTag:="tb1gp2")
        End Select
        
    End Sub
    ```

- Support function to reload ribbon after initial load.

    ```vb
    ' VBA
    ' Refresh the ribbon
    Sub RibbonRefresh(TabTag As String, Optional TabID As String, Optional GroupTag As String)

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
    ```

### GetSelectedItemID

- Generate the default selected item of a dropdown control, dynamically.

    ```xml
    <!-- XML -->
    <dropDown id="tb2gp2dd1" 
            label="Dropdown 1" 
            sizeString="WWWWWWWWWWW" 
            getSelectedItemID="GetDefaultItemID" 
            onAction="GetSelectedItemID">
    ```
    ```vb
    ' VBA
    ' Get default item to display by ID
    Sub GetDefaultItemID(ByRef control As IRibbonControl, ByRef returnedVal As Variant)

        returnedVal = MySelectedItemID
        
    End Sub
    ```

### GetLabel

- Generate label attribute of control, dynamically.

    ```xml
    <!-- XML -->
    <button id="tb1gp1mn1_02" 
            getLabel="LabelNextGroup"
            onAction="ChangeGroup"
            imageMso="FormControlGroupBox" />
    ```
    ```vb
    ' VBA
    ' Generate button label based on the opposite state to MySelectedGroupTag
    Sub LabelNextGroup(control As IRibbonControl, ByRef returnedVal)

        Select Case MySelectedGroupTag
            Case "tb1gp2": returnedVal = "Group 3"
            Case "tb1gp3": returnedVal = "Group 2"
        End Select

    End Sub
    ```

- Get user's choice from dropdown control.

    ```vb
    ' VBA
    ' Get user selected item by ID
    Sub GetSelectedItemID(control As IRibbonControl, ID As String, index As Integer)

        MySelectedItemID = ID

    End Sub
    ```  
    
<!-- ### Examples
<img src="../src/img/img-customuiribbon-t100.gif" width="60%" height="60%"> -->



---
[Excel](https://github.com/ry4nyeo/lab-excel) > 
[CustomUI Ribbon Template](../docs/pgs-customuiribbon.md)

