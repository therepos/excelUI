# CustomUI
CustomUI implements embedded Excel ribbon with advanced customisation. It also overrides any existing ribbon entirely.

- [Install] Excel addins .xlam file [CustomUI Example](https://github.com/therepos/msexcel/blob/main/apps/xlam/customui-example.xlam). 
- Read/Write embedded XML file with [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).

![Features](/img/img-commonaddin-tabmain.png)

## Documentation

The following examples are based on [CustomUI Sample](https://github.com/therepos/msexcel/blob/main/apps/xlam/customui-sample.xlam) to illustrate the concept.  
See official Microsoft [documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview) for more details.

### Implementation
```plain title="Minimum setup for a customUI add-in"
MsExcel.xlam
├── VBA Modules
│   ├── Controls    # vba functions of each button
│   └── Ribbon      # vba controls the ribbon
└── CustomUI        # xml displays the ribbon
```

```vba title="VBA Controls Module"
Sub WorkbookArial(control As IRibbonControl)
  
    On Error GoTo ErrorHandler
    
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
    
ErrorHandler:
    Exit Sub
    
End Sub
```

```vba title="VBA Ribbon Module"
Public Ribbon As IRibbonUI
Public MySelectedTabTag As String
Public MySelectedGroupTag As String

Sub RibbonOnLoad(Rib As IRibbonUI)
    Set Ribbon = Rib
    MySelectedTabTag = "t1"
    MySelectedGroupTag = "t1g1"
    
End Sub
```

```xml title="CustomUI"
<!-- XML -->
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonOnLoad">
    <ribbon>
        <tabs>
            <!-- Tab --> 
            <tab id="t1" tag="t1" label="TabLabel">
                <!-- Group -->
                <group id="t1g1" tag="t1g1" label="Workbook" autoScale="false">   
                    <!-- Box is a layout container in a group --> 
                    <box id="t1g1x1" boxStyle="horizontal">                                           
                        <!-- Button --> 
                        <button id="t1g1b1" label="Arial 8" size="large" onAction="WorkbookArial" imageMso="CharacterBorder" />   
                    </box>                                                                
                </group>                                                
            </tab>                                   
        </tabs>
    </ribbon>
</customUI>
```

### OnLoad

```xml title="Initialise state of controls on load"
<!-- XML -->
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" 
            onLoad="RibbonOnLoad">
```
```vb title="Declare global variables and object"
Public Ribbon As IRibbonUI
Public MySelectedTabTag As String
Public MySelectedGroupTag As String
Public MySelectedItemID As String
```
```vb title="Initialise the ribbon"
Sub RibbonOnLoad(Rib As IRibbonUI)

    Set Ribbon = Rib
    MySelectedTabTag = "tb1"
    MySelectedGroupTag = "tb1gp2"
    MySelectedItemID = "tb2gp2dd1_01"
    
End Sub
```

### GetVisible

```xml title="Generate visible attribute of tab control, dynamically"
<!-- XML -->
<tab id="tb1" 
    tag="tb1" 
    label="Tab 1" 
    getVisible="ShowTab">
```
```vb title="Set visible attribute of tab control based on the MySelectedTabTag"
Sub ShowTab(control As IRibbonControl, ByRef visible)

    If control.Tag Like MySelectedTabTag Then
        visible = True
    Else
        visible = False
    End If

End Sub
```
```xml title="Generate visible attribute of group control, dynamically"
<!-- XML -->
<group id="tb1gp3" 
    tag="tb1gp3" 
    label="Group 3" 
    getVisible="ShowGroup">
```
```vb title="Set visible attribute of group control based on the MySelectedGroupTag"
Sub ShowGroup(control As IRibbonControl, ByRef visible)

    If control.Tag Like MySelectedGroupTag Then
        visible = True
    Else
        visible = False
    End If

End Sub
```

### OnAction

```xml title="Execute change of tab"
<!-- XML -->
<button id="tb1gp1mn1_01" 
        label="Show Tab 2" 
        onAction="ChangeTab" 
        imageMso="ControlTabControl" />
```
```vb title="Display ribbon tab on demand"
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

```xml title="Execute change of group"
<!-- XML -->
<button id="tb1gp1mn1_02" 
        getLabel="LabelNextGroup"
        onAction="ChangeGroup"
        imageMso="FormControlGroupBox" />
    ```
```vb title="Display tab group on demand"
Sub ChangeGroup(control As IRibbonControl)

    Select Case MySelectedGroupTag
        Case "tb1gp2": Call RibbonRefresh(TabTag:=MySelectedTabTag, GroupTag:="tb1gp3")
        Case "tb1gp3": Call RibbonRefresh(TabTag:=MySelectedTabTag, GroupTag:="tb1gp2")
    End Select
    
End Sub
```

```vb title="Refresh the ribbon"
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

```xml title="Generate the default selected item of a dropdown control, dynamically"
<dropDown id="tb2gp2dd1" 
        label="Dropdown 1" 
        sizeString="WWWWWWWWWWW" 
        getSelectedItemID="GetDefaultItemID" 
        onAction="GetSelectedItemID">
```
```vb title="Get default item to display by ID"
Sub GetDefaultItemID(ByRef control As IRibbonControl, ByRef returnedVal As Variant)

    returnedVal = MySelectedItemID
    
End Sub
```

### GetLabel

```xml title="Generate label attribute of control, dynamically"
<!-- XML -->
<button id="tb1gp1mn1_02" 
        getLabel="LabelNextGroup"
        onAction="ChangeGroup"
        imageMso="FormControlGroupBox" />
```
```vb title="Generate button label based on the opposite state to MySelectedGroupTag"
Sub LabelNextGroup(control As IRibbonControl, ByRef returnedVal)

    Select Case MySelectedGroupTag
        Case "tb1gp2": returnedVal = "Group 3"
        Case "tb1gp3": returnedVal = "Group 2"
    End Select

End Sub
```

```vb title="Get user selected item by ID"
Sub GetSelectedItemID(control As IRibbonControl, ID As String, index As Integer)

    MySelectedItemID = ID

End Sub
```  

## Gallery
<img src="/msexcel/img/CharacterBorder.png" alt="CharacterBorder" title="CharacterBorder" />
<img src="/msexcel/img/CharacterShading.png" alt="CharacterShading" title="CharacterShading" />
<img src="/msexcel/img/CombineShapesGallery.png" alt="CombineShapesGallery" title="CombineShapesGallery" />
<img src="/msexcel/img/CreateMacro.png" alt="CreateMacro" title="CreateMacro" />
<img src="/msexcel/img/BlogManageAccounts.png" alt="BlogManageAccounts" title="BlogManageAccounts" />
<img src="/msexcel/img/Clear.png" alt="Clear" title="Clear" />
<img src="/msexcel/img/AutoFilterClassic.png" alt="AutoFilterClassic" title="AutoFilterClassic" />
<img src="/msexcel/img/AudioNoteDelete.png" alt="AudioNoteDelete" title="AudioNoteDelete" />
<img src="/msexcel/img/ApproveApprovalRequest.png" alt="ApproveApprovalRequest" title="ApproveApprovalRequest" />
<img src="/msexcel/img/ChangeToDeclineInvitation.png" alt="ChangeToDeclineInvitation" title="ChangeToDeclineInvitation" />
<img src="/msexcel/img/AutoFormat.png" alt="AutoFormat" title="AutoFormat" />
<img src="/msexcel/img/CustomEquationsGallery.png" alt="CustomEquationsGallery" title="CharacterShading" />
<img src="/msexcel/img/AlignLeft.png" alt="AlignLeft" title="AlignLeft" />
<img src="/msexcel/img/AppointmentColor0.png" alt="AppointmentColor3" title="AppointmentColor3" />
<img src="/msexcel/img/AppointmentColor1.png" alt="AppointmentColor3" title="AppointmentColor3" />
<img src="/msexcel/img/AppointmentColor2.png" alt="AppointmentColor3" title="AppointmentColor3" />
<img src="/msexcel/img/AppointmentColor3.png" alt="AppointmentColor3" title="AppointmentColor3" />
<img src="/msexcel/img/AppointmentColor4.png" alt="AppointmentColor4" title="AppointmentColor4" />
<img src="/msexcel/img/AppointmentColor5.png" alt="AppointmentColor5" title="AppointmentColor5" />
<img src="/msexcel/img/AppointmentColor6.png" alt="AppointmentColor6" title="AppointmentColor6" />
<img src="/msexcel/img/AppointmentColor7.png" alt="AppointmentColor7" title="AppointmentColor7" />
<img src="/msexcel/img/AppointmentColor8.png" alt="AppointmentColor8" title="AppointmentColor8" />
<img src="/msexcel/img/AppointmentColor9.png" alt="AppointmentColor9" title="AppointmentColor9" />
<img src="/msexcel/img/AppointmentColor10.png" alt="AppointmentColor9" title="AppointmentColor10" />
<img src="/msexcel/img/img-mso-BlackAndWhiteBlack.png" alt="BlackAndWhiteBlack" title="BlackAndWhiteBlack" />
<img src="/msexcel/img/img-mso-BlackAndWhiteAutomatic.png" alt="BlackAndWhiteAutomatic" title="BlackAndWhiteAutomatic" />
<img src="/msexcel/img/BlackAndWhite.png" alt="BlackAndWhite" title="BlackAndWhite" />
<img src="/msexcel/img/img-mso-BlackAndWhiteGrayWithWhiteFill.png" alt="BlackAndWhiteGrayWithWhiteFill" title="BlackAndWhiteGrayWithWhiteFill" />
<img src="/msexcel/img/img-mso-BlackAndWhiteInverseGrayscale.png" alt="BlackAndWhiteInverseGrayscale" title="BlackAndWhiteInverseGrayscale" />
<img src="/msexcel/img/img-mso-ArrangeByLogContact.png" alt="ArrangeByLogContact" title="ArrangeByLogContact" />
<img src="/msexcel/img/img-mso-AnimationStartDropdown.png" alt="AnimationStartDropdown" title="AnimationStartDropdown" />
<img src="/msexcel/img/img-mso-BlogHomePage.png" alt="BlogHomePage" title="BlogHomePage" />
<img src="/msexcel/img/img-mso-BlogCategoriesRefresh.png" alt="BlogCategoriesRefresh" title="BlogCategoriesRefresh" />
<img src="/msexcel/img/Chart3DColumnChart.png" alt="Chart3DColumnChart" title="Chart3DColumnChart" />
---

<img src="/msexcel/img/img-mso-ContactCardCallOther.png" alt="ContactCardCallOther" title="ContactCardCallOther" />
<img src="/msexcel/img/img-mso-ChartInsert.png" alt="ChartInsert" title="ChartInsert" />
<img src="/msexcel/img/img-mso-ChartAreaChart.png" alt="ChartAreaChart" title="ChartAreaChart" />
<img src="/msexcel/img/img-mso-Chart3DConeChart.png" alt="Chart3DConeChart" title="Chart3DConeChart" />
<img src="/msexcel/img/img-mso-Chart3DBarChart.png" alt="Chart3DBarChart" title="Chart3DBarChart" />
---

## Resources
- [ImageMSO Gallery](https://bert-toolkit.com/imagemso-list.html)

<!-- Links -->

[Install]: https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460#:~:text=COM%20add%2Din-,Click%20the%20File%20tab%2C%20click%20Options%2C%20and%20then%20click%20the,install%2C%20and%20then%20click%20OK.