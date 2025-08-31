# CustomUI
CustomUI implements embedded Excel ribbon with advanced customisation. It also overrides any existing ribbon entirely.

- [Install] Excel addins .xlam file [CustomUI Example](https://github.com/therepos/msexcel/releases/latest/download/excelUI.zip). 
- Read/Write embedded XML file with [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).

![Features](/img/img-commonaddin-tabmain.png)

## Documentation

The following examples are based on [CustomUI Minimal](https://github.com/therepos/msexcel/blob/main/apps/xlam/customui-minimal.xlam) to illustrate the concept.  
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
<img src="/msexcel/img/img-mso-All.png" />
---
<img src="/msexcel/img/img-mso-ListMacros.png" alt="ListMacros" title="ListMacros" />
<img src="/msexcel/img/img-mso-CancelRequest.png" alt="CancelRequest" title="CancelRequest" />
<img src="/msexcel/img/img-mso-WorkflowComplete.png" alt="WorkflowComplete" title="WorkflowComplete" />
<img src="/msexcel/img/img-mso-Info.png" alt="Info" title="Info" />
<img src="/msexcel/img/img-mso-DiagramTargetInsertClassic.png" alt="DiagramTargetInsertClassic" title="DiagramTargetInsertClassic" />
<img src="/msexcel/img/img-mso-HighImportance.png" alt="HighImportance" title="HighImportance" />
<img src="/msexcel/img/img-mso-TrustCenter.png" alt="TrustCenter" title="TrustCenter" />
<img src="/msexcel/img/img-mso-AdpStoredProcedureQueryMakeTable.png" alt="AdpStoredProcedureQueryMakeTable" title="AdpStoredProcedureQueryMakeTable" />
<img src="/msexcel/img/img-mso-FormRegionSave.png" alt="FormRegionSave" title="FormRegionSave" />
---
<img src="/msexcel/img/img-mso-OutlinePromoteToHeading.png" alt="OutlinePromoteToHeading" title="OutlinePromoteToHeading" />
<img src="/msexcel/img/img-mso-OutlineDemoteToBodyText.png" alt="OutlineDemoteToBodyText" title="OutlineDemoteToBodyText" />
<img src="/msexcel/img/img-mso-MessagePrevious.png" alt="MessagePrevious" title="MessagePrevious" />
<img src="/msexcel/img/img-mso-MessageNext.png" alt="MessageNext" title="MessageNext" />
<img src="/msexcel/img/img-mso-WebGoForward.png" alt="WebGoForward" title="WebGoForward" />
<img src="/msexcel/img/img-mso-AddressBook.png" alt="AddressBook" title="AddressBook" />
<img src="/msexcel/img/img-mso-FontColorMoreColorsDialog.png" alt="FontColorMoreColorsDialog" title="FontColorMoreColorsDialog" />
<img src="/msexcel/img/img-mso-ObjectEditPoints.png" alt="ObjectEditPoints" title="ObjectEditPoints" />
<img src="/msexcel/img/img-mso-Recurrence.png" alt="Recurrence" title="Recurrence" />
<img src="/msexcel/img/img-mso-SetPertWeights.png" alt="SetPertWeights" title="SetPertWeights" />
<img src="/msexcel/img/img-mso-RmsInvokeBrowser.png" alt="RmsInvokeBrowser" title="RmsInvokeBrowser" />
<img src="/msexcel/img/img-mso-VisibilityHidden.png" alt="VisibilityHidden" title="VisibilityHidden" />
<img src="/msexcel/img/img-mso-ShowClipboard.png" alt="ShowClipboard" title="ShowClipboard" />
<img src="/msexcel/img/img-mso-QueryUnionQuery.png" alt="QueryUnionQuery" title="QueryUnionQuery" />
<img src="/msexcel/img/img-mso-AutoFilterClassic.png" alt="AutoFilterClassic" title="AutoFilterClassic" />
<img src="/msexcel/img/img-mso-ChangesDiscardAndRefresh.png" alt="ChangesDiscardAndRefresh" title="ChangesDiscardAndRefresh" />
<img src="/msexcel/img/img-mso-AutoFormat.png" alt="AutoFormat" title="AutoFormat" />
<img src="/msexcel/img/img-mso-CharacterBorder.png" alt="CharacterBorder" title="CharacterBorder" />
<img src="/msexcel/img/img-mso-CharacterShading.png" alt="CharacterShading" title="CharacterShading" />
<img src="/msexcel/img/img-mso-DollarSign.png" alt="DollarSign" title="DollarSign" />
<img src="/msexcel/img/img-mso-BlogCategoriesRefresh.png" alt="BlogCategoriesRefresh" title="BlogCategoriesRefresh" />
<img src="/msexcel/img/img-mso-BlogHomePage.png" alt="BlogHomePage" title="BlogHomePage" />
<img src="/msexcel/img/img-mso-HyperlinksVerify.png" alt="HyperlinksVerify" title="HyperlinksVerify" />
---
<img src="/msexcel/img/img-mso-BordersAll.png" alt="BordersAll" title="BordersAll" />
<img src="/msexcel/img/img-mso-AppointmentColor0.png" alt="AppointmentColor0" title="AppointmentColor0" />
<img src="/msexcel/img/img-mso-AppointmentColor1.png" alt="AppointmentColor1" title="AppointmentColor1" />
<img src="/msexcel/img/img-mso-AppointmentColor3.png" alt="AppointmentColor3" title="AppointmentColor3" />
<img src="/msexcel/img/img-mso-AppointmentColor4.png" alt="AppointmentColor4" title="AppointmentColor4" />
<img src="/msexcel/img/img-mso-AppointmentColor5.png" alt="AppointmentColor5" title="AppointmentColor5" />
<img src="/msexcel/img/img-mso-AppointmentColor6.png" alt="AppointmentColor6" title="AppointmentColor6" />
<img src="/msexcel/img/img-mso-AppointmentColor7.png" alt="AppointmentColor7" title="AppointmentColor7" />
<img src="/msexcel/img/img-mso-AppointmentColor8.png" alt="AppointmentColor8" title="AppointmentColor8" />
<img src="/msexcel/img/img-mso-AppointmentColor9.png" alt="AppointmentColor9" title="AppointmentColor9" />
<img src="/msexcel/img/img-mso-AppointmentColor10.png" alt="AppointmentColor10" title="AppointmentColor10" />
<img src="/msexcel/img/img-mso-AppointmentColorDialog.png" alt="AppointmentColorDialog" title="AppointmentColorDialog" />
<img src="/msexcel/img/img-mso-AppointmentBusy.png" alt="AppointmentBusy" title="AppointmentBusy" />
<img src="/msexcel/img/img-mso-AppointmentOutOfOffice.png" alt="AppointmentOutOfOffice" title="AppointmentOutOfOffice" />
<img src="/msexcel/img/img-mso-BlackAndWhite.png" alt="BlackAndWhite" title="BlackAndWhite" />
<img src="/msexcel/img/img-mso-BlackAndWhiteAutomatic.png" alt="BlackAndWhiteAutomatic" title="BlackAndWhiteAutomatic" />
<img src="/msexcel/img/img-mso-BlackAndWhiteBlack.png" alt="BlackAndWhiteBlack" title="BlackAndWhiteBlack" />
<img src="/msexcel/img/img-mso-BlackAndWhiteGrayWithWhiteFill.png" alt="BlackAndWhiteGrayWithWhiteFill" title="BlackAndWhiteGrayWithWhiteFill" />
<img src="/msexcel/img/img-mso-BlackAndWhiteInverseGrayscale.png" alt="BlackAndWhiteInverseGrayscale" title="BlackAndWhiteInverseGrayscale" />
<img src="/msexcel/img/img-mso-DataGraphicIconSet.png" alt="DataGraphicIconSet" title="DataGraphicIconSet" />
<img src="/msexcel/img/img-mso-ShapesDuplicate.png" alt="ShapesDuplicate" title="ShapesDuplicate" />
<img src="/msexcel/img/img-mso-CreateMacro.png" alt="CreateMacro" title="CreateMacro" />
<img src="/msexcel/img/img-mso-CustomEquationsGallery.png" alt="CustomEquationsGallery" title="CustomEquationsGallery" />
<img src="/msexcel/img/img-mso-EquationMatrixGallery.png" alt="EquationMatrixGallery" title="EquationMatrixGallery" />
<img src="/msexcel/img/img-mso-DiagramRadialInsertClassic.png" alt="DiagramRadialInsertClassic" title="DiagramRadialInsertClassic" />
<img src="/msexcel/img/img-mso-SmartArtInsert.png" alt="SmartArtInsert" title="SmartArtInsert" />
<img src="/msexcel/img/img-mso-ChartInsert.png" alt="ChartInsert" title="ChartInsert" />
<img src="/msexcel/img/img-mso-Chart3DBarChart.png" alt="Chart3DBarChart" title="Chart3DBarChart" />
<img src="/msexcel/img/img-mso-Chart3DConeChart.png" alt="Chart3DConeChart" title="Chart3DConeChart" />
<img src="/msexcel/img/img-mso-Chart3DPieChart.png" alt="Chart3DPieChart" title="Chart3DPieChart" />
<img src="/msexcel/img/img-mso-ChartAreaChart.png" alt="ChartAreaChart" title="ChartAreaChart" />
<img src="/msexcel/img/img-mso-ChartRadarChart.png" alt="ChartRadarChart" title="ChartRadarChart" />
<img src="/msexcel/img/img-mso-WordArtFormatDialog.png" alt="WordArtFormatDialog" title="WordArtFormatDialog" />
<img src="/msexcel/img/img-mso-BuildingBlocksOrganizer.png" alt="BuildingBlocksOrganizer" title="BuildingBlocksOrganizer" />
<img src="/msexcel/img/img-mso-WindowsCascade.png" alt="WindowsCascade" title="WindowsCascade" />
---
<img src="/msexcel/img/img-mso-ContactPictureMenu.png" alt="ContactPictureMenu" title="ContactPictureMenu" />
<img src="/msexcel/img/img-mso-ContactCardCallOther.png" alt="ContactCardCallOther" title="ContactCardCallOther" />
<img src="/msexcel/img/img-mso-ShapeSmileyFace.png" alt="ShapeSmileyFace" title="ShapeSmileyFace" />
<img src="/msexcel/img/img-mso-ShowContactPage.png" alt="ShowContactPage" title="ShowContactPage" />
<img src="/msexcel/img/img-mso-ReminderSound.png" alt="ReminderSound" title="ReminderSound" />
<img src="/msexcel/img/img-mso-SpeechMicrophone.png" alt="SpeechMicrophone" title="SpeechMicrophone" />
<img src="/msexcel/img/img-mso-LassoSelect.png" alt="LassoSelect" title="LassoSelect" />
<img src="/msexcel/img/img-mso-InkEraseMode.png" alt="InkEraseMode" title="InkEraseMode" />
<img src="/msexcel/img/img-mso-SignaturesLoading.png" alt="SignaturesLoading" title="SignaturesLoading" />
<img src="/msexcel/img/img-mso-SignatureShow.png" alt="SignatureShow" title="SignatureShow" />
<img src="/msexcel/img/img-mso-ZoomPrintPreviewExcel.png" alt="ZoomPrintPreviewExcel" title="ZoomPrintPreviewExcel" />
<img src="/msexcel/img/img-mso-AudioNoteDelete.png" alt="AudioNoteDelete" title="AudioNoteDelete" />
<img src="/msexcel/img/img-mso-CondolatoryEvent.png" alt="CondolatoryEvent" title="CondolatoryEvent" />
<img src="/msexcel/img/img-mso-ShapeSeal8.png" alt="ShapeSeal8" title="ShapeSeal8" />
<img src="/msexcel/img/img-mso-ToolboxVideo.png" alt="ToolboxVideo" title="ToolboxVideo" />
<img src="/msexcel/img/img-mso-NewNote.png" alt="NewNote" title="NewNote" />
<img src="/msexcel/img/img-mso-EditBusinessCard.png" alt="EditBusinessCard" title="EditBusinessCard" />
<img src="/msexcel/img/img-mso-GoToMail.png" alt="GoToMail" title="GoToMail" />
<img src="/msexcel/img/img-mso-AnimationOnClick.png" alt="AnimationOnClick" title="AnimationOnClick" />
<img src="/msexcel/img/img-mso-AnimationStartDropdown.png" alt="AnimationStartDropdown" title="AnimationStartDropdown" />
---
<img src="/msexcel/img/img-mso-Pushpin.png" alt="Pushpin" title="Pushpin" />
<img src="/msexcel/img/img-mso-Piggy.png" alt="Piggy" title="Piggy" />
<img src="/msexcel/img/img-mso-PanAndZoomWindow.png" alt="PanAndZoomWindow" title="PanAndZoomWindow" />
<img src="/msexcel/img/img-mso-PanningHand.png" alt="PanningHand" title="PanningHand" />
<img src="/msexcel/img/img-mso-PositionAbsoluteMarks.png" alt="PositionAbsoluteMarks" title="PositionAbsoluteMarks" />
<img src="/msexcel/img/img-mso-GridShowHide.png" alt="GridShowHide" title="GridShowHide" />
<img src="/msexcel/img/img-mso-ViewSheetGridlines.png" alt="ViewSheetGridlines" title="ViewSheetGridlines" />
<img src="/msexcel/img/img-mso-HorizontalSpacingDecrease.png" alt="HorizontalSpacingDecrease" title="HorizontalSpacingDecrease" />
<img src="/msexcel/img/img-mso-HorizontalSpacingIncrease.png" alt="HorizontalSpacingIncrease" title="HorizontalSpacingIncrease" />
<img src="/msexcel/img/img-mso-ObjectsGroup.png" alt="ObjectsGroup" title="ObjectsGroup" />
<img src="/msexcel/img/img-mso-ObjectsUngroup.png" alt="ObjectsUngroup" title="ObjectsUngroup" />
<img src="/msexcel/img/img-mso-CreateFormBlankForm.png" alt="CreateFormBlankForm" title="CreateFormBlankForm" />
<img src="/msexcel/img/img-mso-DeleteRows.png" alt="DeleteRows" title="DeleteRows" />
<img src="/msexcel/img/img-mso-SelectRecord.png" alt="SelectRecord" title="SelectRecord" />
<img src="/msexcel/img/img-mso-FormatPainter.png" alt="FormatPainter" title="FormatPainter" />
<img src="/msexcel/img/img-mso-MoreTextureOptions.png" alt="MoreTextureOptions" title="MoreTextureOptions" />
<img src="/msexcel/img/img-mso-TextPictureFill.png" alt="TextPictureFill" title="TextPictureFill" />
<img src="/msexcel/img/img-mso-PictureBulletsInsert.png" alt="PictureBulletsInsert" title="PictureBulletsInsert" />
---
<img src="/msexcel/img/img-mso-Undo.png" alt="Undo" title="Undo" />
<img src="/msexcel/img/img-mso-Redo.png" alt="Redo" title="Redo" />
<img src="/msexcel/img/img-mso-ShapeCloud.png" alt="ShapeCloud" title="ShapeCloud" />
<img src="/msexcel/img/img-mso-ShapeSeal8.png" alt="ShapeSeal8" title="ShapeSeal8" />
<img src="/msexcel/img/img-mso-VisibilityHidden.png" alt="VisibilityHidden" title="VisibilityHidden" />
<img src="/msexcel/img/img-mso-ShowClipboard.png" alt="ShowClipboard" title="ShowClipboard" />
<img src="/msexcel/img/img-mso-PanAndZoomWindow.png" alt="PanAndZoomWindow" title="PanAndZoomWindow" />
<img src="/msexcel/img/img-mso-LassoSelect.png" alt="LassoSelect" title="LassoSelect" />
<img src="/msexcel/img/img-mso-PanningHand.png" alt="PanningHand" title="PanningHand" />
<img src="/msexcel/img/img-mso-BlogCategoriesRefresh.png" alt="BlogCategoriesRefresh" title="BlogCategoriesRefresh" />
<img src="/msexcel/img/img-mso-MsnLogo.png" alt="MsnLogo" title="MsnLogo" />
<img src="/msexcel/img/img-mso-LinkBarCustom.png" alt="LinkBarCustom" title="LinkBarCustom" />
<img src="/msexcel/img/img-mso-BlackAndWhiteDontShow.png" alt="BlackAndWhiteDontShow" title="BlackAndWhiteDontShow" />
<img src="/msexcel/img/img-mso-CellFillColorPicker.png" alt="CellFillColorPicker" title="CellFillColorPicker" />
<img src="/msexcel/img/img-mso-ControlTabControl.png" alt="ControlTabControl" title="ControlTabControl" />
<img src="/msexcel/img/img-mso-FileNewDefault.png" alt="FileNewDefault" title="FileNewDefault" />
<img src="/msexcel/img/img-mso-FileOpenUsingBackstage.png" alt="FileOpenUsingBackstage" title="FileOpenUsingBackstage" />
<img src="/msexcel/img/img-mso-FilePrintQuick.png" alt="FilePrintQuick" title="FilePrintQuick" />
<img src="/msexcel/img/img-mso-FileSendAsAttachment.png" alt="FileSendAsAttachment" title="FileSendAsAttachment" />
<img src="/msexcel/img/img-mso-FontPaneShowHide.png" alt="FontPaneShowHide" title="FontPaneShowHide" />
<img src="/msexcel/img/img-mso-HeaderCell.png" alt="HeaderCell" title="HeaderCell" />
<img src="/msexcel/img/img-mso-OutlineSubtotals.png" alt="OutlineSubtotals" title="OutlineSubtotals" />
<img src="/msexcel/img/img-mso-PointerModeOptions.png" alt="PointerModeOptions" title="PointerModeOptions" />
<img src="/msexcel/img/img-mso-PrintPreviewAndPrint.png" alt="PrintPreviewAndPrint" title="PrintPreviewAndPrint" />
<img src="/msexcel/img/img-mso-Risks.png" alt="Risks" title="Risks" />
<img src="/msexcel/img/img-mso-StartAfterPrevious.png" alt="StartAfterPrevious" title="StartAfterPrevious" />

---

## Resources
- [ImageMSO Gallery](https://bert-toolkit.com/imagemso-list.html)

<!-- Links -->

[Install]: https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460#:~:text=COM%20add%2Din-,Click%20the%20File%20tab%2C%20click%20Options%2C%20and%20then%20click%20the,install%2C%20and%20then%20click%20OK.