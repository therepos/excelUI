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
<img src="/msexcel/img/img-mso-ListMacros.png" alt="ListMacros" title="ListMacros" />
<img src="/msexcel/img/img-mso-CancelRequest.png" alt="CancelRequest" title="CancelRequest" />
<img src="/msexcel/img/img-mso-Info.png" alt="Info" title="Info" />
<img src="/msexcel/img/img-mso-DiagramTargetInsertClassic.png" alt="DiagramTargetInsertClassic" title="DiagramTargetInsertClassic" />
<img src="/msexcel/img/img-mso-Risks.png" alt="Risks" title="Risks" />
<img src="/msexcel/img/img-mso-HighImportance.png" alt="HighImportance" title="HighImportance" />
<img src="/msexcel/img/img-mso-TrustCenter.png" alt="TrustCenter" title="TrustCenter" />
<img src="/msexcel/img/img-mso-FileSave.png" alt="FileSave" title="FileSave" />
<img src="/msexcel/img/img-mso-SourceControlCheckOut.png" alt="SourceControlCheckOut" title="SourceControlCheckOut" />
<img src="/msexcel/img/img-mso-SourceControlCheckIn.png" alt="SourceControlCheckIn" title="SourceControlCheckIn" />
<img src="/msexcel/img/img-mso-Folder.png" alt="Folder" title="Folder" />
<img src="/msexcel/img/img-mso-FileNew.png" alt="FileNew" title="FileNew" />
<img src="/msexcel/img/img-mso-CopyFolder.png" alt="CopyFolder" title="CopyFolder" />
<img src="/msexcel/img/img-mso-CreateMailRule.png" alt="CreateMailRule" title="CreateMailRule" />
<img src="/msexcel/img/img-mso-FilePrint.png" alt="FilePrint" title="FilePrint" />
<img src="/msexcel/img/img-mso-InsertDrawingCanvas.png" alt="InsertDrawingCanvas" title="InsertDrawingCanvas" />
<img src="/msexcel/img/img-mso-AutoFormat.png" alt="AutoFormat" title="AutoFormat" />
<img src="/msexcel/img/img-mso-OpenStartPage.png" alt="OpenStartPage" title="OpenStartPage" />
<img src="/msexcel/img/img-mso-SmartArtChangeColorsGallery.png" alt="SmartArtChangeColorsGallery" title="SmartArtChangeColorsGallery" />
<img src="/msexcel/img/img-mso-OutlinePromoteToHeading.png" alt="OutlinePromoteToHeading" title="OutlinePromoteToHeading" />
<img src="/msexcel/img/img-mso-OutlineDemoteToBodyText.png" alt="OutlineDemoteToBodyText" title="OutlineDemoteToBodyText" />
<img src="/msexcel/img/img-mso-LeftArrow2.png" alt="LeftArrow2" title="LeftArrow2" />
<img src="/msexcel/img/img-mso-OutlineMoveUp.png" alt="OutlineMoveUp" title="OutlineMoveUp" />
<img src="/msexcel/img/img-mso-OutlineMoveDown.png" alt="OutlineMoveDown" title="OutlineMoveDown" />
<img src="/msexcel/img/img-mso-RightArrow2.png" alt="RightArrow2" title="RightArrow2" />
<img src="/msexcel/img/img-mso-ViewGoBack.png" alt="ViewGoBack" title="ViewGoBack" />
<img src="/msexcel/img/img-mso-Pushpin.png" alt="Pushpin" title="Pushpin" />
<img src="/msexcel/img/img-mso-Lock.png" alt="Lock" title="Lock" />
<img src="/msexcel/img/img-mso-AdpPrimaryKey.png" alt="AdpPrimaryKey" title="AdpPrimaryKey" />
<img src="/msexcel/img/img-mso-MacroDefault.png" alt="MacroDefault" title="MacroDefault" />
<img src="/msexcel/img/img-mso-FrontPageToggleBookmark.png" alt="FrontPageToggleBookmark" title="FrontPageToggleBookmark" />
<img src="/msexcel/img/img-mso-ViewFullScreenView.png" alt="ViewFullScreenView" title="ViewFullScreenView" />
<img src="/msexcel/img/img-mso-ZoomPrintPreviewExcel.png" alt="ZoomPrintPreviewExcel" title="ZoomPrintPreviewExcel" />
<img src="/msexcel/img/img-mso-FilterByResource.png" alt="FilterByResource" title="FilterByResource" />
<img src="/msexcel/img/img-mso-AddressBook.png" alt="AddressBook" title="AddressBook" />
<img src="/msexcel/img/img-mso-SetPertWeights.png" alt="SetPertWeights" title="SetPertWeights" />
<img src="/msexcel/img/img-mso-QueryUnionQuery.png" alt="QueryUnionQuery" title="QueryUnionQuery" />
<img src="/msexcel/img/img-mso-SpeechMicrophone.png" alt="SpeechMicrophone" title="SpeechMicrophone" />
<img src="/msexcel/img/img-mso-AudioNoteDelete.png" alt="AudioNoteDelete" title="AudioNoteDelete" />
<img src="/msexcel/img/img-mso-CondolatoryEvent.png" alt="CondolatoryEvent" title="CondolatoryEvent" />
<img src="/msexcel/img/img-mso-Head.png" alt="Head" title="Head" />
<img src="/msexcel/img/img-mso-StartAfterPrevious.png" alt="StartAfterPrevious" title="StartAfterPrevious" />
<img src="/msexcel/img/img-mso-DiagramRadialInsertClassic.png" alt="DiagramRadialInsertClassic" title="DiagramRadialInsertClassic" />
<img src="/msexcel/img/img-mso-Breakpoint.png" alt="Breakpoint" title="Breakpoint" />
<img src="/msexcel/img/img-mso-HappyFace.png" alt="HappyFace" title="HappyFace" />
<img src="/msexcel/img/img-mso-SadFace.png" alt="SadFace" title="SadFace" />
<img src="/msexcel/img/img-mso-Calculator.png" alt="Calculator" title="Calculator" />
<img src="/msexcel/img/img-mso-AutoDial.png" alt="AutoDial" title="AutoDial" />
<img src="/msexcel/img/img-mso-Piggy.png" alt="Piggy" title="Piggy" />
<img src="/msexcel/img/img-mso-MagicEightBall.png" alt="MagicEightBall" title="MagicEightBall" />
<img src="/msexcel/img/img-mso-DollarSign.png" alt="DollarSign" title="DollarSign" />
<img src="/msexcel/img/img-mso-VisibilityVisible.png" alt="VisibilityVisible" title="VisibilityVisible" />
<img src="/msexcel/img/img-mso-VisibilityHidden.png" alt="VisibilityHidden" title="VisibilityHidden" />
<img src="/msexcel/img/img-mso-AppointmentColorDialog.png" alt="AppointmentColorDialog" title="AppointmentColorDialog" />
<img src="/msexcel/img/img-mso-AppointmentColor0.png" alt="AppointmentColor0" title="AppointmentColor0" />
<img src="/msexcel/img/img-mso-AppointmentColor1.png" alt="AppointmentColor1" title="AppointmentColor1" />
<img src="/msexcel/img/img-mso-AppointmentColor2.png" alt="AppointmentColor2" title="AppointmentColor2" />
<img src="/msexcel/img/img-mso-AppointmentColor3.png" alt="AppointmentColor3" title="AppointmentColor3" />
<img src="/msexcel/img/img-mso-AppointmentColor4.png" alt="AppointmentColor4" title="AppointmentColor4" />
<img src="/msexcel/img/img-mso-AppointmentColor5.png" alt="AppointmentColor5" title="AppointmentColor5" />
<img src="/msexcel/img/img-mso-AppointmentColor6.png" alt="AppointmentColor6" title="AppointmentColor6" />
<img src="/msexcel/img/img-mso-AppointmentColor7.png" alt="AppointmentColor7" title="AppointmentColor7" />
<img src="/msexcel/img/img-mso-AppointmentColor8.png" alt="AppointmentColor8" title="AppointmentColor8" />
<img src="/msexcel/img/img-mso-AppointmentColor9.png" alt="AppointmentColor9" title="AppointmentColor9" />
<img src="/msexcel/img/img-mso-AppointmentColor10.png" alt="AppointmentColor10" title="AppointmentColor10" />
<img src="/msexcel/img/img-mso-AppointmentBusy.png" alt="AppointmentBusy" title="AppointmentBusy" />
<img src="/msexcel/img/img-mso-AppointmentOutOfOffice.png" alt="AppointmentOutOfOffice" title="AppointmentOutOfOffice" />
<img src="/msexcel/img/img-mso-BlackAndWhiteAutomatic.png" alt="BlackAndWhiteAutomatic" title="BlackAndWhiteAutomatic" />
<img src="/msexcel/img/img-mso-BlackAndWhiteBlack.png" alt="BlackAndWhiteBlack" title="BlackAndWhiteBlack" />
<img src="/msexcel/img/img-mso-BlackAndWhiteBlackWithGrayscaleFill.png" alt="BlackAndWhiteBlackWithGrayscaleFill" title="BlackAndWhiteBlackWithGrayscaleFill" />
<img src="/msexcel/img/img-mso-BlackAndWhiteBlackWithWhiteFill.png" alt="BlackAndWhiteBlackWithWhiteFill" title="BlackAndWhiteBlackWithWhiteFill" />
<img src="/msexcel/img/img-mso-BlackAndWhiteDontShow.png" alt="BlackAndWhiteDontShow" title="BlackAndWhiteDontShow" />
<img src="/msexcel/img/img-mso-BlackAndWhiteGrayscale.png" alt="BlackAndWhiteGrayscale" title="BlackAndWhiteGrayscale" />
<img src="/msexcel/img/img-mso-BlackAndWhiteWhite.png" alt="BlackAndWhiteWhite" title="BlackAndWhiteWhite" />
<img src="/msexcel/img/img-mso-BlackAndWhiteInverseGrayscale.png" alt="BlackAndWhiteInverseGrayscale" title="BlackAndWhiteInverseGrayscale" />
<img src="/msexcel/img/img-mso-ViewDisplayInHighContrast.png" alt="ViewDisplayInHighContrast" title="ViewDisplayInHighContrast" />
<img src="/msexcel/img/img-mso-ViewGoForward.png" alt="ViewGoForward" title="ViewGoForward" />
<img src="/msexcel/img/img-mso-Chart3DColumnChart.png" alt="Chart3DColumnChart" title="Chart3DColumnChart" />
<img src="/msexcel/img/img-mso-Chart3DConeChart.png" alt="Chart3DConeChart" title="Chart3DConeChart" />
<img src="/msexcel/img/img-mso-ChartAreaChart.png" alt="ChartAreaChart" title="ChartAreaChart" />
<img src="/msexcel/img/img-mso-Chart3DBarChart.png" alt="Chart3DBarChart" title="Chart3DBarChart" />
<img src="/msexcel/img/img-mso-Chart3DPieChart.png" alt="Chart3DPieChart" title="Chart3DPieChart" />
<img src="/msexcel/img/img-mso-ChartTypeOtherInsertGallery.png" alt="ChartTypeOtherInsertGallery" title="ChartTypeOtherInsertGallery" />
<img src="/msexcel/img/img-mso-DatabaseCopyDatabaseFile.png" alt="DatabaseCopyDatabaseFile" title="DatabaseCopyDatabaseFile" />
<img src="/msexcel/img/img-mso-CategoryCollapse.png" alt="CategoryCollapse" title="CategoryCollapse" />
<img src="/msexcel/img/img-mso-FormulaMoreFunctionsMenu.png" alt="FormulaMoreFunctionsMenu" title="FormulaMoreFunctionsMenu" />
<img src="/msexcel/img/img-mso-VisualBasicReferences.png" alt="VisualBasicReferences" title="VisualBasicReferences" />
<img src="/msexcel/img/img-mso-PictureReflectionGalleryItem.png" alt="PictureReflectionGalleryItem" title="PictureReflectionGalleryItem" />
<img src="/msexcel/img/img-mso-CreateReportBlankReport.png" alt="CreateReportBlankReport" title="CreateReportBlankReport" />
<img src="/msexcel/img/img-mso-SlideMasterClipArtPlaceholderInsert.png" alt="SlideMasterClipArtPlaceholderInsert" title="SlideMasterClipArtPlaceholderInsert" />
<img src="/msexcel/img/img-mso-SlideMasterMediaPlaceholderInsert.png" alt="SlideMasterMediaPlaceholderInsert" title="SlideMasterMediaPlaceholderInsert" />
<img src="/msexcel/img/img-mso-BlackAndWhite.png" alt="BlackAndWhite" title="BlackAndWhite" />
<img src="/msexcel/img/img-mso-DataGraphicIconSet.png" alt="DataGraphicIconSet" title="DataGraphicIconSet" />
<img src="/msexcel/img/img-mso-CharacterBorder.png" alt="CharacterBorder" title="CharacterBorder" />
<img src="/msexcel/img/img-mso-CharacterShading.png" alt="CharacterShading" title="CharacterShading" />
<img src="/msexcel/img/img-mso-Delete.png" alt="Delete" title="Delete" />
<img src="/msexcel/img/img-mso-TagMarkComplete.png" alt="TagMarkComplete" title="TagMarkComplete" />
<img src="/msexcel/img/img-mso-TableDesign.png" alt="TableDesign" title="TableDesign" />
<img src="/msexcel/img/img-mso-RecurrenceEdit.png" alt="RecurrenceEdit" title="RecurrenceEdit" />
<img src="/msexcel/img/img-mso-EquationMatrixGallery.png" alt="EquationMatrixGallery" title="EquationMatrixGallery" />
<img src="/msexcel/img/img-mso-EquationOptions.png" alt="EquationOptions" title="EquationOptions" />
<img src="/msexcel/img/img-mso-ShapeFillColorPickerClassic.png" alt="ShapeFillColorPickerClassic" title="ShapeFillColorPickerClassic" />
<img src="/msexcel/img/img-mso-FindDialog.png" alt="FindDialog" title="FindDialog" />
<img src="/msexcel/img/img-mso-FormatPainter.png" alt="FormatPainter" title="FormatPainter" />
<img src="/msexcel/img/img-mso-Bullets.png" alt="Bullets" title="Bullets" />
<img src="/msexcel/img/img-mso-ResultsPaneStartFindAndReplace.png" alt="ResultsPaneStartFindAndReplace" title="ResultsPaneStartFindAndReplace" />
<img src="/msexcel/img/img-mso-PositionAbsoluteMarks.png" alt="PositionAbsoluteMarks" title="PositionAbsoluteMarks" />
<img src="/msexcel/img/img-mso-ControlsGallery.png" alt="ControlsGallery" title="ControlsGallery" />
<img src="/msexcel/img/img-mso-LassoSelect.png" alt="LassoSelect" title="LassoSelect" />
<img src="/msexcel/img/img-mso-ShapesDuplicate.png" alt="ShapesDuplicate" title="ShapesDuplicate" />
<img src="/msexcel/img/img-mso-RunDialog.png" alt="RunDialog" title="RunDialog" />
<img src="/msexcel/img/img-mso-ControlTabControl.png" alt="ControlTabControl" title="ControlTabControl" />
<img src="/msexcel/img/img-mso-InkEraseMode.png" alt="InkEraseMode" title="InkEraseMode" />
<img src="/msexcel/img/img-mso-WordArtFormatDialog.png" alt="WordArtFormatDialog" title="WordArtFormatDialog" />
<img src="/msexcel/img/img-mso-MsnLogo.png" alt="MsnLogo" title="MsnLogo" />
<img src="/msexcel/img/img-mso-NewAlert.png" alt="NewAlert" title="NewAlert" />
<img src="/msexcel/img/img-mso-GoToMail.png" alt="GoToMail" title="GoToMail" />
<img src="/msexcel/img/img-mso-EnvelopesAndLabelsDialog.png" alt="EnvelopesAndLabelsDialog" title="EnvelopesAndLabelsDialog" />
<img src="/msexcel/img/img-mso-RmsSendBizcard.png" alt="RmsSendBizcard" title="RmsSendBizcard" />
<img src="/msexcel/img/img-mso-RmsSendBizcardDesign.png" alt="RmsSendBizcardDesign" title="RmsSendBizcardDesign" />
<img src="/msexcel/img/img-mso-PersonaStatusAway.png" alt="PersonaStatusAway" title="PersonaStatusAway" />
<img src="/msexcel/img/img-mso-PersonaStatusBusy.png" alt="PersonaStatusBusy" title="PersonaStatusBusy" />
<img src="/msexcel/img/img-mso-PersonaStatusOnline.png" alt="PersonaStatusOnline" title="PersonaStatusOnline" />
<img src="/msexcel/img/img-mso-RmsInvokeBrowser.png" alt="RmsInvokeBrowser" title="RmsInvokeBrowser" />
<img src="/msexcel/img/img-mso-RmsNavigationBar.png" alt="RmsNavigationBar" title="RmsNavigationBar" />
<img src="/msexcel/img/img-mso-SelectionPaneHidden.png" alt="SelectionPaneHidden" title="SelectionPaneHidden" />
<img src="/msexcel/img/img-mso-WebServerDiscussions.png" alt="WebServerDiscussions" title="WebServerDiscussions" />
<img src="/msexcel/img/img-mso-ChangeBinding.png" alt="ChangeBinding" title="ChangeBinding" />
<img src="/msexcel/img/img-mso-ShowClipboard.png" alt="ShowClipboard" title="ShowClipboard" />
<img src="/msexcel/img/img-mso-_3DStyle.png" alt="_3DStyle" title="_3DStyle" />
<img src="/msexcel/img/img-mso-HyperlinksVerify.png" alt="HyperlinksVerify" title="HyperlinksVerify" />
<img src="/msexcel/img/img-mso-ToolboxVideo.png" alt="ToolboxVideo" title="ToolboxVideo" />
<img src="/msexcel/img/img-mso-TableSelect.png" alt="TableSelect" title="TableSelect" />
<img src="/msexcel/img/img-mso-WindowUnhide.png" alt="WindowUnhide" title="WindowUnhide" />
<img src="/msexcel/img/img-mso-WindowsCascade.png" alt="WindowsCascade" title="WindowsCascade" />
<img src="/msexcel/img/img-mso-SignatureShow.png" alt="SignatureShow" title="SignatureShow" />
<img src="/msexcel/img/img-mso-OutlineSubtotals.png" alt="OutlineSubtotals" title="OutlineSubtotals" />
<img src="/msexcel/img/img-mso-SelectRecord.png" alt="SelectRecord" title="SelectRecord" />
<img src="/msexcel/img/img-mso-SlideShowUseRehearsedTimings.png" alt="SlideShowUseRehearsedTimings" title="SlideShowUseRehearsedTimings" />
<img src="/msexcel/img/img-mso-MailMergeResultsPreview.png" alt="MailMergeResultsPreview" title="MailMergeResultsPreview" />
<img src="/msexcel/img/img-mso-PanAndZoomWindow.png" alt="PanAndZoomWindow" title="PanAndZoomWindow" />
<img src="/msexcel/img/img-mso-SignaturesLoading.png" alt="SignaturesLoading" title="SignaturesLoading" />
<img src="/msexcel/img/img-mso-TableBorderPenColorPicker.png" alt="TableBorderPenColorPicker" title="TableBorderPenColorPicker" />
<img src="/msexcel/img/img-mso-SlidesPerPageSlideOutline.png" alt="SlidesPerPageSlideOutline" title="SlidesPerPageSlideOutline" />
<img src="/msexcel/img/img-mso-AnimationOnClick.png" alt="AnimationOnClick" title="AnimationOnClick" />
<img src="/msexcel/img/img-mso-ShapeFillTextureGallery.png" alt="ShapeFillTextureGallery" title="ShapeFillTextureGallery" />
<img src="/msexcel/img/img-mso-DesignXml.png" alt="DesignXml" title="DesignXml" />
<img src="/msexcel/img/img-mso-FontSchemes.png" alt="FontSchemes" title="FontSchemes" />
<img src="/msexcel/img/img-mso-SmartArtLayoutGallery.png" alt="SmartArtLayoutGallery" title="SmartArtLayoutGallery" />
<img src="/msexcel/img/img-mso-HorizontalSpacingDecrease.png" alt="HorizontalSpacingDecrease" title="HorizontalSpacingDecrease" />
<img src="/msexcel/img/img-mso-HorizontalSpacingIncrease.png" alt="HorizontalSpacingIncrease" title="HorizontalSpacingIncrease" />
<img src="/msexcel/img/img-mso-ObjectEditPoints.png" alt="ObjectEditPoints" title="ObjectEditPoints" />
<img src="/msexcel/img/img-mso-ObjectsGroup.png" alt="ObjectsGroup" title="ObjectsGroup" />
<img src="/msexcel/img/img-mso-ObjectsUngroup.png" alt="ObjectsUngroup" title="ObjectsUngroup" />
<img src="/msexcel/img/img-mso-ChangesDiscardAndRefresh.png" alt="ChangesDiscardAndRefresh" title="ChangesDiscardAndRefresh" />
<img src="/msexcel/img/img-mso-FrameCreateAbove.png" alt="FrameCreateAbove" title="FrameCreateAbove" />
<img src="/msexcel/img/img-mso-FrameCreateBelow.png" alt="FrameCreateBelow" title="FrameCreateBelow" />
<img src="/msexcel/img/img-mso-FrameCreateLeft.png" alt="FrameCreateLeft" title="FrameCreateLeft" />
<img src="/msexcel/img/img-mso-FrameCreateRight.png" alt="FrameCreateRight" title="FrameCreateRight" />
<img src="/msexcel/img/img-mso-ViewSlideSorterView.png" alt="ViewSlideSorterView" title="ViewSlideSorterView" />
<img src="/msexcel/img/img-mso-SlidesPerPage9Slides.png" alt="SlidesPerPage9Slides" title="SlidesPerPage9Slides" />
<img src="/msexcel/img/img-mso-ViewSheetGridlines.png" alt="ViewSheetGridlines" title="ViewSheetGridlines" />
<img src="/msexcel/img/img-mso-GridSettings.png" alt="GridSettings" title="GridSettings" />
<img src="/msexcel/img/img-mso-ViewGridlinesFrontPage.png" alt="ViewGridlinesFrontPage" title="ViewGridlinesFrontPage" />
<img src="/msexcel/img/img-mso-Redo.png" alt="Redo" title="Redo" />
<img src="/msexcel/img/img-mso-Undo.png" alt="Undo" title="Undo" />
<img src="/msexcel/img/img-mso-ModuleInsert.png" alt="ModuleInsert" title="ModuleInsert" />
<img src="/msexcel/img/img-mso-_3DPerspectiveDecrease.png" alt="_3DPerspectiveDecrease" title="_3DPerspectiveDecrease" />
<img src="/msexcel/img/img-mso-_3DPerspectiveIncrease.png" alt="_3DPerspectiveIncrease" title="_3DPerspectiveIncrease" />
<img src="/msexcel/img/img-mso-LinkBarCustom.png" alt="LinkBarCustom" title="LinkBarCustom" />
<img src="/msexcel/img/img-mso-ShapeCloud.png" alt="ShapeCloud" title="ShapeCloud" />
<img src="/msexcel/img/img-mso-RecordsRefreshMenu.png" alt="RecordsRefreshMenu" title="RecordsRefreshMenu" />

---

## Resources
- [ImageMSO Gallery](https://bert-toolkit.com/imagemso-list.html)

<!-- Links -->

[Install]: https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460#:~:text=COM%20add%2Din-,Click%20the%20File%20tab%2C%20click%20Options%2C%20and%20then%20click%20the,install%2C%20and%20then%20click%20OK.