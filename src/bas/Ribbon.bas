Public RibbonUI As IRibbonUI

Public Sub RibbonOnLoad(r As IRibbonUI)
    Set RibbonUI = r
End Sub

Public Sub RunByName(control As IRibbonControl)
    Dim macro As String
    macro = control.Tag
    If Len(macro) = 0 Then macro = control.ID
    On Error GoTo errh
    Application.Run macro
    Exit Sub
errh:
    MsgBox "Macro not found: " & macro, vbExclamation
End Sub

Public Sub GetHighlightLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastHighlight", "Green")
End Sub

Public Sub GetWorkbookFontLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastWbFont", "Arial")
End Sub

Public Sub GetWorkbookFontSizeLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastWbFontSize", "10")
End Sub

Public Sub GetSheetFontLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastShFont", "Arial")
End Sub

Public Sub GetSheetTabLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastShTab", "Green")
End Sub

Public Sub GetSelNumberLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelNumber", "Accounting")
End Sub

Public Sub GetSelCaseLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelCase", "Proper")
End Sub
