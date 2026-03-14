Public RibbonUI As IRibbonUI

Public Sub RibbonOnLoad(R As IRibbonUI)
    Set RibbonUI = R
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

Public Sub GetSheetTabColor(control As IRibbonControl, ByRef returnedVal)
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastShTab", "Green")
    
    Dim clr As Long
    Select Case last
        Case "Green":  clr = RGB(204, 255, 204)
        Case "Red":    clr = RGB(255, 204, 204)
        Case "Yellow": clr = RGB(255, 255, 0)
        Case Else:     clr = RGB(255, 255, 255)
    End Select
    
    Set returnedVal = CreateColorIcon(clr)
    
End Sub

Public Sub GetHighlightColor(control As IRibbonControl, ByRef returnedVal)
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastHighlight", "Green")
    
    Dim clr As Long
    Select Case last
        Case "Green":  clr = RGB(204, 255, 204)
        Case "Red":    clr = RGB(255, 204, 204)
        Case "Yellow": clr = RGB(255, 255, 0)
        Case Else:     clr = RGB(255, 255, 255)
    End Select
    
    Set returnedVal = CreateColorIcon(clr)
    
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

Public Sub GetSelNumberLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelNumber", "Accounting")
End Sub

Public Sub GetSelCaseLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelCase", "Proper")
End Sub

Private Function CreateColorIcon(clr As Long) As IPictureDisp
    Dim tmp As String
    tmp = Environ("TEMP") & "\excelui_icon.bmp"
    
    ' Create a 16x16 BMP file
    Dim f As Integer
    f = FreeFile
    Open tmp For Binary As #f
    
    ' BMP header (54 bytes) + 16x16x3 pixel data (768 bytes) = 822 bytes
    Dim bmp(0 To 821) As Byte
    
    ' BM signature
    bmp(0) = 66: bmp(1) = 77
    ' File size = 822
    bmp(2) = 54: bmp(3) = 3: bmp(4) = 0: bmp(5) = 0
    ' Data offset = 54
    bmp(10) = 54
    ' DIB header size = 40
    bmp(14) = 40
    ' Width = 16
    bmp(18) = 16
    ' Height = 16
    bmp(22) = 16
    ' Planes = 1
    bmp(26) = 1
    ' Bits per pixel = 24
    bmp(28) = 24
    
    ' Fill pixels with color (BGR order)
    Dim R As Byte, G As Byte, B As Byte
    R = clr And &HFF
    G = (clr \ &H100) And &HFF
    B = (clr \ &H10000) And &HFF
    
    Dim i As Long
    For i = 54 To 821 Step 3
        bmp(i) = B
        bmp(i + 1) = G
        bmp(i + 2) = R
    Next i
    
    Put #f, , bmp
    Close #f
    
    Set CreateColorIcon = LoadPicture(tmp)
    Kill tmp
End Function
