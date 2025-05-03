Attribute VB_Name = "Ribbon"
Public Ribbon As IRibbonUI
Public MyTag As String
Public MySelectedFont As String
Public MySelectedFontSize As String

Private Sub CheckExpiry()
'   Purpose: Unload Excel add-in if product expired
'   Reference: https://www.automateexcel.com/vba/date-variable/
'   Reference: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._workbook.builtindocumentproperties?redirectedfrom=MSDN&view=excel-pia#Microsoft_Office_Interop_Excel__Workbook_BuiltinDocumentProperties
'   Notes:
'   - %SystemRoot%\System32
'   - %UserProfile%\Application Data\Microsoft\Office\Recent
'   - Assumes GetUTCTimeDate() always return "12:00:00 am" when user offline
'   - ThisWorkbook.BuiltinDocumentProperties("Creation Date")

    Dim expiryDate, internetDate As Date
    Dim FileName, filepath As String
    expiryDate = DateSerial(2022, 12, 31)
    internetDate = GetUTCTimeDate()

'   Error handling for unavailable internet time
    If internetDate = TimeValue("12:00:00 am") Then
        internetDate = DateSerial(2022, 12, 31)
    End If

    If expiryDate <= internetDate Then
        MsgBox "Trial period expired." & vbNewLine & "Please visit our website to update the addin."
'   Disabled RemoveAddin for master code
        'Call RemoveAddin
    End If
        
End Sub

Private Function GetUTCTimeDate() As Date
'   Purpose: Get Internet Time
'   Reference: https://stackoverflow.com/questions/48371398/get-date-from-internet-and-compare-to-system-clock-on-workbook-open
'   Reference: https://stackoverflow.com/questions/551613/check-for-active-internet-connection
'   Reference: http://excelerator.solutions/2017/08/28/excel-http-get-request/
'   Requirement: Microsoft Scripting Runtime, Microsoft Internet Controls, and Microsoft WinHTTP

    Dim UTCDateTime As String
    Dim arrDT() As String
    Dim http As Object
    Dim UTCDate As String
    Dim UTCTime As String

    Const NetTime As String = "https://www.time.gov/"
    
    On Error Resume Next
    Set http = CreateObject("Microsoft.XMLHTTP")
    On Error GoTo 0

    On Error Resume Next
    http.Open "GET", NetTime & Now(), False, "", ""
    http.send
    
    UTCDateTime = http.GetResponseHeader("Date")
    UTCDate = Mid(UTCDateTime, InStr(UTCDateTime, ",") + 2)
    UTCDate = Left(UTCDate, InStrRev(UTCDate, " ") - 1)
    UTCTime = Mid(UTCDate, InStrRev(UTCDate, " ") + 1)
    UTCDate = Left(UTCDate, InStrRev(UTCDate, " ") - 1)
    GetUTCTimeDate = DateValue(UTCDate)
    On Error GoTo 0

End Function

Private Sub RemoveAddin()
'   Purpose: Remove Excel add-in programmatically
'   Reference: https://answers.microsoft.com/en-us/msoffice/forum/all/remove-an-excel-2007-addin-programatically-in-vba/2fffffdf-dfc9-4723-8924-66a08e4b62ac
         
    Dim FileName As String
    Dim A As AddIn
    
    FileName = ThisWorkbook.Name
    
    For Each A In Application.AddIns
        If A.Name = FileName Then
            If A.Installed = True Then
                A.Installed = False
            Else
                Exit For
            End If
            Exit For
        End If
    Next
    Workbooks(FileName).Close SaveChanges:=False
    
End Sub

Sub RibbonOnLoad(Rib As IRibbonUI)
'   Purpose: Callback for customUI.onLoad
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Updated: 2022FEB25

    Set Ribbon = Rib
    MyTag = "tabMain"
    MySelectedFont = "ddSelectionFont01"
    MySelectedFontSize = "ddSelectionFontSize03"
    
End Sub

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

Sub GetVisible(control As IRibbonControl, ByRef visible)
'   Purpose: See the ThisWorkbook module for a option to Show by default
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Updated: 2022FEB25

    If control.Tag Like MyTag Then
        visible = True
    Else
        visible = False
    End If

End Sub

Sub DisplayTabMain(control As IRibbonControl)
'   Purpose: Display ribbon tab on demand (Selection)
'   Updated: 2022FEB25

    Call RibbonRefresh(Tag:="tabMain", TabID:="tabMain")
    
End Sub

Sub DisplayTabSettings(control As IRibbonControl)
'   Purpose: Display ribbon tab on demand (Selection)
'   Updated: 2022FEB25

    Call RibbonRefresh(Tag:="tabSettings", TabID:="tabSettings")
    
End Sub

Sub GetDefaultFontID(ByRef control As IRibbonControl, ByRef returnedVal As Variant)
'   Purpose: Get default font selection
'   Updated: 2022FEB27

    returnedVal = MySelectedFont
    
End Sub

Sub GetDefaultFontSizeID(ByRef control As IRibbonControl, ByRef returnedVal As Variant)
'   Purpose: Get default font selection
'   Updated: 2022FEB27

    returnedVal = MySelectedFontSize
    
End Sub

Sub GetSelectedFontID(control As IRibbonControl, ID As String, index As Integer)
'   Purpose: Get user font selection
'   Updated: 2022FEB27

    MySelectedFont = ID

End Sub

Sub GetSelectedFontSizeID(control As IRibbonControl, ID As String, index As Integer)
'   Purpose: Get user font selection
'   Updated: 2022FEB27

    MySelectedFontSize = ID

End Sub

Sub GetLabelWorkbookFont(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFont
        Case "ddSelectionFont01": returnedVal = "Arial"
        Case "ddSelectionFont02": returnedVal = "Verdana"
        Case "ddSelectionFont03": returnedVal = "Times"
    End Select

End Sub


Sub GetLabelWorkbookFontSize(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": returnedVal = 8
        Case "ddSelectionFontSize02": returnedVal = 9
        Case "ddSelectionFontSize03": returnedVal = 10
        Case "ddSelectionFontSize04": returnedVal = 11
    End Select

End Sub

Sub GetLabelSelectionFont(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFont
        Case "ddSelectionFont01": returnedVal = "Arial"
        Case "ddSelectionFont02": returnedVal = "Verdana"
        Case "ddSelectionFont03": returnedVal = "Times"
    End Select

End Sub

Sub GetLabelSheetFont(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFont
        Case "ddSelectionFont01": returnedVal = "Arial"
        Case "ddSelectionFont02": returnedVal = "Verdana"
        Case "ddSelectionFont03": returnedVal = "Times"
    End Select

End Sub


Sub GetLabelSheetFontSize(control As IRibbonControl, ByRef returnedVal)
'   Purpose: Generate font label
'   Updated: 2022MAR12

    Select Case MySelectedFontSize
        Case "ddSelectionFontSize01": returnedVal = 8
        Case "ddSelectionFontSize02": returnedVal = 9
        Case "ddSelectionFontSize03": returnedVal = 10
        Case "ddSelectionFontSize04": returnedVal = 11
    End Select

End Sub
