Sample VBA codes to further explore.

```vba

Option Explicit On

'*****************************************
' NOTE:
' Sample instructional commentaries
' Source: epace.xla
'*****************************************

Public Sub SetTool(ByVal bItem As Document, ByRef Cancel As Boolean)

    Dim strContext As String
    strContext = "SetTool(bItem=" & bItem & ")"

    On Error GoTo ERROR_HANDLER
    With ScopeTimer(strContext)
        Dim oDllFacade As Object
        Set oDllFacade = GetDllFacade
        oDllFacade.SetTool(bItem)
        Set oEventHandler = New EventHandler
        Set oEventHandler.cmdApplication = Application
        Set oEventHandler.cmdPrintToolbarButton = CommandBars("Standard").Controls("&Print")
        
        If oEventHandler Is Nothing Then
            HandleError 0, "oEventHandler error", strFunctionName
        End If

	#If VBA7 And Win64 Then
            Dim hKey As LongPtr
	#Else
            Dim hKey As Long
	#End If

        Public Const REG_SZ As Long = 1
        Public Const REG_DWORD As Long = 2
        Dim sValue As String
        Dim vValue As Variant

        Select Case iType
            Case REG_SZ
                sValue = vValue & Chr$(0)
            Case REG_DWORD
                sValue = vValue
        End Select

PROC_EXIT:
        Set oDllFacade = Nothing

    End With
    Exit Sub

ERROR_HANDLER:
    If (Err <> 0) Then
        HandleError Err.Number, Err.Description, strContext
    Else
        HandleError 0, "Unknown Error", strContext
    End If

End Sub

```

