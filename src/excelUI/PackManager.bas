Attribute VB_Name = "PackManager"
' ============================================================
' MODULE: PackManager
' Add to your base excelUI.xlam as a new module
' ============================================================

Private pPackWorkbook As Workbook

Sub PackLoad()
    
    ' Check if a pack is already loaded
    If Not pPackWorkbook Is Nothing Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("A pack is already loaded: " & pPackWorkbook.Name & vbCrLf & vbCrLf & _
            "Do you want to unload it first and load a different one?", _
            vbYesNo + vbQuestion, "Pack Manager")
        If answer = vbNo Then Exit Sub
        pPackWorkbook.Close SaveChanges:=False
        Set pPackWorkbook = Nothing
    End If
    
    ' Prompt user to select a pack
    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Add-in Files (*.xlam), *.xlam", _
        Title:="Select a Function Pack to load")
    
    If filePath = "False" Then Exit Sub
    
    On Error GoTo LoadError
    Set pPackWorkbook = Workbooks.Open( _
        Filename:=filePath, _
        ReadOnly:=True, _
        AddToMru:=False)
    
    MsgBox "Pack loaded: " & pPackWorkbook.Name & vbCrLf & vbCrLf & _
        "A new tab should now be visible on the ribbon.", _
        vbInformation, "Pack Manager"
    Exit Sub
    
LoadError:
    MsgBox "Failed to load pack:" & vbCrLf & Err.Description, _
        vbExclamation, "Pack Manager"
    Set pPackWorkbook = Nothing
End Sub

Sub PackUnload()
    
    If pPackWorkbook Is Nothing Then
        MsgBox "No pack is currently loaded.", vbInformation, "Pack Manager"
        Exit Sub
    End If
    
    Dim packName As String
    packName = pPackWorkbook.Name
    pPackWorkbook.Close SaveChanges:=False
    Set pPackWorkbook = Nothing
    
    MsgBox "Pack unloaded: " & packName, vbInformation, "Pack Manager"
    
End Sub

