# WIP

### Read to File
```vb
Sub ReadDelimitedTextFileIntoArray()
'   Reference: https://www.automateexcel.com/vba/read-import-text-file/
'   Reference: https://www.automateexcel.com/vba/write-to-text-file/

    Dim Delimiter As String
    Dim TextFile As Integer
    Dim FilePath As String
    Dim FileContent As String
    Dim LineArray() As String
    Dim DataArray() As String
    Dim TempArray() As String
    Dim rw As Long, col As Long
 
    Delimiter = vbTab 'the delimiter that is used in your text file
    FilePath = ThisWorkbook.Path & "\TestFileTab.txt"
    rw = 1
    
    TextFile = FreeFile
    Open FilePath For Input As TextFile
    FileContent = Input(LOF(TextFile), TextFile)
    Close TextFile
 
    LineArray() = Split(FileContent, vbNewLine) 'change vbNewLine to vbCrLf or vbLf depending on the line separator that is used in your text file
    For x = LBound(LineArray) To UBound(LineArray)
        If Len(Trim(LineArray(x))) <> 0 Then
           TempArray = Split(LineArray(x), Delimiter)
           col = UBound(TempArray)
   ReDim Preserve DataArray(col, rw)
           For y = LBound(TempArray) To UBound(TempArray)
       DataArray(y, rw) = TempArray(y)
       Cells(x + 1, y + 1).Value = DataArray(y, rw)  'this code will start pasting the text file’s content from the active worksheet’s A1 (Cell(1,1)) cell
           Next y
        End If
        rw = rw + 1
     Next x
 
End Sub
```
### Install AddIn
 Chip Pearson
```vb
Sub Install_Addin()

   Dim AI as excel.addin
   Set AI = Application.Addins.Add("C:Add_In.xlam")
   AI.Installed = True
   Application.Addins("Add_in").Installed = True

End Sub
```
