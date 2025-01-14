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

