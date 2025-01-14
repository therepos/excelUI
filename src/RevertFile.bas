Sub RevertFile(control As IRibbonControl)
'   Purpose: Revert macro changes
'   Reference: https://www.excelforum.com/excel-programming-vba-macros/491103-undoing-a-macro.html
'   Reference: https://stackoverflow.com/questions/33813806/is-it-possible-to-undo-a-macro-action#:~:text=1)%20Have%20the%20macro%20save,did%20whatever%20the%20macro%20does.
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler

    wkname = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    ActiveWorkbook.Close SaveChanges:=False
    
    Workbooks.Open FileName:=wkname

ErrorHandler:
    Exit Sub

End Sub

