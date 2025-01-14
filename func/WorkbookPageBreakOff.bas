Sub WorkbookPageBreakOff(control As IRibbonControl)
'   Purpose: This removes all page breaks for all worksheets in the workbook
'   Reference: www.DedicatedExcel.com
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
        
    For Each ws In Sheets
        ws.DisplayPageBreaks = False
    Next ws
 
     For Each ws In Sheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws

    Call sourceSheet.Activate

ErrorHandler:
    Exit Sub

End Sub

