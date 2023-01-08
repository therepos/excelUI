Sub SheetPageBreakOff(control As IRibbonControl)
'   Purpose: This removes all page breaks for worksheet
'   Reference: www.DedicatedExcel.com
'   Updated: 2022MAR12

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.Activate
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub

