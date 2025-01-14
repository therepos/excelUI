```vba
Sub Macro1()
'   Purpose: Copy and paste a range to another workbook
'   Source: https://www.excelcampus.com/vba/copy-paste-another-workbook/
'   Source: https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet.cells

Dim wsc As Worksheet    ' worksheet to copy from
Dim wsp As Worksheet    ' worksheet to paste to
Dim wscLastRow As Long
Dim wspLastRow As Long

Set wsc = ActiveSheet
Set wsp = Workbooks("Modules Tracker.xlsx").Worksheets("L_Rename")

' wspLastRow = wsp.Cells(wsp.Rows.Count, "A").End(xlUp).Offset(1).Row

' Workbooks.Open "C:\Users\XP13R\Google Drive\Modules Tracker.xlsx"
                    
    wscLastRow = wsc.Range("A" & Rows.Count).End(xlUp).row
    Range("A1:A" & wscLastRow).Select
    Selection.Replace What:="~?", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=":", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    ' Clear contents of destination range
    wspLastRow = wsp.Cells(wsp.Rows.Count, "A").End(xlUp).Offset(1).row
    wsp.Range("A3:A" & wspLastRow).ClearContents
    
    ' Copy and paste to destination
    wsc.Range("A1:A" & wscLastRow).Copy wsp.Range("A3")
    ActiveWorkbook.Close savechanges:=False  
    
End Sub
```