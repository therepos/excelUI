Function XGETPAGENUMBER(CurrentCell As Range) As String
'   Purpose: Return page number of a cell
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim VPC As Integer, HPC As Integer
    Dim VerticalPageBreak As VPageBreak, HorizontalPageBreak As HPageBreak
    Dim NumPage As Integer

    If ActiveSheet.PageSetup.Order = xlDownThenOver Then
        HPC = ActiveSheet.HPageBreaks.Count + 1
        VPC = 1
    Else
        VPC = ActiveSheet.VPageBreaks.Count + 1
        HPC = 1
    End If

    NumPage = 1
    For Each VerticalPageBreak In ActiveSheet.VPageBreaks
      If VerticalPageBreak.Location.Column > CurrentCell.Cells.Column Then Exit For
      NumPage = NumPage + HPC
    Next VerticalPageBreak
    
    For Each HorizontalPageBreak In ActiveSheet.HPageBreaks
      If HorizontalPageBreak.Location.row > CurrentCell.Cells.row Then Exit For
      NumPage = NumPage + VPC
    Next HorizontalPageBreak

    XGETPAGENUMBER = NumPage

End Function

