Function XCELLFORMULA(rCell As Range) As String
'   Purpose: Return cell formula
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save

    XCELLFORMULA = rCell.Formula

End Function

