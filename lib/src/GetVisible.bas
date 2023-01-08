Sub GetVisible(control As IRibbonControl, ByRef visible)
'   Purpose: See the ThisWorkbook module for a option to Show by default
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Updated: 2022FEB25

    If control.Tag Like MyTag Then
        visible = True
    Else
        visible = False
    End If

End Sub

