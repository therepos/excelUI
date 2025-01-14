Sub RibbonOnLoad(Rib As IRibbonUI)
'   Purpose: Callback for customUI.onLoad
'   Reference: https://www.rondebruin.nl/win/s2/win012.htm
'   Note: 
'    -Public Ribbon As IRibbonUI
'   - Public MyTag As String
'   - Public MySelectedFont As String
'   - Public MySelectedFontSize As String
'   Updated: 2022FEB25

    Set Ribbon = Rib
    MyTag = "tabMain"
    MySelectedFont = "ddSelectionFont01"
    MySelectedFontSize = "ddSelectionFontSize03"
    
End Sub

