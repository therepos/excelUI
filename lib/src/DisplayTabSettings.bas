Sub DisplayTabSettings(control As IRibbonControl)
'   Purpose: Display ribbon tab on demand (Selection)
'   Updated: 2022FEB25

    Call RibbonRefresh(Tag:="tabSettings", TabID:="tabSettings")
    
End Sub

