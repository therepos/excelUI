Sub FindReplaceAll()
'	Purpose: Find & Replace text/values throughout a specific sheet
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save
	
	Dim sht As Worksheet

' 	Store a specfic sheet to a variable
	Set sht = Sheets("json")

' 	Perform the Find/Replace All
	sht.Cells.Replace what:="flat_model:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="floor_area_sqm:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="remaining_lease_period:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="dist_mrt:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="dist_sch:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="dist_raffles:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="dist_mall:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="code_maturity:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="code_storey:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="code_type:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="code_town:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="add_lat:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="add_lon:", Replacement:="", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="{", Replacement:="[", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False
	sht.Cells.Replace what:="}", Replacement:="]", _
		LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
		SearchFormat:=False, ReplaceFormat:=False

ErrorHandler:
    Exit Sub

End Sub