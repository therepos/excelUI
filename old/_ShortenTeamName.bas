Sub ShortenTeamName()
' 	Purpose: Replace team names with shorter name
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Selection.Replace "Liverpool", "LIVE"
    Selection.Replace "Manchester City", "MANC"
    Selection.Replace "Leicester City", "LEIC"
    Selection.Replace "Chelsea", "CHEL"
    Selection.Replace "Manchester United", "MANU"
    Selection.Replace "Wolverhampton Wanderers", "WOLV"
    Selection.Replace "Tottenham Hotspur", "TOTT"
    Selection.Replace "Sheffield United", "SHEF"
    Selection.Replace "Crystal Palace", "CRYS"
    Selection.Replace "Arsenal", "ARSE"
    Selection.Replace "Burnley", "BURN"
    Selection.Replace "Everton", "EVER"
    Selection.Replace "Newcastle United", "NEWU"
    Selection.Replace "Southampton", "SOUT"
    Selection.Replace "Brighton & Hove Albion", "BRIG"
    Selection.Replace "Watford", "WATF"
    Selection.Replace "West Ham United", "WEST"
    Selection.Replace "AFC Bournemouth", "AFCB"
    Selection.Replace "Aston Villa", "ASTO"
    Selection.Replace "Norwich City", "NORW"
    Selection.Replace "Man City", "MANC"
    Selection.Replace "Leicester", "LEIC"
    Selection.Replace "Man Utd", "MANU"
    Selection.Replace "Manchester Utd", "MANU"
    Selection.Replace "Wolves", "WOLV"
    Selection.Replace "Tottenham", "TOTT"
    Selection.Replace "Sheffield Utd", "SHEF"
    Selection.Replace "Newcastle", "NEWU"
    Selection.Replace "Newcastle Utd", "NEWU"
    Selection.Replace "Southampton", "SOUT"
    Selection.Replace "Brighton", "BRIG"
    Selection.Replace "West Ham", "WEST"
    Selection.Replace "Bournemouth", "AFCB"
    Selection.Replace "Norwich", "NORW"

ErrorHandler:
    Exit Sub

End Sub