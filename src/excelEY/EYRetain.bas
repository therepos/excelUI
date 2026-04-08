Attribute VB_Name = "EYRetain"
' ============================================================
' RETAIN AVAILABILITY GENERATOR
' Auto-detects sheet structure and generates "Retain Availability"
' in a NEW workbook
' Works on any tab regardless of column layout
' ============================================================

' --- Helper: locale-safe DD/MM/YYYY parser ---
Private Function ParseDMY(ByVal s As String) As Date
    Dim parts() As String
    Dim sep As String
    If InStr(s, "/") > 0 Then
        sep = "/"
    ElseIf InStr(s, "-") > 0 Then
        sep = "-"
    ElseIf InStr(s, ".") > 0 Then
        sep = "."
    Else
        ParseDMY = CDate(s)
        Exit Function
    End If
    parts = Split(s, sep)
    If UBound(parts) < 2 Then
        ParseDMY = CDate(s)
        Exit Function
    End If
    ParseDMY = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
End Function

' --- Helper: find a column by header keywords ---
Private Function FindColumnByHeader(ws As Worksheet, headerRow As Long, lastCol As Long, ParamArray keywords() As Variant) As Long
    ' Returns the column number whose header matches any of the keywords (case-insensitive)
    ' Returns 0 if not found
    Dim c As Long
    Dim kw As Variant
    Dim cellText As String
    
    For c = 1 To lastCol
        If Not IsError(ws.Cells(headerRow, c).Value) Then
            cellText = LCase(Trim(CStr(ws.Cells(headerRow, c).Value & "")))
            For Each kw In keywords
                If cellText = LCase(CStr(kw)) Then
                    FindColumnByHeader = c
                    Exit Function
                End If
            Next kw
        End If
    Next c
    
    ' Fallback: partial match
    For c = 1 To lastCol
        If Not IsError(ws.Cells(headerRow, c).Value) Then
            cellText = LCase(Trim(CStr(ws.Cells(headerRow, c).Value & "")))
            For Each kw In keywords
                If InStr(cellText, LCase(CStr(kw))) > 0 Then
                    FindColumnByHeader = c
                    Exit Function
                End If
            Next kw
        End If
    Next c
    
    FindColumnByHeader = 0
End Function

Sub GenerateRetainAvailability()
    
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim wbOut As Workbook
    Dim lastRow As Long, lastCol As Long
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dateStartCol As Long
    Dim srcSheetName As String
    
    ' --- Column references (auto-detected) ---
    Dim colStatus As Long
    Dim colName As Long
    Dim colRank As Long
    
    ' --- Validate source workbook ---
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is open." & vbCrLf & vbCrLf & _
            "Please open the Retain file and try again.", _
            vbExclamation, "Retain Availability Generator"
        Exit Sub
    End If
    
    ' --- Use whatever sheet the user is currently on ---
    Set wsSrc = ActiveSheet
    srcSheetName = wsSrc.Name
    
    ' ============================================================
    ' AUTO-DETECT: Find header row (first row with dates)
    ' ============================================================
    Dim scanRow As Long, scanCol As Long
    Dim foundHeaderRow As Long
    Dim foundDateStartCol As Long
    foundHeaderRow = 0
    foundDateStartCol = 0
    
    For scanRow = 1 To 20
        For scanCol = 1 To 50
            If Not IsEmpty(wsSrc.Cells(scanRow, scanCol).Value) Then
                If Not IsError(wsSrc.Cells(scanRow, scanCol).Value) Then
                    If IsDate(wsSrc.Cells(scanRow, scanCol).Value) Then
                        ' Confirm it's a real date (not just a number that looks like a date)
                        ' Check that the next cell is also a date (dates come in sequences)
                        If scanCol < 50 Then
                            If Not IsError(wsSrc.Cells(scanRow, scanCol + 1).Value) Then
                                If IsDate(wsSrc.Cells(scanRow, scanCol + 1).Value) Then
                                    foundHeaderRow = scanRow
                                    foundDateStartCol = scanCol
                                    GoTo FoundDates
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next scanCol
    Next scanRow
    
FoundDates:
    If foundHeaderRow = 0 Then
        MsgBox "Could not find date headers on the active sheet """ & srcSheetName & """." & vbCrLf & vbCrLf & _
            "The script looks for a row (within rows 1-20) that contains consecutive dates." & vbCrLf & _
            "Please switch to the correct tab and try again.", _
            vbExclamation, "Retain Availability Generator"
        Exit Sub
    End If
    
    headerRow = foundHeaderRow
    dateStartCol = foundDateStartCol
    dataStartRow = headerRow + 1
    
    ' ============================================================
    ' AUTO-DETECT: Find Status, Name, and Rank columns
    ' ============================================================
    ' Look in the header row itself for column labels
    ' Some sheets have labels in the same row as dates, others one row above
    
    Dim labelRow As Long
    Dim tempLastCol As Long
    tempLastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
    
    ' Try header row first
    colStatus = FindColumnByHeader(wsSrc, headerRow, tempLastCol, "Status")
    colName = FindColumnByHeader(wsSrc, headerRow, tempLastCol, "Person Name", "ResourceName", "Resource Name", "Name", "Staff Name")
    colRank = FindColumnByHeader(wsSrc, headerRow, tempLastCol, "Rank", "Current Grade", "Grade", "Current Grade Order")
    
    ' If not found, try one row above (some sheets split label row and date row)
    If headerRow > 1 Then
        If colStatus = 0 Then colStatus = FindColumnByHeader(wsSrc, headerRow - 1, tempLastCol, "Status")
        If colName = 0 Then colName = FindColumnByHeader(wsSrc, headerRow - 1, tempLastCol, "Person Name", "ResourceName", "Resource Name", "Name", "Staff Name")
        If colRank = 0 Then colRank = FindColumnByHeader(wsSrc, headerRow - 1, tempLastCol, "Rank", "Current Grade", "Grade", "Current Grade Order")
    End If
    
    ' Validation
    If colName = 0 Then
        MsgBox "Could not find the staff name column on """ & srcSheetName & """." & vbCrLf & vbCrLf & _
            "Expected a header like ""Person Name"" or ""ResourceName"" in row " & headerRow & " or " & headerRow - 1 & "." & vbCrLf & _
            "Please check the sheet structure.", _
            vbExclamation, "Retain Availability Generator"
        Exit Sub
    End If
    
    ' Default fallbacks if Status or Rank not found
    If colStatus = 0 Then
        ' Assume column A
        colStatus = 1
    End If
    If colRank = 0 Then
        ' Assume column after name
        colRank = colName + 1
    End If
    
    ' --- Now safe to change app settings ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Determine data extent ---
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, colName).End(xlUp).Row
    lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
    
    ' --- Prompt for date range ---
    Dim startWeek As Date, endWeek As Date
    Dim inputStart As String, inputEnd As String
    
    Dim firstDate As Date, lastDate As Date
    
    ' Safe read of first/last date headers
    If IsError(wsSrc.Cells(headerRow, dateStartCol).Value) Then
        MsgBox "The first date header cell contains an error." & vbCrLf & _
            "Please fix the source data and try again.", vbExclamation, "Retain Availability Generator"
        GoTo Cleanup
    End If
    firstDate = wsSrc.Cells(headerRow, dateStartCol).Value
    
    ' Find the actual last date column
    Dim lastDateCol As Long
    lastDateCol = dateStartCol
    Dim sc As Long
    For sc = dateStartCol To lastCol
        If Not IsError(wsSrc.Cells(headerRow, sc).Value) Then
            If IsDate(wsSrc.Cells(headerRow, sc).Value) Then
                lastDateCol = sc
            End If
        End If
    Next sc
    
    If IsError(wsSrc.Cells(headerRow, lastDateCol).Value) Then
        MsgBox "The last date header cell contains an error." & vbCrLf & _
            "Please fix the source data and try again.", vbExclamation, "Retain Availability Generator"
        GoTo Cleanup
    End If
    lastDate = wsSrc.Cells(headerRow, lastDateCol).Value
    
    inputStart = InputBox("Enter START week date (DD/MM/YYYY):" & vbCrLf & vbCrLf & _
        "Source tab: " & srcSheetName & vbCrLf & _
        "Detected layout: dates in row " & headerRow & ", names in col " & Chr(64 + colName) & vbCrLf & _
        "Available range: " & Format(firstDate, "DD/MM/YYYY") & " to " & Format(lastDate, "DD/MM/YYYY"), _
        "Retain Availability Generator", Format(Date, "DD/MM/YYYY"))
    
    If inputStart = "" Then
        MsgBox "Cancelled.", vbInformation
        GoTo Cleanup
    End If
    
    inputEnd = InputBox("Enter END week date (DD/MM/YYYY):", _
        "Retain Availability Generator", Format(DateAdd("ww", 12, ParseDMY(inputStart)), "DD/MM/YYYY"))
    
    If inputEnd = "" Then
        MsgBox "Cancelled.", vbInformation
        GoTo Cleanup
    End If
    
    ' Locale-safe date parsing
    startWeek = ParseDMY(inputStart)
    endWeek = ParseDMY(inputEnd)
    
    ' --- Create new workbook ---
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Sheets(1)
    wsOut.Name = "Retain Availability"
    
    ' --- Collect display date columns ---
    Dim dateCols() As Long
    Dim dateVals() As Date
    Dim numDates As Long
    numDates = 0
    
    Dim c As Long
    For c = dateStartCol To lastCol
        If Not IsError(wsSrc.Cells(headerRow, c).Value) Then
            If IsDate(wsSrc.Cells(headerRow, c).Value) Then
                Dim dt As Date
                dt = wsSrc.Cells(headerRow, c).Value
                If dt >= startWeek And dt <= endWeek Then
                    numDates = numDates + 1
                    ReDim Preserve dateCols(1 To numDates)
                    ReDim Preserve dateVals(1 To numDates)
                    dateCols(numDates) = c
                    dateVals(numDates) = dt
                End If
            End If
        End If
    Next c
    
    If numDates = 0 Then
        MsgBox "No dates found in the specified range!", vbExclamation
        wbOut.Close False
        GoTo Cleanup
    End If
    
    ' --- Title (includes source tab name) ---
    wsOut.Range("A1").Value = "Retain Availability (" & srcSheetName & ")"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    
    ' --- PH annotations (row 2) ---
    ' Check the row above the header row for PH info
    Dim phRow As Long
    phRow = headerRow - 1
    If phRow < 1 Then phRow = 1
    
    Dim i As Long
    For i = 1 To numDates
        Dim phVal As Variant
        If Not IsError(wsSrc.Cells(phRow, dateCols(i)).Value) Then
            phVal = wsSrc.Cells(phRow, dateCols(i)).Value
            If Len(Trim(CStr(phVal & ""))) > 0 And Not IsNumeric(phVal) Then
                wsOut.Cells(2, 3 + i).Value = "PH: " & Format(dateVals(i), "D MMM")
                wsOut.Cells(2, 3 + i).Font.Size = 8
                wsOut.Cells(2, 3 + i).Font.Bold = True
            End If
        End If
    Next i
    
    ' --- Month headers (row 3) ---
    Dim prevMonth As String
    prevMonth = ""
    For i = 1 To numDates
        Dim monthName As String
        monthName = Format(dateVals(i), "MMM")
        If monthName <> prevMonth Then
            wsOut.Cells(3, 3 + i).Value = monthName
            wsOut.Cells(3, 3 + i).Font.Bold = True
            wsOut.Cells(3, 3 + i).Font.Size = 11
            prevMonth = monthName
        End If
    Next i
    
    ' --- Column headers (row 4) ---
    wsOut.Cells(4, 1).Value = "No."
    wsOut.Cells(4, 2).Value = "Staff Name"
    wsOut.Cells(4, 3).Value = "Rank"
    
    For i = 1 To numDates
        wsOut.Cells(4, 3 + i).Value = Day(dateVals(i))
        wsOut.Cells(4, 3 + i).HorizontalAlignment = xlCenter
    Next i
    
    ' Format header row
    Dim headerRange As Range
    Set headerRange = wsOut.Range(wsOut.Cells(4, 1), wsOut.Cells(4, 3 + numDates))
    With headerRange
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(255, 255, 0)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
    
    ' Column widths
    wsOut.Columns(1).ColumnWidth = 5
    wsOut.Columns(2).ColumnWidth = 40
    wsOut.Columns(3).ColumnWidth = 7
    For i = 1 To numDates
        wsOut.Columns(3 + i).ColumnWidth = 9
    Next i
    
    ' --- Write staff data ---
    Dim outRow As Long
    Dim staffNum As Long
    outRow = 5
    staffNum = 0
    
    Dim R As Long
    For R = dataStartRow To lastRow
        Dim staffName As String
        If IsError(wsSrc.Cells(R, colName).Value) Then GoTo NextStaff
        staffName = Trim(Replace(CStr(wsSrc.Cells(R, colName).Value & ""), vbLf, " "))
        If Len(staffName) = 0 Then GoTo NextStaff
        
        Dim status As String
        If IsError(wsSrc.Cells(R, colStatus).Value) Then
            status = ""
        Else
            status = LCase(Trim(Replace(CStr(wsSrc.Cells(R, colStatus).Value & ""), vbLf, " ")))
        End If
        
        Dim rank As String
        If IsError(wsSrc.Cells(R, colRank).Value) Then
            rank = ""
        Else
            rank = Trim(CStr(wsSrc.Cells(R, colRank).Value & ""))
        End If
        
        ' Abbreviate rank
        Dim rankAbbr As String
        Select Case rank
            Case "Senior-Grade 3", "Senior 3": rankAbbr = "S3"
            Case "Senior-Grade 2", "Senior 2": rankAbbr = "S2"
            Case "Senior-Grade 1", "Senior 1", "Senior 1 ": rankAbbr = "S1"
            Case "Assistant-Grade 2", "Assistant 2": rankAbbr = "A2"
            Case "Assistant-Grade 1", "Assistant 1": rankAbbr = "A1"
            Case "Assistant": rankAbbr = "A1"
            Case "Staff/Assistant 1": rankAbbr = "S/A1"
            Case "Staff/Assistant 2": rankAbbr = "S/A2"
            Case "Risk Consulting": rankAbbr = "RC"
            Case Else
                If InStr(LCase(rank), "intern") > 0 Then
                    rankAbbr = "Intern"
                ElseIf InStr(LCase(rank), "senior") > 0 And InStr(LCase(rank), "1") > 0 Then
                    rankAbbr = "S1"
                ElseIf InStr(LCase(rank), "senior") > 0 And InStr(LCase(rank), "2") > 0 Then
                    rankAbbr = "S2"
                ElseIf InStr(LCase(rank), "senior") > 0 And InStr(LCase(rank), "3") > 0 Then
                    rankAbbr = "S3"
                ElseIf InStr(LCase(rank), "assistant") > 0 And InStr(LCase(rank), "1") > 0 Then
                    rankAbbr = "A1"
                ElseIf InStr(LCase(rank), "assistant") > 0 And InStr(LCase(rank), "2") > 0 Then
                    rankAbbr = "A2"
                ElseIf InStr(LCase(rank), "staff") > 0 And InStr(LCase(rank), "1") > 0 Then
                    rankAbbr = "S/A1"
                ElseIf InStr(LCase(rank), "staff") > 0 And InStr(LCase(rank), "2") > 0 Then
                    rankAbbr = "S/A2"
                Else
                    rankAbbr = rank
                End If
        End Select
        
        staffNum = staffNum + 1
        
        ' --- Clean staff name ---
        Dim displayName As String
        displayName = CleanStaffName(staffName, startWeek, endWeek)
        
        wsOut.Cells(outRow, 1).Value = staffNum
        wsOut.Cells(outRow, 1).HorizontalAlignment = xlCenter
        wsOut.Cells(outRow, 2).Value = displayName
        wsOut.Cells(outRow, 3).Value = rankAbbr
        wsOut.Cells(outRow, 3).HorizontalAlignment = xlCenter
        
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 3)).Borders.LineStyle = xlContinuous
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 3)).Borders.Weight = xlThin
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 3)).Font.Size = 10
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 3)).Interior.Color = RGB(255, 255, 255)
        
        For i = 1 To numDates
            Dim outCol As Long
            outCol = 3 + i
            Dim oCell As Range
            Set oCell = wsOut.Cells(outRow, outCol)
            
            oCell.Borders.LineStyle = xlContinuous
            oCell.Borders.Weight = xlThin
            oCell.HorizontalAlignment = xlCenter
            oCell.Font.Size = 10
            oCell.Interior.Color = RGB(255, 255, 255)
            
            ' Check if inactive and not yet started
            If InStr(status, "inactive") > 0 Then
                Dim wefPos As Long
                wefPos = InStr(status, "wef")
                If wefPos > 0 Then
                    Dim wefStr As String
                    wefStr = Mid(status, wefPos + 4)
                    On Error Resume Next
                    Dim wefDate As Date
                    Dim dayNum As String, monStr As String
                    dayNum = ""
                    monStr = ""
                    Dim pos As Long
                    pos = 1
                    Do While pos <= Len(wefStr) And IsNumeric(Mid(wefStr, pos, 1))
                        dayNum = dayNum & Mid(wefStr, pos, 1)
                        pos = pos + 1
                    Loop
                    monStr = Trim(Mid(wefStr, pos))
                    If Len(monStr) >= 3 Then monStr = Left(monStr, 3)
                    
                    If Len(dayNum) > 0 And Len(monStr) > 0 Then
                        wefDate = DateSerial(Year(Date), Month(DateValue("1 " & monStr & " 2026")), CInt(dayNum))
                        If dateVals(i) < wefDate Then
                            oCell.Interior.Color = RGB(0, 0, 0)
                            GoTo NextDate
                        End If
                    End If
                    On Error GoTo 0
                Else
                    oCell.Interior.Color = RGB(0, 0, 0)
                    GoTo NextDate
                End If
            End If
            
            ' Check if staff has left
            Dim nameLower As String
            nameLower = LCase(staffName)
            If InStr(nameLower, "last day") > 0 Then
                Dim ldPos As Long
                ldPos = InStr(nameLower, "last day")
                Dim ldStr As String
                ldStr = Mid(nameLower, ldPos + 9)
                ldStr = Replace(ldStr, ":", "")
                ldStr = Replace(ldStr, ")", "")
                ldStr = Trim(ldStr)
                On Error Resume Next
                Dim ldDay As String, ldMon As String
                ldDay = ""
                ldMon = ""
                Dim p2 As Long
                p2 = 1
                Do While p2 <= Len(ldStr) And IsNumeric(Mid(ldStr, p2, 1))
                    ldDay = ldDay & Mid(ldStr, p2, 1)
                    p2 = p2 + 1
                Loop
                ldMon = Trim(Mid(ldStr, p2))
                If Len(ldMon) >= 3 Then ldMon = Left(ldMon, 3)
                If Len(ldDay) > 0 And Len(ldMon) > 0 Then
                    Dim lastDayDate As Date
                    lastDayDate = DateSerial(Year(Date), Month(DateValue("1 " & ldMon & " 2026")), CInt(ldDay))
                    If dateVals(i) > lastDayDate Then
                        oCell.Interior.Color = RGB(0, 0, 0)
                        GoTo NextDate
                    End If
                End If
                On Error GoTo 0
            End If
            
            ' Safe read of source cell
            Dim cellVal As Variant
            If IsError(wsSrc.Cells(R, dateCols(i)).Value) Then
                GoTo NextDate
            End If
            cellVal = wsSrc.Cells(R, dateCols(i)).Value
            
            Dim totalPct As Long
            totalPct = ExtractTotalPercent(CStr(cellVal & ""))
            
            Dim avail As Long
            avail = 100 - totalPct
            If avail < 0 Then avail = 0
            
            If avail > 0 Then
                oCell.Value = avail & "%"
                oCell.Interior.Color = RGB(255, 255, 0)
            End If
            
NextDate:
        Next i
        
        outRow = outRow + 1
NextStaff:
    Next R
    
    ' Freeze panes
    wsOut.Range("D5").Select
    ActiveWindow.FreezePanes = True
    
    ' Clean display
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.View = xlNormalView
    
    ' White background for rows 1-3
    Dim rngTop As Range
    Set rngTop = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(3, 3 + numDates))
    Dim cel As Range
    For Each cel In rngTop
        If cel.Interior.ColorIndex = xlNone Then
            cel.Interior.Color = RGB(255, 255, 255)
        End If
    Next cel
    
    ' Page setup
    wsOut.PageSetup.FitToPagesWide = 1
    wsOut.PageSetup.FitToPagesTall = 999
    
    MsgBox "Retain Availability generated in new workbook!" & vbCrLf & _
        "Source tab: " & srcSheetName & vbCrLf & _
        "Layout: row " & headerRow & ", names in col " & Chr(64 + colName) & ", dates from col " & Chr(64 + dateStartCol) & vbCrLf & _
        staffNum & " staff processed" & vbCrLf & _
        numDates & " weeks displayed", vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Function CleanStaffName(rawName As String, rangeStart As Date, rangeEnd As Date) As String
    Dim result As String
    Dim tmpName As String
    tmpName = rawName
    
    Dim firstParen As Long
    firstParen = InStr(tmpName, "(")
    
    If firstParen = 0 Then
        CleanStaffName = Trim(tmpName)
        Exit Function
    End If
    
    result = Trim(Left(tmpName, firstParen - 1))
    
    Dim remaining As String
    remaining = Mid(tmpName, firstParen)
    
    Dim pStart As Long, pEnd As Long
    Dim block As String, lowerBlock As String
    
    Do
        pStart = InStr(remaining, "(")
        If pStart = 0 Then Exit Do
        
        pEnd = InStr(pStart, remaining, ")")
        If pEnd = 0 Then
            block = Mid(remaining, pStart)
            remaining = ""
        Else
            block = Mid(remaining, pStart, pEnd - pStart + 1)
            remaining = Mid(remaining, pEnd + 1)
        End If
        
        lowerBlock = LCase(block)
        
        If InStr(lowerBlock, "last day") > 0 Then
            result = result & " " & block
            
        ElseIf InStr(lowerBlock, "ey scale") > 0 Then
            result = result & " " & block
            
        ElseIf InStr(lowerBlock, " to ") > 0 And InStr(lowerBlock, "join date") = 0 Then
            result = result & " " & block
            
        ElseIf InStr(lowerBlock, "join date") > 0 Then
            Dim jdText As String
            Dim jdPos As Long
            jdPos = InStr(lowerBlock, "join date")
            jdText = Mid(lowerBlock, jdPos + 10)
            jdText = Replace(jdText, ":", "")
            jdText = Replace(jdText, ")", "")
            jdText = Trim(jdText)
            
            On Error Resume Next
            Dim joinDate As Date
            joinDate = CDate(jdText)
            If Err.Number = 0 Then
                If joinDate >= DateAdd("m", -1, rangeStart) Then
                    result = result & " " & block
                End If
            End If
            On Error GoTo 0
            
        Else
            result = result & " " & block
        End If
    Loop
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanStaffName = Trim(result)
End Function

Function ExtractTotalPercent(cellText As String) As Long
    Dim total As Long
    total = 0
    
    If Len(Trim(cellText)) = 0 Then
        ExtractTotalPercent = 0
        Exit Function
    End If
    
    If IsNumeric(cellText) Then
        Dim numVal As Double
        numVal = CDbl(cellText)
        If numVal = 0 Then
            ExtractTotalPercent = 0
        ElseIf numVal > 0 And numVal <= 1 Then
            ExtractTotalPercent = CLng(numVal * 100)
        Else
            ExtractTotalPercent = CLng(numVal)
        End If
        Exit Function
    End If
    
    Dim i As Long
    Dim numStr As String
    numStr = ""
    
    For i = 1 To Len(cellText)
        Dim ch As String
        ch = Mid(cellText, i, 1)
        
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        ElseIf ch = "%" And Len(numStr) > 0 Then
            total = total + CLng(numStr)
            numStr = ""
        Else
            numStr = ""
        End If
    Next i
    
    ExtractTotalPercent = total
End Function