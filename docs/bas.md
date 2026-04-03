# VBA modules

_This file is generated automatically from `.bas` files in `src/bas`._

## Module `EYProject`

### `ExtractJobsByManager`

```vbnet
Sub ExtractJobsByManager()
    
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim wbOut As Workbook
    Dim lastRow As Long, lastCol As Long
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dateStartCol As Long
    
    ' --- Validate source workbook/sheet ---
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is open." & vbCrLf & vbCrLf & _
            "Please open the Retain file (containing the ""ER and P&C"" sheet) and try again.", _
            vbExclamation, "Extract Jobs by Manager"
        Exit Sub
    End If
    
    Dim sheetFound As Boolean
    sheetFound = False
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = "ER and P&C" Then
            sheetFound = True
            Exit For
        End If
    Next ws
    
    If Not sheetFound Then
        MsgBox "Could not find the ""ER and P&C"" sheet in the active workbook." & vbCrLf & vbCrLf & _
            "Please open the Retain file and make sure it is the active workbook, then try again.", _
            vbExclamation, "Extract Jobs by Manager"
        Exit Sub
    End If
    
    ' --- Prompt for manager name ---
    Dim mgrName As String
    mgrName = InputBox("Enter manager name to search for:" & vbCrLf & vbCrLf & _
        "This will find all job assignments where this name appears" & vbCrLf & _
        "in parentheses, e.g. (Ryan), (Nida/Ryan), (JS/Ryan)", _
        "Extract Jobs by Manager", "Ryan")
    
    If mgrName = "" Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Configuration ---
    Set wsSrc = ActiveWorkbook.Sheets("ER and P&C")
    headerRow = 7
    dataStartRow = 8
    dateStartCol = 4
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).Row
    lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
    
    ' --- Get ALL date headers ---
    Dim allDates() As Date
    Dim allDateCols() As Long
    Dim numAllDates As Long
    numAllDates = 0
    
    Dim c As Long
    For c = dateStartCol To lastCol
        If IsDate(wsSrc.Cells(headerRow, c).Value) Then
            numAllDates = numAllDates + 1
            ReDim Preserve allDates(1 To numAllDates)
            ReDim Preserve allDateCols(1 To numAllDates)
            allDates(numAllDates) = wsSrc.Cells(headerRow, c).Value
            allDateCols(numAllDates) = c
        End If
    Next c
    
    ' --- Scan and collect results ---
    Dim resJob() As String
    Dim resStaff() As String
    Dim resRank() As String
    Dim resPct() As Long
    Dim resDateIdx() As Long
    Dim numRes As Long
    numRes = 0
    
    Dim searchLower As String
    searchLower = LCase(mgrName)
    
    Dim r As Long, i As Long
    For r = dataStartRow To lastRow
        Dim staffRaw As String
        staffRaw = CStr(wsSrc.Cells(r, 2).Value & "")
        If Len(Trim(staffRaw)) = 0 Then GoTo NextRow
        
        Dim staffName As String
        Dim parenPos As Long
        staffName = Replace(staffRaw, vbLf, " ")
        parenPos = InStr(staffName, "(")
        If parenPos > 0 Then staffName = Left(staffName, parenPos - 1)
        staffName = Trim(staffName)
        Do While InStr(staffName, "  ") > 0
            staffName = Replace(staffName, "  ", " ")
        Loop
        
        Dim rank As String
        rank = Trim(CStr(wsSrc.Cells(r, 3).Value & ""))
        Dim rankAbbr As String
        Select Case rank
            Case "Senior-Grade 3": rankAbbr = "S3"
            Case "Senior-Grade 2": rankAbbr = "S2"
            Case "Senior-Grade 1": rankAbbr = "S1"
            Case "Assistant-Grade 2": rankAbbr = "A2"
            Case "Assistant-Grade 1": rankAbbr = "A1"
            Case "Assistant": rankAbbr = "A1"
            Case "Staff/Assistant 1": rankAbbr = "S/A1"
            Case Else
                If InStr(LCase(rank), "intern") > 0 Then
                    rankAbbr = "Intern"
                Else
                    rankAbbr = rank
                End If
        End Select
        
        For i = 1 To numAllDates
            Dim cellVal As String
            cellVal = CStr(wsSrc.Cells(r, allDateCols(i)).Value & "")
            If InStr(LCase(cellVal), searchLower) = 0 Then GoTo NextDate2
            
            ' ============================================================
            ' CRITICAL: Rejoin split lines before parsing
            ' Some cells have the manager on a separate line, e.g.:
            '   "50% - WSG \n(Nida/Ryan)"
            ' We must merge "(..." lines back onto the previous "%" line
            ' ============================================================
            Dim rawLines() As String
            rawLines = Split(Replace(cellVal, vbLf, Chr(10)), Chr(10))
            
            ' Build merged lines array
            Dim merged() As String
            Dim numMerged As Long
            numMerged = 0
            
            Dim m As Long
            For m = LBound(rawLines) To UBound(rawLines)
                Dim stripped As String
                stripped = Trim(rawLines(m))
                If Len(stripped) = 0 Then GoTo NextRawLine
                
                If Left(stripped, 1) = "(" And numMerged > 0 Then
                    ' This line starts with "(" - merge onto previous line if it had "%"
                    If InStr(merged(numMerged), "%") > 0 Then
                        merged(numMerged) = Trim(merged(numMerged)) & " " & stripped
                    Else
                        numMerged = numMerged + 1
                        ReDim Preserve merged(1 To numMerged)
                        merged(numMerged) = stripped
                    End If
                Else
                    numMerged = numMerged + 1
                    ReDim Preserve merged(1 To numMerged)
                    merged(numMerged) = stripped
                End If
NextRawLine:
            Next m
            
            ' Now process merged lines
            Dim j As Long
            For j = 1 To numMerged
                Dim ln As String
                ln = Trim(merged(j))
                If Len(ln) = 0 Then GoTo NextLine
                If InStr(LCase(ln), searchLower) = 0 Then GoTo NextLine
                
                ' --- Extract job name and TOTAL percentage from the line ---
                ' Sum all percentages in the line
                Dim totalPct As Long
                totalPct = 0
                Dim numStr As String
                numStr = ""
                Dim k As Long
                For k = 1 To Len(ln)
                    Dim ch As String
                    ch = Mid(ln, k, 1)
                    If ch >= "0" And ch <= "9" Then
                        numStr = numStr & ch
                    ElseIf ch = "%" And Len(numStr) > 0 Then
                        totalPct = totalPct + CLng(numStr)
                        numStr = ""
                    Else
                        numStr = ""
                    End If
                Next k
                
                If totalPct = 0 Then GoTo NextLine
                
                Dim pctVal As Long
                pctVal = totalPct
                
                ' Find the LAST "%" position to get the actual job name
                Dim lastPctPos As Long
                lastPctPos = 0
                For k = Len(ln) To 1 Step -1
                    If Mid(ln, k, 1) = "%" Then
                        lastPctPos = k
                        Exit For
                    End If
                Next k
                
                ' Job name is between last "% - " and "("
                Dim afterPct As String
                afterPct = Mid(ln, lastPctPos + 1)
                afterPct = Trim(afterPct)
                If Left(afterPct, 1) = "-" Then afterPct = Trim(Mid(afterPct, 2))
                
                Dim mgrParen As Long
                mgrParen = InStr(afterPct, "(")
                Dim jobName As String
                If mgrParen > 0 Then
                    jobName = Trim(Left(afterPct, mgrParen - 1))
                Else
                    jobName = Trim(afterPct)
                End If
                
                ' Clean trailing dashes/spaces
                Do While Len(jobName) > 0 And (Right(jobName, 1) = "-" Or Right(jobName, 1) = " ")
                    jobName = Left(jobName, Len(jobName) - 1)
                Loop
                
                If Len(jobName) = 0 Then GoTo NextLine
                
                ' Store result
                numRes = numRes + 1
                ReDim Preserve resJob(1 To numRes)
                ReDim Preserve resStaff(1 To numRes)
                ReDim Preserve resRank(1 To numRes)
                ReDim Preserve resPct(1 To numRes)
                ReDim Preserve resDateIdx(1 To numRes)
                resJob(numRes) = jobName
                resStaff(numRes) = staffName
                resRank(numRes) = rankAbbr
                resPct(numRes) = pctVal
                resDateIdx(numRes) = i
                
NextLine:
            Next j
NextDate2:
        Next i
NextRow:
    Next r
    
    If numRes = 0 Then
        MsgBox "No jobs found for """ & mgrName & """.", vbInformation
        GoTo Cleanup
    End If
    
    ' --- Build unique job+staff rows (sorted by job then staff) ---
    Dim rowKeys() As String
    Dim rowJob() As String
    Dim rowStaff() As String
    Dim rowRank() As String
    Dim numRows As Long
    numRows = 0
    
    For i = 1 To numRes
        Dim rKey As String
        rKey = LCase(resJob(i)) & "|" & LCase(resStaff(i))
        Dim found As Boolean
        found = False
        Dim s As Long
        For s = 1 To numRows
            If rowKeys(s) = rKey Then
                found = True
                Exit For
            End If
        Next s
        If Not found Then
            numRows = numRows + 1
            ReDim Preserve rowKeys(1 To numRows)
            ReDim Preserve rowJob(1 To numRows)
            ReDim Preserve rowStaff(1 To numRows)
            ReDim Preserve rowRank(1 To numRows)
            rowKeys(numRows) = rKey
            rowJob(numRows) = resJob(i)
            rowStaff(numRows) = resStaff(i)
            rowRank(numRows) = resRank(i)
        End If
    Next i
    
    ' Sort rows by job name then staff
    Dim swapped As Boolean
    Dim tmp As String
    Do
        swapped = False
        For s = 1 To numRows - 1
            If rowKeys(s) > rowKeys(s + 1) Then
                tmp = rowKeys(s): rowKeys(s) = rowKeys(s + 1): rowKeys(s + 1) = tmp
                tmp = rowJob(s): rowJob(s) = rowJob(s + 1): rowJob(s + 1) = tmp
                tmp = rowStaff(s): rowStaff(s) = rowStaff(s + 1): rowStaff(s + 1) = tmp
                tmp = rowRank(s): rowRank(s) = rowRank(s + 1): rowRank(s + 1) = tmp
                swapped = True
            End If
        Next s
    Loop While swapped
    
    ' --- Determine date range to display ---
    Dim dateHasData() As Boolean
    ReDim dateHasData(1 To numAllDates)
    For i = 1 To numRes
        dateHasData(resDateIdx(i)) = True
    Next i
    
    Dim firstIdx As Long, lastIdx As Long
    firstIdx = 0
    lastIdx = 0
    For i = 1 To numAllDates
        If dateHasData(i) Then
            If firstIdx = 0 Then firstIdx = i
            lastIdx = i
        End If
    Next i
    
    If firstIdx > 1 Then firstIdx = firstIdx - 1
    If lastIdx < numAllDates Then lastIdx = lastIdx + 1
    
    Dim dispDates() As Date
    Dim dispDateIdx() As Long
    Dim numDispDates As Long
    numDispDates = 0
    For i = firstIdx To lastIdx
        numDispDates = numDispDates + 1
        ReDim Preserve dispDates(1 To numDispDates)
        ReDim Preserve dispDateIdx(1 To numDispDates)
        dispDates(numDispDates) = allDates(i)
        dispDateIdx(numDispDates) = i
    Next i
    
    ' --- Build pct lookup ---
    Dim pctLookupKey() As String
    Dim pctLookupVal() As Long
    Dim numLookup As Long
    numLookup = 0
    
    For i = 1 To numRes
        rKey = LCase(resJob(i)) & "|" & LCase(resStaff(i))
        Dim rowIdx As Long
        rowIdx = 0
        For s = 1 To numRows
            If rowKeys(s) = rKey Then
                rowIdx = s
                Exit For
            End If
        Next s
        If rowIdx = 0 Then GoTo NextRes
        
        Dim lookupKey As String
        lookupKey = CStr(rowIdx) & "|" & CStr(resDateIdx(i))
        
        found = False
        For s = 1 To numLookup
            If pctLookupKey(s) = lookupKey Then
                pctLookupVal(s) = pctLookupVal(s) + resPct(i)
                found = True
                Exit For
            End If
        Next s
        If Not found Then
            numLookup = numLookup + 1
            ReDim Preserve pctLookupKey(1 To numLookup)
            ReDim Preserve pctLookupVal(1 To numLookup)
            pctLookupKey(numLookup) = lookupKey
            pctLookupVal(numLookup) = resPct(i)
        End If
NextRes:
    Next i
    
    ' --- Create output workbook ---
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Sheets(1)
    wsOut.Name = "Jobs by " & mgrName
    
    ' --- Title ---
    wsOut.Range("A1").Value = "Project Timeline - " & mgrName
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    
    ' --- Month headers (row 2) ---
    Dim prevMonth As String
    prevMonth = ""
    Dim dateColStart As Long
    dateColStart = 5
    
    For i = 1 To numDispDates
        Dim mName As String
        mName = Format(dispDates(i), "MMM YYYY")
        If mName <> prevMonth Then
            wsOut.Cells(2, dateColStart + i - 1).Value = Format(dispDates(i), "MMM")
            wsOut.Cells(2, dateColStart + i - 1).Font.Bold = True
            wsOut.Cells(2, dateColStart + i - 1).Font.Size = 10
            prevMonth = mName
        End If
    Next i
    
    ' --- Column headers (row 3) ---
    wsOut.Cells(3, 1).Value = "No."
    wsOut.Cells(3, 2).Value = "Job Name"
    wsOut.Cells(3, 3).Value = "Staff"
    wsOut.Cells(3, 4).Value = "Rank"
    
    For i = 1 To numDispDates
        wsOut.Cells(3, dateColStart + i - 1).Value = Day(dispDates(i))
        wsOut.Cells(3, dateColStart + i - 1).HorizontalAlignment = xlCenter
    Next i
    
    ' Format header row
    Dim headerRange As Range
    Set headerRange = wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(3, dateColStart + numDispDates - 1))
    With headerRange
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
    
    ' Column widths
    wsOut.Columns(1).ColumnWidth = 4
    wsOut.Columns(2).ColumnWidth = 30
    wsOut.Columns(3).ColumnWidth = 25
    wsOut.Columns(4).ColumnWidth = 6
    For i = 1 To numDispDates
        wsOut.Columns(dateColStart + i - 1).ColumnWidth = 5
    Next i
    
    ' --- Color palette for jobs ---
    Dim colors(0 To 9) As Long
    colors(0) = RGB(79, 129, 189)
    colors(1) = RGB(192, 80, 77)
    colors(2) = RGB(155, 187, 89)
    colors(3) = RGB(128, 100, 162)
    colors(4) = RGB(75, 172, 198)
    colors(5) = RGB(247, 150, 70)
    colors(6) = RGB(119, 147, 60)
    colors(7) = RGB(180, 120, 80)
    colors(8) = RGB(150, 60, 90)
    colors(9) = RGB(60, 140, 140)
    
    ' --- Write data rows ---
    Dim outRow As Long
    outRow = 4
    Dim prevJobName As String
    prevJobName = ""
    Dim jobColor As Long
    Dim jobColorIdx As Long
    jobColorIdx = -1
    
    For s = 1 To numRows
        ' New job group - insert separator
        If LCase(rowJob(s)) <> LCase(prevJobName) Then
            If prevJobName <> "" Then
                wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, dateColStart + numDispDates - 1)).Interior.Color = RGB(255, 255, 255)
                wsOut.Rows(outRow).RowHeight = 4
                outRow = outRow + 1
            End If
            prevJobName = rowJob(s)
            jobColorIdx = (jobColorIdx + 1) Mod 10
            jobColor = colors(jobColorIdx)
        End If
        
        wsOut.Cells(outRow, 1).Value = s
        wsOut.Cells(outRow, 1).HorizontalAlignment = xlCenter
        wsOut.Cells(outRow, 2).Value = rowJob(s)
        wsOut.Cells(outRow, 3).Value = rowStaff(s)
        wsOut.Cells(outRow, 4).Value = rowRank(s)
        wsOut.Cells(outRow, 4).HorizontalAlignment = xlCenter
        
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 4)).Font.Size = 9
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 4)).Borders.LineStyle = xlContinuous
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 4)).Borders.Weight = xlThin
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 4)).Interior.Color = RGB(255, 255, 255)
        
        ' Write date cells
        For i = 1 To numDispDates
            Dim oCol As Long
            oCol = dateColStart + i - 1
            Dim oCell As Range
            Set oCell = wsOut.Cells(outRow, oCol)
            
            oCell.Borders.LineStyle = xlContinuous
            oCell.Borders.Weight = xlHairline
            oCell.Borders.Color = RGB(200, 200, 200)
            oCell.HorizontalAlignment = xlCenter
            oCell.Font.Size = 8
            oCell.Interior.Color = RGB(245, 245, 245)
            
            ' Look up percentage
            lookupKey = CStr(s) & "|" & CStr(dispDateIdx(i))
            Dim pctFound As Long
            pctFound = 0
            Dim lk As Long
            For lk = 1 To numLookup
                If pctLookupKey(lk) = lookupKey Then
                    pctFound = pctLookupVal(lk)
                    Exit For
                End If
            Next lk
            
            If pctFound > 0 Then
                oCell.Value = pctFound & "%"
                oCell.Interior.Color = jobColor
                oCell.Font.Color = RGB(255, 255, 255)
                oCell.Font.Bold = True
            End If
        Next i
        
        outRow = outRow + 1
    Next s
    
    ' --- Freeze panes ---
    wsOut.Cells(4, dateColStart).Select
    ActiveWindow.FreezePanes = True
    
    ' Clean display
    ActiveWindow.DisplayGridlines = False
    
    ' White background for title rows
    Dim cel As Range
    For Each cel In wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(2, dateColStart + numDispDates - 1))
        If cel.Interior.ColorIndex = xlNone Then
            cel.Interior.Color = RGB(255, 255, 255)
        End If
    Next cel
    
    ' Page setup
    wsOut.PageSetup.Orientation = 2
    wsOut.PageSetup.FitToPagesWide = 1
    wsOut.PageSetup.FitToPagesTall = 999
    
    ' --- Count unique jobs ---
    Dim uniqueJobs As Long
    uniqueJobs = 0
    prevJobName = ""
    For s = 1 To numRows
        If LCase(rowJob(s)) <> LCase(prevJobName) Then
            uniqueJobs = uniqueJobs + 1
            prevJobName = rowJob(s)
        End If
    Next s
    
    MsgBox "Project timeline generated for """ & mgrName & """!" & vbCrLf & _
        uniqueJobs & " jobs found" & vbCrLf & _
        numRows & " job-staff assignments" & vbCrLf & _
        numDispDates & " weeks displayed", vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
```

## Module `EYRetain`

### `GenerateRetainAvailability`

```vbnet
Sub GenerateRetainAvailability()
    
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim wbOut As Workbook
    Dim lastRow As Long, lastCol As Long
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dateStartCol As Long
    
    ' --- Validate source workbook/sheet FIRST (before changing app settings) ---
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is open." & vbCrLf & vbCrLf & _
            "Please open the Retain file (containing the ""ER and P&C"" sheet) and try again.", _
            vbExclamation, "Retain Availability Generator"
        Exit Sub
    End If
    
    Dim sheetFound As Boolean
    sheetFound = False
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = "ER and P&C" Then
            sheetFound = True
            Exit For
        End If
    Next ws
    
    If Not sheetFound Then
        MsgBox "Could not find the ""ER and P&C"" sheet in the active workbook." & vbCrLf & vbCrLf & _
            "Please open the Retain file and make sure it is the active workbook, then try again.", _
            vbExclamation, "Retain Availability Generator"
        Exit Sub
    End If
    
    ' --- Now safe to change app settings ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Configuration ---
    Set wsSrc = ActiveWorkbook.Sheets("ER and P&C")
    headerRow = 7
    dataStartRow = 8
    dateStartCol = 4 ' Column D
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).Row
    lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
    
    ' --- Prompt for date range ---
    Dim startWeek As Date, endWeek As Date
    Dim inputStart As String, inputEnd As String
    
    Dim firstDate As Date, lastDate As Date
    firstDate = wsSrc.Cells(headerRow, dateStartCol).Value
    lastDate = wsSrc.Cells(headerRow, lastCol).Value
    
    inputStart = InputBox("Enter START week date (DD/MM/YYYY):" & vbCrLf & vbCrLf & _
        "Available range: " & Format(firstDate, "DD/MM/YYYY") & " to " & Format(lastDate, "DD/MM/YYYY"), _
        "Retain Availability Generator", Format(Date, "DD/MM/YYYY"))
    
    If inputStart = "" Then
        MsgBox "Cancelled.", vbInformation
        GoTo Cleanup
    End If
    
    inputEnd = InputBox("Enter END week date (DD/MM/YYYY):", _
        "Retain Availability Generator", Format(DateAdd("ww", 12, CDate(inputStart)), "DD/MM/YYYY"))
    
    If inputEnd = "" Then
        MsgBox "Cancelled.", vbInformation
        GoTo Cleanup
    End If
    
    startWeek = CDate(inputStart)
    endWeek = CDate(inputEnd)
    
    ' --- Create new workbook ---
    Set wbOut = Workbooks.Add(xlWBATWorksheet) ' New workbook with 1 sheet
    Set wsOut = wbOut.Sheets(1)
    wsOut.Name = "Retain Availability"
    
    ' --- Collect display date columns ---
    Dim dateCols() As Long
    Dim dateVals() As Date
    Dim numDates As Long
    numDates = 0
    
    Dim c As Long
    For c = dateStartCol To lastCol
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
    Next c
    
    If numDates = 0 Then
        MsgBox "No dates found in the specified range!", vbExclamation
        wbOut.Close False
        GoTo Cleanup
    End If
    
    ' --- Title ---
    wsOut.Range("A1").Value = "Retain Availability (ER and P&C)"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14
    
    ' --- PH annotations (row 2) ---
    Dim i As Long
    For i = 1 To numDates
        Dim phVal As Variant
        phVal = wsSrc.Cells(6, dateCols(i)).Value
        If Len(Trim(CStr(phVal & ""))) > 0 And Not IsNumeric(phVal) Then
            wsOut.Cells(2, 3 + i).Value = "PH: " & Format(dateVals(i), "D MMM")
            wsOut.Cells(2, 3 + i).Font.Size = 8
            wsOut.Cells(2, 3 + i).Font.Bold = True
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
    
    Dim r As Long
    For r = dataStartRow To lastRow
        Dim staffName As String
        staffName = Trim(Replace(CStr(wsSrc.Cells(r, 2).Value & ""), vbLf, " "))
        If Len(staffName) = 0 Then GoTo NextStaff
        
        Dim status As String
        status = LCase(Trim(Replace(CStr(wsSrc.Cells(r, 1).Value & ""), vbLf, " ")))
        
        Dim rank As String
        rank = Trim(CStr(wsSrc.Cells(r, 3).Value & ""))
        
        ' Abbreviate rank
        Dim rankAbbr As String
        Select Case rank
            Case "Senior-Grade 3": rankAbbr = "S3"
            Case "Senior-Grade 2": rankAbbr = "S2"
            Case "Senior-Grade 1": rankAbbr = "S1"
            Case "Assistant-Grade 2": rankAbbr = "A2"
            Case "Assistant-Grade 1": rankAbbr = "A1"
            Case "Assistant": rankAbbr = "A1"
            Case "Staff/Assistant 1": rankAbbr = "S/A1"
            Case Else
                If InStr(LCase(rank), "intern") > 0 Then
                    rankAbbr = "Intern"
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
                        wefDate = DateSerial(2026, Month(DateValue("1 " & monStr & " 2026")), CInt(dayNum))
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
                    lastDayDate = DateSerial(2026, Month(DateValue("1 " & ldMon & " 2026")), CInt(ldDay))
                    If dateVals(i) > lastDayDate Then
                        oCell.Interior.Color = RGB(0, 0, 0)
                        GoTo NextDate
                    End If
                End If
                On Error GoTo 0
            End If
            
            ' Calculate availability
            Dim cellVal As Variant
            cellVal = wsSrc.Cells(r, dateCols(i)).Value
            
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
    Next r
    
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
        staffNum & " staff processed" & vbCrLf & _
        numDates & " weeks displayed", vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
```

### `CleanStaffName`

```vbnet
Function CleanStaffName(rawName As String, rangeStart As Date, rangeEnd As Date) As String
    ' Parses staff name and handles multiple parenthetical blocks individually.
    ' - "(Join date: ...)" -> only keep if join date is near/within display range
    ' - "(Last day: ...)" -> always keep
    ' - "(EY SCALE ...)" -> always keep
    ' - "(4 May 26 to 31 Jul 26)" -> always keep (intern ranges)
    
    Dim result As String
    Dim tmpName As String
    tmpName = rawName
    
    ' Extract the base name (everything before first parenthesis)
    Dim firstParen As Long
    firstParen = InStr(tmpName, "(")
    
    If firstParen = 0 Then
        CleanStaffName = Trim(tmpName)
        Exit Function
    End If
    
    result = Trim(Left(tmpName, firstParen - 1))
    
    ' Process each parenthetical block
    Dim remaining As String
    remaining = Mid(tmpName, firstParen)
    
    Dim pStart As Long, pEnd As Long
    Dim block As String, lowerBlock As String
    
    Do
        pStart = InStr(remaining, "(")
        If pStart = 0 Then Exit Do
        
        pEnd = InStr(pStart, remaining, ")")
        If pEnd = 0 Then
            ' Unclosed paren - take rest
            block = Mid(remaining, pStart)
            remaining = ""
        Else
            block = Mid(remaining, pStart, pEnd - pStart + 1)
            remaining = Mid(remaining, pEnd + 1)
        End If
        
        lowerBlock = LCase(block)
        
        ' Decide whether to keep this block
        If InStr(lowerBlock, "last day") > 0 Then
            ' Always keep Last day
            result = result & " " & block
            
        ElseIf InStr(lowerBlock, "ey scale") > 0 Then
            ' Always keep EY SCALE
            result = result & " " & block
            
        ElseIf InStr(lowerBlock, " to ") > 0 And InStr(lowerBlock, "join date") = 0 Then
            ' Intern date range - always keep
            result = result & " " & block
            
        ElseIf InStr(lowerBlock, "join date") > 0 Then
            ' Parse join date - only keep if recent
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
                ' Keep only if join date is within 1 month before range start or later
                If joinDate >= DateAdd("m", -1, rangeStart) Then
                    result = result & " " & block
                End If
            End If
            On Error GoTo 0
            
        Else
            ' Unknown - keep it
            result = result & " " & block
        End If
    Loop
    
    ' Clean up double spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanStaffName = Trim(result)
End Function
```

### `ExtractTotalPercent`

```vbnet
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
```

## Module `Functions`

### `XFIRSTUSEDCOL`

```vbnet
Public Function XFIRSTUSEDCOL(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    result = Rng.Find(What:="*", _
                      After:=Rng.Cells(1), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlNext, _
                      MatchCase:=False).Column
    XFIRSTUSEDCOL = result
    If Err.Number <> 0 Then XFIRSTUSEDCOL = Rng.Column + Rng.Columns.Count - 1
End Function
```

### `XLASTUSEDCOL`

```vbnet
Public Function XLASTUSEDCOL(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    result = Rng.Find(What:="*", _
                      After:=Rng.Cells(1), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Column
    XLASTUSEDCOL = result
    If Err.Number <> 0 Then XLASTUSEDCOL = Rng.Column + Rng.Columns.Count - 1
End Function
```

### `XFIRSTUSEDROW`

```vbnet
Public Function XFIRSTUSEDROW(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    If IsEmpty(Rng.Cells(1)) Then
        result = Rng.Find(What:="*", _
                          After:=Rng.Cells(1), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False).Row
    Else
        result = Rng.Cells(1).Row
    End If
    XFIRSTUSEDROW = result
    If Err.Number <> 0 Then XFIRSTUSEDROW = 0
End Function
```

### `XLASTUSEDROW`

```vbnet
Public Function XLASTUSEDROW(Rng As Range) As Long
    Dim result As Long
    On Error Resume Next
    result = Rng.Find(What:="*", _
                      After:=Rng.Cells(1), _
                      Lookat:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Row
    XLASTUSEDROW = result
    If Err.Number <> 0 Then XLASTUSEDROW = 0
End Function
```

### `XRELEVANTAREA`

```vbnet
Public Function XRELEVANTAREA(rngTarget As Range) As Range
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    firstRow = XFIRSTUSEDROW(rngTarget)
    firstCol = XFIRSTUSEDCOL(rngTarget)
    lastRow = XLASTUSEDROW(rngTarget)
    lastCol = XLASTUSEDCOL(rngTarget)
    Set XRELEVANTAREA = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))
End Function
```

### `XCOLUMNWIDTH`

```vbnet
Public Function XCOLUMNWIDTH(target As Range) As Double
    Application.ScreenUpdating = False
    XCOLUMNWIDTH = target.ColumnWidth
    Application.ScreenUpdating = True
End Function
```

### `XGETBOLD`

```vbnet
Public Function XGETBOLD(pWorkRng As Range)
    XGETBOLD = pWorkRng.Font.Bold
End Function
```

### `XGETINDENTLEVEL`

```vbnet
Public Function XGETINDENTLEVEL(targetCell As Range)
    XGETINDENTLEVEL = targetCell.IndentLevel
End Function
```

### `XCOUNTCOLOR`

```vbnet
Public Function XCOUNTCOLOR(CountRange As Range, CountColor As Range)
    Dim CountColorValue As Long
    Dim TotalCount As Long
    Dim rCell As Range
    CountColorValue = CountColor.Interior.ColorIndex
    For Each rCell In CountRange
        If rCell.Interior.ColorIndex = CountColorValue Then
            TotalCount = TotalCount + 1
        End If
    Next rCell
    XCOUNTCOLOR = TotalCount
End Function
```

### `XEXTRACTAFTER`

```vbnet
Public Function XEXTRACTAFTER(rngWord As Range, strWord As String) As String
    On Error GoTo ErrorHandler
    Application.Volatile
    Dim lngStart As Long, lngEnd As Long, tempResult As String
    lngStart = InStr(1, rngWord, strWord)
    If lngStart = 0 Then
        XEXTRACTAFTER = "Not found": Exit Function
    End If
    lngEnd = InStr(lngStart + Len(strWord), rngWord, Len(rngWord))
    If lngEnd = 0 Then lngEnd = Len(rngWord)
    tempResult = Mid(rngWord, lngStart + Len(strWord), lngEnd - lngStart)
    XEXTRACTAFTER = Trim(tempResult)
    Exit Function
ErrorHandler:
    XEXTRACTAFTER = Err.Description
End Function
```

### `XEXTRACTBEFORE`

```vbnet
Public Function XEXTRACTBEFORE(rngWord As Range, strWord As String) As String
    On Error GoTo ErrorHandler
    Application.Volatile
    Dim lngEnd As Long, tempResult As String
    lngEnd = InStr(1, rngWord, strWord)
    If lngEnd = 0 Then
        XEXTRACTBEFORE = "Not found": Exit Function
    End If
    tempResult = Left(rngWord, lngEnd - 1)
    XEXTRACTBEFORE = Trim(tempResult)
    Exit Function
ErrorHandler:
    XEXTRACTBEFORE = Err.Description
End Function
```

### `XHYPERACTIVE`

```vbnet
Public Function XHYPERACTIVE(ByRef Rng As Range)
    Dim strAddress As String, strTextDisplay As String
    Dim target As Range
    Application.DisplayAlerts = False
    On Error Resume Next
    Set target = Application.InputBox(Title:="Create Hyperlink", _
                                      Prompt:="Select a cell to create hyperlink", _
                                      Type:=8)
    On Error GoTo 0
    Application.DisplayAlerts = True
    If Rng Is Nothing Or target Is Nothing Then Exit Function
    strAddress = Chr(39) & target.Parent.Name & Chr(39) & "!" & target.Address
    If WorksheetFunction.CountA(Rng) = 0 Then
        strTextDisplay = target.Parent.Name
    Else
        strTextDisplay = Rng.Value
    End If
    With ActiveSheet.Hyperlinks
        .Add Anchor:=Rng, Address:="", SubAddress:=strAddress, TextToDisplay:=strTextDisplay
    End With
End Function
```

## Module `Ribbon`

### `RibbonOnLoad`

```vbnet
Public Sub RibbonOnLoad(R As IRibbonUI)
    Set RibbonUI = R
End Sub
```

### `RunByName`

```vbnet
Public Sub RunByName(control As IRibbonControl)
    Dim macro As String
    macro = control.Tag
    If Len(macro) = 0 Then macro = control.ID
    On Error GoTo errh
    Application.Run macro
    Exit Sub
errh:
    MsgBox "Macro not found: " & macro, vbExclamation
End Sub
```

### `GetSheetTabColor`

```vbnet
Public Sub GetSheetTabColor(control As IRibbonControl, ByRef returnedVal)
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastShTab", "Green")
    
    Dim clr As Long
    Select Case last
        Case "Green":  clr = RGB(204, 255, 204)
        Case "Red":    clr = RGB(255, 204, 204)
        Case "Yellow": clr = RGB(255, 255, 0)
        Case Else:     clr = RGB(255, 255, 255)
    End Select
    
    Set returnedVal = CreateColorIcon(clr)
    
End Sub
```

### `GetHighlightColor`

```vbnet
Public Sub GetHighlightColor(control As IRibbonControl, ByRef returnedVal)
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastHighlight", "Green")
    
    Dim clr As Long
    Select Case last
        Case "Green":  clr = RGB(204, 255, 204)
        Case "Red":    clr = RGB(255, 204, 204)
        Case "Yellow": clr = RGB(255, 255, 0)
        Case Else:     clr = RGB(255, 255, 255)
    End Select
    
    Set returnedVal = CreateColorIcon(clr)
    
End Sub
```

### `GetWorkbookFontLabel`

```vbnet
Public Sub GetWorkbookFontLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastWbFont", "Arial")
End Sub
```

### `GetWorkbookFontSizeLabel`

```vbnet
Public Sub GetWorkbookFontSizeLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastWbFontSize", "10")
End Sub
```

### `GetSheetFontLabel`

```vbnet
Public Sub GetSheetFontLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastShFont", "Arial")
End Sub
```

### `GetSelNumberLabel`

```vbnet
Public Sub GetSelNumberLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelNumber", "Accounting")
End Sub
```

### `GetSelCaseLabel`

```vbnet
Public Sub GetSelCaseLabel(control As IRibbonControl, ByRef label)
    label = GetSetting("ExcelUI", "Preferences", "LastSelCase", "Proper")
End Sub
```

### `CreateColorIcon`

```vbnet
Private Function CreateColorIcon(clr As Long) As IPictureDisp
    Dim tmp As String
    tmp = Environ("TEMP") & "\excelui_icon.bmp"
    
    ' Create a 16x16 BMP file
    Dim f As Integer
    f = FreeFile
    Open tmp For Binary As #f
    
    ' BMP header (54 bytes) + 16x16x3 pixel data (768 bytes) = 822 bytes
    Dim bmp(0 To 821) As Byte
    
    ' BM signature
    bmp(0) = 66: bmp(1) = 77
    ' File size = 822
    bmp(2) = 54: bmp(3) = 3: bmp(4) = 0: bmp(5) = 0
    ' Data offset = 54
    bmp(10) = 54
    ' DIB header size = 40
    bmp(14) = 40
    ' Width = 16
    bmp(18) = 16
    ' Height = 16
    bmp(22) = 16
    ' Planes = 1
    bmp(26) = 1
    ' Bits per pixel = 24
    bmp(28) = 24
    
    ' Fill pixels with color (BGR order)
    Dim R As Byte, G As Byte, B As Byte
    R = clr And &HFF
    G = (clr \ &H100) And &HFF
    B = (clr \ &H10000) And &HFF
    
    Dim i As Long
    For i = 54 To 821 Step 3
        bmp(i) = B
        bmp(i + 1) = G
        bmp(i + 2) = R
    Next i
    
    Put #f, , bmp
    Close #f
    
    Set CreateColorIcon = LoadPicture(tmp)
    Kill tmp
End Function
```

## Module `Subs`

### `CaseProper`

```vbnet
Sub CaseProper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = StrConv(cell, vbProperCase)
    Next cell
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelCase", "Proper"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `CaseSentence`

```vbnet
Sub CaseSentence()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    
    Application.ScreenUpdating = False
    For Each Rng In XRELEVANTAREA(WorkRng)
        xValue = Rng.Value
        xStart = True
        For i = 1 To VBA.Len(xValue)
            ch = Mid(xValue, i, 1)
            Select Case ch
                Case "."
                xStart = True
                Case "?"
                xStart = True
                Case "a" To "z"
                If xStart Then
                    ch = UCase(ch)
                    xStart = False
                End If
                Case "A" To "Z"
                If xStart Then
                    xStart = False
                Else
                    ch = LCase(ch)
                End If
            End Select
            Mid(xValue, i, 1) = ch
        Next
        Rng.Value = xValue
    Next
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelCase", "Sentence"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `CaseUpper`

```vbnet
Sub CaseUpper()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = UCase(cell)
    Next cell
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelCase", "Upper"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `CellTrim`

```vbnet
Sub CellTrim()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    For Each cell In XRELEVANTAREA(Rng)
        cell.Value = Trim(cell)
    Next cell
    Application.ScreenUpdating = True

ErrorHandler:
    Exit Sub
    
End Sub
```

### `ETLSAM`

```vbnet
Sub ETLSAM()
'   Purpose: Format Identified Misstatements

    Dim rngA, rngB, rngTitle As Range
    Dim lastRow, lastCol As Long
    Dim strEGA As String
    
    lastRow = ActiveSheet.UsedRange.Rows.Count
    lastCol = ActiveSheet.UsedRange.Columns.Count
    
    Set rngTitle = ActiveSheet.Range("A2")
    Set rngA = ActiveSheet.UsedRange
    Set rngB = ActiveSheet.Range(Cells(8, 1), Cells(lastRow, lastCol))
    Set rngAmount = ActiveSheet.Range(Cells(9, lastCol), Cells(lastRow, lastCol))
    
    If Not rngTitle = "Summary of Corrected and Uncorrected Misstatements" Then
        Exit Sub
    End If
    
    'Remove blanks from range
    With rngB
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    'Copy down blank cells
    rngB.Select
    With rngB
        On Error GoTo skiperror
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    End With
    
skiperror:

    'Copy paste as values
    rngB.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Format amount to accounting
    rngAmount.Select
    Call FormatAccounting
    
    'Format workbook to Arial
    Call WorkbookArial
    Call WorkbookPageBreakOff
    
End Sub
```

### `ETLTB`

```vbnet
Sub ETLTB()
'   Purpose: Format Aura TB Export

    Dim rngA, rngB, rngTitle, rngAccount As Range
    Dim lastRow, lastCol As Long
    Dim strEGA As String
    
    lastRow = ActiveSheet.UsedRange.Rows.Count
    lastCol = ActiveSheet.UsedRange.Columns.Count
    
    Set rngTitle = ActiveSheet.Range("A1")
    Set rngA = ActiveSheet.UsedRange
    Set rngB = ActiveSheet.Range(Cells(2, 1), Cells(lastRow, 3))
    Set rngAccount = ActiveSheet.Range(Cells(2, 5), Cells(lastRow, 5))
    
    'Check if EGA is TB
    If Not rngTitle = "FSLI No." Then
        Exit Sub
    End If
       
    'Remove blank rows
    Call SheetRemoveBlankRows
    
    'Copy down blank cells
    With rngB
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    rngB.Select
    With rngB
        On Error GoTo skiperror
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    End With

skiperror:

    'Copy paste as values
    rngB.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Format amount to accounting
'    rngAmount.Select
'    Call FormatAccounting
    
'    Format workbook to Arial
    Call WorkbookArial
    Call WorkbookPageBreakOff
    
    rngAccount.Select
    Call RemoveBlankRows
    
End Sub
```

### `FormatAccounting`

```vbnet
Sub FormatAccounting()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range, rngB As Range
    Set rngSelection = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Accounting"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormatCellRed`

```vbnet
Sub FormatCellRed()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(122, 24, 24)
    Rng.Font.Color = RGB(255, 255, 255)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormatDateDDMMM`

```vbnet
Sub FormatDateDDMMM()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlCenter
        c.NumberFormat = "dd-mmm-yy"
    Next c
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormatDateDDMMYY`

```vbnet
Sub FormatDateDDMMYY()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlCenter
        c.NumberFormat = "dd/mm/yy"
    Next c
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormatFontBlue`

```vbnet
Sub FormatFontBlue()
    
    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Font.Color = RGB(0, 112, 192)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormatFontGreen`

```vbnet
Sub FormatFontGreen()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Font.Color = RGB(0, 176, 80)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormatFontOrange`

```vbnet
Sub FormatFontOrange()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Font.Color = RGB(237, 125, 49)
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormatHighlightGreen`

```vbnet
Sub FormatHighlightGreen()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(204, 255, 204)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Green"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate

ErrorHandler:
    Exit Sub
End Sub
```

### `FormatHighlightRed`

```vbnet
Sub FormatHighlightRed()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 204, 204)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Red"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate

ErrorHandler:
    Exit Sub
End Sub
```

### `FormatHighlightYellow`

```vbnet
Sub FormatHighlightYellow()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = RGB(255, 255, 0)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Yellow"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
End Sub
```

### `FormatHighlightReset`

```vbnet
Sub FormatHighlightReset()

    On Error GoTo ErrorHandler
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Interior.Color = xlNone
    Rng.Font.Color = RGB(0, 0, 0)
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastHighlight", "Clear"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
End Sub
```

### `FormatTableBordersGrey`

```vbnet
Sub FormatTableBordersGrey()

    On Error GoTo ErrorHandler
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormatTextToValue`

```vbnet
Sub FormatTextToValue()

    On Error GoTo ErrorHandler
    
    Dim rngSelection As Range
    Set rngSelection = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(rngSelection)
        c.WrapText = False
        c.HorizontalAlignment = xlLeft
        c.NumberFormat = "General"
        c.Value = c.Value
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `FormulaAbsolute`

```vbnet
Sub FormulaAbsolute()
'   Purpose: Convert selected values to absolute

    On Error GoTo ErrorHandler

    Dim Rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set Rng = Selection

    Dim regexTargetText As Object
    Set regexTargetText = New RegExp
    With regexTargetText
    .Pattern = "ABS"
    .Global = False
    End With

    Dim c As Range
    For Each c In Rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            If regexTargetText.Test(myFormula) Then
                myFormula = regexTargetText.Replace(myFormula, "")
                myFormula = Replace(myFormula, "(", "")
                myFormula = Replace(myFormula, ")", "")
                c.Formula = "=" & myFormula
            Else
                c.Formula = "=ABS(" & myFormula & ")"
            End If
        Else
            cellValue = c.Value
            c.Formula = "=ABS(" & cellValue & ")"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Absolute"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `FormulaReverseSign`

```vbnet
Sub FormulaReverseSign()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(Rng)
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            If Left(myFormula, 1) = "-" Then
                c.Formula = "=" & Right(myFormula, Len(myFormula) - 1)
            Else
                c.Formula = "=-" & myFormula
            End If
        Else
                c.Formula = "=-" & c.Value
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Reverse Sign"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormulaRound`

```vbnet
Sub FormulaRound()

    On Error GoTo ErrorHandler

    Dim Rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set Rng = Selection

    Dim regexTargetText As Object
    Set regexTargetText = New RegExp
    With regexTargetText
    .Pattern = "ROUNDDOWN"
    .Global = False
    End With

    Dim c As Range
    For Each c In Rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            If regexTargetText.Test(myFormula) Then
                myFormula = regexTargetText.Replace(myFormula, "")
                myFormula = Replace(myFormula, "(", "")
                myFormula = Replace(myFormula, ",0)", "")
                c.Formula = "=" & myFormula
            Else
                c.Formula = "=ROUNDDOWN(" & myFormula & ",0)"
            End If
        Else
            cellValue = c.Value
            c.Formula = "=ROUNDDOWN(" & cellValue & ",0)"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Round"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `FormulaThousands`

```vbnet
Sub FormulaThousands()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In XRELEVANTAREA(Rng)
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            c.Formula = "=ROUND(" & myFormula & "/1000,0)"
        Else
            c.Formula = "=ROUND(" & c.Value & "/1000,0)"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
    SaveSetting "ExcelUI", "Preferences", "LastSelNumber", "Thousands"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `FormulaToValue`

```vbnet
Sub FormulaToValue()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Copy
    Rng.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `InsertArrowDown`

```vbnet
Sub InsertArrowDown()

    On Error GoTo ErrorHandler
    
    Dim X1 As Long
    Dim X2 As Long
    Dim Y1 As Long
    Dim Y2 As Long
    
    Dim Line1 As Shape
    
    Dim mX1 As Long
    Dim mY1 As Long
    Dim mX2 As Long
    Dim mY2 As Long
    
    Dim Line2 As Shape
    
    Dim lCell As Range
    
    Set lCell = Selection.Cells(Selection.Rows.Count, Selection.Columns.Count) 'Last Cell
        
    Application.ScreenUpdating = False
    
    With Selection
        X1 = .Left + 10
        Y1 = .Top
    End With
        
    With lCell
        X2 = .Left + 10
        Y2 = .Top + .Height - 1.5
    End With
        
    With ActiveSheet.Shapes
        Set Line1 = .AddLine(X1, Y1, X2, Y2)
        Line1.Line.Weight = 0.5
        Line1.Line.BeginArrowheadStyle = msoArrowheadNone
        Line1.Line.EndArrowheadStyle = msoArrowheadTriangle
        Line1.Line.EndArrowheadWidth = msoArrowheadWidthMedium
        Line1.Line.EndArrowheadLength = msoArrowheadLengthMedium
        Line1.Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    With lCell
        mX1 = .Left + .Width / 2 - 4
        mX2 = .Left + .Width / 2 + 4
        mY1 = .Top + .Height - 1
        mY2 = .Top + .Height - 1
    End With
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `InsertBlankSheet`

```vbnet
Sub InsertBlankSheet()

    ActiveSheet.Select
    Sheets.Add.Name = "SourceData >>>"
    ActiveSheet.Tab.ColorIndex = 1
    
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Intentionally left blank"
    Range("B2").Select
    Selection.Font.Italic = True
    ActiveSheet.Select
    
End Sub
```

### `InsertColumnWidth`

```vbnet
Sub InsertColumnWidth()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Dim myFormula As String
    Set Rng = Selection

    Application.ScreenUpdating = False
    For Each c In Rng
        c.Formula = "=" & "XCOLUMNWIDTH(" & c.Address & ")"
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0.0_);_((#,##0.0);_(""-""??_);_(@_)"
    Next c
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `InsertCrossReference`

```vbnet
Sub InsertCrossReference()

    Call XHYPERACTIVE(Selection)

End Sub
```

### `InsertTimestamp`

```vbnet
Sub InsertTimestamp()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Value = Now
    Rng.NumberFormat = "dd-mmm-yy"
    Rng.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `InsertWorkdone`

```vbnet
Sub InsertWorkdone()

    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Selection
    
    Application.ScreenUpdating = False
    Rng.Value = "Keys to Workdone:"
    Rng.Font.Bold = True
    Rng.Offset(1, 0) = "TB"
    Rng.Offset(1, 1) = ": Agreed to current year trial balance."
    Rng.Offset(2, 0) = "PY"
    Rng.Offset(2, 1) = ": Agreed to prior year audited balance."
    Rng.Offset(3, 0) = "imm"
    Rng.Offset(3, 1) = ": Immaterial (below SUM), suggest to leave."
    Rng.Offset(4, 0) = "^"
    Rng.Offset(4, 1) = ": Casted."
    Rng.Offset(5, 0) = "Cal"
    Rng.Offset(5, 1) = ": Calculation checked."
    Rng.Offset(1, 0).Characters(1, 3).Font.Color = RGB(0, 112, 192)
    Rng.Offset(2, 0).Characters(1, 3).Font.Color = RGB(255, 51, 0)
    Rng.Offset(3, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(4, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(5, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
    Rng.Offset(1, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(2, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(3, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(4, 0).Characters(1, 3).Font.Bold = True
    Rng.Offset(5, 0).Characters(1, 3).Font.Bold = True
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `RemoveBlankCells`

```vbnet
Sub RemoveBlankCells()
'   Purpose: Remove blank cells in selection

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
        
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlUp
    Else
        If Application.WorksheetFunction.CountA(Selection) = 0 Then
            Selection.Delete Shift:=xlUp
        End If
    End If
    
    Selection.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub
```

### `RemoveBlankRows`

```vbnet
Sub RemoveBlankRows()
'   Purpose: Remove blank rows in selection

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
        
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.EntireRow.Delete
    Else
        If Application.WorksheetFunction.CountA(Selection) = 0 Then
            Selection.EntireRow.Delete
        End If
    End If
    
    Selection.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetColumnsFS`

```vbnet
Sub SheetColumnsFS()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Columns.ColumnWidth = 11
    Columns("A").ColumnWidth = 1
    Columns("B").ColumnWidth = 45
    Columns("C").ColumnWidth = 5
    Columns("D").ColumnWidth = 11
    Columns("E").ColumnWidth = 11
    Columns("F").ColumnWidth = 11
    Columns("G").ColumnWidth = 11
    Columns("H").ColumnWidth = 11
    Columns("I").ColumnWidth = 11
    Columns("J").ColumnWidth = 11
    Columns("K").ColumnWidth = 11
    Application.ScreenUpdating = True

ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetColumnsNTA4X`

```vbnet
Sub SheetColumnsNTA4X()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Columns.ColumnWidth = 11
    Columns("A").ColumnWidth = 1
    Columns("B").ColumnWidth = 5
    Columns("C").ColumnWidth = 45
    Columns("D").ColumnWidth = 11
    Columns("E").ColumnWidth = 11
    Columns("F").ColumnWidth = 11
    Columns("G").ColumnWidth = 11
    Columns("H").ColumnWidth = 11
    Columns("I").ColumnWidth = 11
    Columns("J").ColumnWidth = 11
    Columns("K").ColumnWidth = 11
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetColumnsWP`

```vbnet
Sub SheetColumnsWP()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Columns.ColumnWidth = 12
    Columns("A").ColumnWidth = 3
    Columns("B").ColumnWidth = 5
    Columns("C").ColumnWidth = 12
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontArial`

```vbnet
Sub SheetFontArial()

    On Error GoTo ErrorHandler
    
    If Not ActiveSheet.ProtectContents Then
        ActiveSheet.Cells.Font.Name = "Arial"
        ActiveSheet.Cells.Font.Size = 8
        ActiveSheet.Activate
        ActiveWindow.Zoom = 100
    Else: Exit Sub
    End If
    
    SaveSetting "ExcelUI", "Preferences", "LastShFont", "Arial"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontEY`

```vbnet
Sub SheetFontEY()

    On Error GoTo ErrorHandler
    
    If Not ActiveSheet.ProtectContents Then
        ActiveSheet.Cells.Font.Name = "EYInterstate Light"
        ActiveSheet.Cells.Font.Size = 8
        ActiveSheet.Activate
        ActiveWindow.Zoom = 100
    Else: Exit Sub
    End If
        
    SaveSetting "ExcelUI", "Preferences", "LastShFont", "EY"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontSize10`

```vbnet
Sub SheetFontSize10()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 10
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFontSize8`

```vbnet
Sub SheetFontSize8()

    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.Font.Size = 8
    ActiveSheet.Activate
    ActiveWindow.Zoom = 100

ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetFormulaToValue`

```vbnet
Sub SheetFormulaToValue()

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    ActiveSheet.Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetPageBreakOff`

```vbnet
Sub SheetPageBreakOff()

    On Error GoTo ErrorHandler
    
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.Activate
    ActiveWindow.DisplayGridlines = False
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetRemoveBlankRows`

```vbnet
Sub SheetRemoveBlankRows()

    On Error GoTo ErrorHandler
    
    Dim SourceRange As Range
    Dim EntireRow As Range
    Set SourceRange = Application.ActiveSheet.UsedRange
    
    Application.ScreenUpdating = False
    If Not (SourceRange Is Nothing) Then
        For i = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
    End If
    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub
        
End Sub
```

### `SheetResetComments`

```vbnet
Sub SheetResetComments()
    
    On Error GoTo ErrorHandler
    
    Dim pComment As Comment
    For Each pComment In Application.ActiveSheet.Comments
       pComment.Shape.Top = pComment.Parent.Top + 5
       pComment.Shape.Left = pComment.Parent.Offset(0, 1).Left + 5
    Next
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabBlack`

```vbnet
Sub SheetTabBlack()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 1
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Black"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabGreen`

```vbnet
Sub SheetTabGreen()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 35
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Green"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `SheetTabRed`

```vbnet
Sub SheetTabRed()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 38
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Red"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabReset`

```vbnet
Sub SheetTabReset()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = xlColorIndexNone
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Reset"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `SheetTabYellow`

```vbnet
Sub SheetTabYellow()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet

    For Each ws In ActiveWindow.SelectedSheets
         ws.Tab.ColorIndex = 6
    Next ws
    
    SaveSetting "ExcelUI", "Preferences", "LastShTab", "Yellow"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
        
ErrorHandler:
    Exit Sub

End Sub
```

### `UnhideSheets`

```vbnet
Sub UnhideSheets()

    On Error GoTo ErrorHandler
    
    For Each sh In Worksheets: sh.Visible = True: Next sh
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `WorkbookArial`

```vbnet
Sub WorkbookArial()
  
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                .Cells.Font.Name = "Arial"
                .Cells.Font.Size = 8
            End If
         End With
    Next ws
    
    For Each ws In Worksheets
        If Not ws.ProtectContents Then
            ws.Activate
            ActiveWindow.Zoom = 100
        End If
    Next
    Application.ScreenUpdating = True
    
    Call sourceSheet.Activate
    
    SaveSetting "ExcelUI", "Preferences", "LastWbFont", "Arial"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `WorkbookEY`

```vbnet
Sub WorkbookEY()
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                .Cells.Font.Name = "EYInterstate Light"
                .Cells.Font.Size = 8
            End If
         End With
    Next ws
    For Each ws In Worksheets
        If Not ws.ProtectContents Then
            ws.Activate
            ActiveWindow.Zoom = 100
        End If
    Next
    Application.ScreenUpdating = True
    
    Call sourceSheet.Activate

    SaveSetting "ExcelUI", "Preferences", "LastWbFont", "EY"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub
    
End Sub
```

### `WorkbookPageBreakOff`

```vbnet
Sub WorkbookPageBreakOff()
  
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
         With ws
            If Not ws.ProtectContents Then
                ws.DisplayPageBreaks = False
                ws.Activate
                ActiveWindow.DisplayGridlines = False
            End If
         End With
    Next ws
    Application.ScreenUpdating = True
    
    Call sourceSheet.Activate

ErrorHandler:
    Exit Sub
           
End Sub
```

### `WorkbookFontSize8`

```vbnet
Sub WorkbookFontSize8()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 8
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "8"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `WorkbookFontSize9`

```vbnet
Sub WorkbookFontSize9()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 9
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "9"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `WorkbookFontSize10`

```vbnet
Sub WorkbookFontSize10()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 10
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "10"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `WorkbookFontSize11`

```vbnet
Sub WorkbookFontSize11()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Size = 11
    Next ws

    SaveSetting "ExcelUI", "Preferences", "LastWbFontSize", "11"
    If Not RibbonUI Is Nothing Then RibbonUI.Invalidate
    
ErrorHandler:
    Exit Sub

End Sub
```

### `RunFormatHighlightRepeat`

```vbnet
Sub RunFormatHighlightRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastHighlight", "Green")
    Select Case last
        Case "Green":  FormatHighlightGreen
        Case "Red":    FormatHighlightRed
        Case "Yellow": FormatHighlightYellow
        Case Else:     FormatHighlightReset
    End Select
End Sub
```

### `RunWorkbookFontRepeat`

```vbnet
Sub RunWorkbookFontRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastWbFont", "Arial")
    Select Case last
        Case "Arial": WorkbookArial
        Case "EY":    WorkbookEY
    End Select
End Sub
```

### `RunWorkbookFontSizeRepeat`

```vbnet
Sub RunWorkbookFontSizeRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastWbFontSize", "10")
    Select Case last
        Case "8":  WorkbookFontSize8
        Case "9":  WorkbookFontSize9
        Case "10": WorkbookFontSize10
        Case "11": WorkbookFontSize11
    End Select
End Sub
```

### `RunSheetFontRepeat`

```vbnet
Sub RunSheetFontRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastShFont", "Arial")
    Select Case last
        Case "Arial": SheetFontArial
        Case "EY":    SheetFontEY
    End Select
End Sub
```

### `RunSheetTabRepeat`

```vbnet
Sub RunSheetTabRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastShTab", "Green")
    Select Case last
        Case "Green":  SheetTabGreen
        Case "Red":    SheetTabRed
        Case "Yellow": SheetTabYellow
        Case Else:     SheetTabReset
    End Select
End Sub
```

### `RunSelNumberRepeat`

```vbnet
Sub RunSelNumberRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastSelNumber", "Accounting")
    Select Case last
        Case "Accounting":   FormatAccounting
        Case "Round":        FormulaRound
        Case "Absolute":     FormulaAbsolute
        Case "Reverse Sign": FormulaReverseSign
    End Select
End Sub
```

### `RunSelCaseRepeat`

```vbnet
Sub RunSelCaseRepeat()
    Dim last As String
    last = GetSetting("ExcelUI", "Preferences", "LastSelCase", "Proper")
    Select Case last
        Case "Proper":   CaseProper
        Case "Upper":    CaseUpper
        Case "Sentence": CaseSentence
    End Select
End Sub
```
