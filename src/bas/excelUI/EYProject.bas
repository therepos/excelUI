' ============================================================
' EXTRACT JOBS BY MANAGER - GANTT/TIMELINE VIEW
' Reads "ER and P&C" sheet and generates a weekly timeline
' showing all jobs managed by a specified person
' ============================================================

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



