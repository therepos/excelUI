' ============================================================
' RETAIN AVAILABILITY GENERATOR
' Reads "ER and P&C" sheet and generates "Retain Availability"
' in a NEW workbook
' ============================================================

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



