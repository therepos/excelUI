' ============================================================
' RETAIN COLOUR NORMALISER
' ------------------------------------------------------------
' Resets the "amber / change" markers on a Retain sheet back to
' the standard booking colours defined in the workbook's own
' "Reference" key:
'
'   Green (D8E4BC) : Staff booking LESS than 70% for the week
'   White          : Staff booking 70% or MORE for the week
'   Amber (FFC000) : "Assistance from Retain team required" -
'                    a temporary change / action marker
'   Pink  (FF0066) : Staff on training (LEFT UNTOUCHED)
'
' For every amber-marked weekly booking cell this recolours it to:
'   - WHITE if the week's booking is >= 70%
'   - GREEN if the week's booking is  < 70%
'
' "Booking" is measured the same way as the availability report
' (ExtractTotalPercent in module EYRetain): BD (Business
' Development) is treated as non-chargeable / available and does
' NOT count towards the 70%. Set USE_BD_EXCLUDED_BOOKING = False
' to instead count every percentage in the cell.
'
' FAIL-SAFE (point 2): an orange/amber cell is detected by its
' COLOUR HUE, not by an exact hex match - so if the user picks a
' "wrong orange" (any other amber/orange shade) it is still
' recognised and normalised. Cells that used a non-standard
' orange are listed in the summary so the wrong colour can be
' spotted and the key colour (FFC000) used next time.
'
' Only cells inside the weekly date grid are touched, so the
' Reference legend swatches and other UI are never altered.
' ============================================================

Private Const COLOUR_THRESHOLD As Long = 70
Private Const USE_BD_EXCLUDED_BOOKING As Boolean = True

Sub NormaliseRetainColours()

    Dim ws As Worksheet

    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is open." & vbCrLf & vbCrLf & _
            "Open the Retain file and try again.", _
            vbExclamation, "Retain Colour Normaliser"
        Exit Sub
    End If

    Set ws = ActiveSheet

    ' ============================================================
    ' AUTO-DETECT the weekly date header row and date columns
    ' (same approach as GenerateRetainAvailability)
    ' ============================================================
    Dim headerRow As Long, dateStartCol As Long
    Dim scanRow As Long, scanCol As Long
    headerRow = 0
    dateStartCol = 0

    For scanRow = 1 To 20
        For scanCol = 1 To 50
            If Not IsEmpty(ws.Cells(scanRow, scanCol).Value) Then
                If Not IsError(ws.Cells(scanRow, scanCol).Value) Then
                    If IsDate(ws.Cells(scanRow, scanCol).Value) Then
                        If scanCol < 50 Then
                            If Not IsError(ws.Cells(scanRow, scanCol + 1).Value) Then
                                If IsDate(ws.Cells(scanRow, scanCol + 1).Value) Then
                                    headerRow = scanRow
                                    dateStartCol = scanCol
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
    If headerRow = 0 Then
        MsgBox "Could not find the weekly date headers on """ & ws.Name & """." & vbCrLf & vbCrLf & _
            "Switch to the correct Retain tab and try again.", _
            vbExclamation, "Retain Colour Normaliser"
        Exit Sub
    End If

    Dim lastCol As Long, lastRow As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    ' --- Collect the weekly date columns ---
    Dim dateCols() As Long, nDates As Long
    nDates = 0
    Dim c As Long
    For c = dateStartCol To lastCol
        If Not IsError(ws.Cells(headerRow, c).Value) Then
            If IsDate(ws.Cells(headerRow, c).Value) Then
                nDates = nDates + 1
                ReDim Preserve dateCols(1 To nDates)
                dateCols(nDates) = c
            End If
        End If
    Next c

    If nDates = 0 Then
        MsgBox "No weekly date columns detected on """ & ws.Name & """.", _
            vbExclamation, "Retain Colour Normaliser"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim greenClr As Long, amberClr As Long, whiteClr As Long
    greenClr = RGB(216, 228, 188)   ' D8E4BC  (standard "below 70%" green)
    amberClr = RGB(255, 192, 0)     ' FFC000  (standard change/action amber)
    whiteClr = RGB(255, 255, 255)

    Dim toWhite As Long, toGreen As Long, wrongOrange As Long
    Dim wrongList As String
    toWhite = 0
    toGreen = 0
    wrongOrange = 0
    wrongList = ""

    Dim dataStartRow As Long
    dataStartRow = headerRow + 1

    Dim R As Long, i As Long
    For R = dataStartRow To lastRow
        For i = 1 To nDates
            Dim cell As Range
            Set cell = ws.Cells(R, dateCols(i))

            If cell.Interior.Pattern <> xlNone Then
                Dim clr As Long
                clr = cell.Interior.Color

                If IsAmberish(clr) Then
                    ' Fail-safe: flag non-standard orange shades
                    If clr <> amberClr Then
                        wrongOrange = wrongOrange + 1
                        If Len(wrongList) < 250 Then _
                            wrongList = wrongList & cell.Address(False, False) & " "
                    End If

                    ' Determine the week's booking %
                    Dim booking As Long
                    If IsError(cell.Value) Then
                        booking = 0
                    ElseIf USE_BD_EXCLUDED_BOOKING Then
                        booking = ExtractTotalPercent(CStr(cell.Value & ""))
                    Else
                        booking = RawTotalPercent(CStr(cell.Value & ""))
                    End If

                    If booking >= COLOUR_THRESHOLD Then
                        cell.Interior.Color = whiteClr
                        toWhite = toWhite + 1
                    Else
                        cell.Interior.Color = greenClr
                        toGreen = toGreen + 1
                    End If
                End If
            End If
        Next i
    Next R

    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Retain colours normalised on """ & ws.Name & """." & vbCrLf & vbCrLf & _
        "Amber markers reset:" & vbCrLf & _
        "   -> WHITE (booking >= " & COLOUR_THRESHOLD & "%): " & toWhite & vbCrLf & _
        "   -> GREEN (booking <  " & COLOUR_THRESHOLD & "%): " & toGreen & vbCrLf & vbCrLf & _
        "Booking basis: " & IIf(USE_BD_EXCLUDED_BOOKING, _
            "BD excluded (chargeable time only)", "all percentages counted")

    If wrongOrange > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
            "FAIL-SAFE: " & wrongOrange & " cell(s) used a NON-STANDARD orange" & vbCrLf & _
            "(the standard amber in the Reference key is FFC000)." & vbCrLf & _
            "They were still normalised. Cells: " & Trim(wrongList)
    End If

    MsgBox msg, vbInformation, "Retain Colour Normaliser"

End Sub

' --- Fail-safe hue test: is this colour in the orange/amber family? ---
' Catches FFC000 and any other orange/amber shade a user might pick,
' while excluding the yellow header (FFFF00), the green marker
' (D8E4BC), the pink marker (FF0066) and white.
Private Function IsAmberish(ByVal clr As Long) As Boolean
    Dim R As Long, G As Long, B As Long
    R = clr Mod 256
    G = (clr \ 256) Mod 256
    B = (clr \ 65536) Mod 256

    ' Strong red, mid green, low blue, with R > G > B.
    ' Bounds are deliberately wide so any orange/amber/gold shade a
    ' user might pick is caught, while yellow (FFFF00), the green
    ' marker (D8E4BC) and the pink marker (FF0066) stay excluded.
    IsAmberish = (R >= 200) And (B <= 150) And _
                 (G >= 100) And (G <= 220) And _
                 (R > G) And (G > B)
End Function

' --- Booking total counting EVERY percentage (BD included) ---
' Used only when USE_BD_EXCLUDED_BOOKING = False. Mirrors the
' original ExtractTotalPercent behaviour before BD exclusion.
Private Function RawTotalPercent(ByVal cellText As String) As Long
    Dim total As Long, i As Long
    Dim numStr As String, ch As String
    total = 0
    numStr = ""

    If Len(Trim(cellText)) = 0 Then
        RawTotalPercent = 0
        Exit Function
    End If

    If IsNumeric(cellText) Then
        Dim v As Double
        v = CDbl(cellText)
        If v = 0 Then
            RawTotalPercent = 0
        ElseIf v > 0 And v <= 1 Then
            RawTotalPercent = CLng(v * 100)
        Else
            RawTotalPercent = CLng(v)
        End If
        Exit Function
    End If

    For i = 1 To Len(cellText)
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

    RawTotalPercent = total
End Function
