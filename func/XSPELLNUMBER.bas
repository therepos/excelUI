Function XSPELLNUMBER(ByVal MyNumber)
'   Purpose: Spell numbers as dollars
'   Source: https://support.microsoft.com/en-us/office/convert-numbers-into-words-a0d166fb-e1ea-4090-95c8-69442cd55d98
'   Reference: https://stackoverflow.com/questions/11155912/how-to-make-vba-function-vba-only-and-disable-it-as-udf/41130822
'   Modification:
'   - Made private XGETHUNDREDS, XGETTENS. Users do not need to access these.
'   - Added "Only" to the end of result.
'   - Fixed additional spacing between words in the original code.
'   - Fixed 0.XX appearing as "No Dollar and XX Cents" to "XX Cents Only"
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "

'   String representation of amount
    MyNumber = Trim(str(MyNumber))
    
'   Position of decimal place 0 if none
    DecimalPlace = InStr(MyNumber, ".")
    
'   Convert cents and set MyNumber to dollar amount
    If DecimalPlace > 0 Then
        Cents = XGETTENS(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    Count = 1
    Do While MyNumber <> ""
        Temp = XGETHUNDREDS(Right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    
    Select Case Dollars
        Case "": Dollars = "" ' Originally "No Dollar"
        Case "One": Dollars = "One Dollar"
        Case Else
        Dollars = Dollars & " Dollars"
    End Select
    
    Select Case Cents
        Case "": Cents = " Only"
        Case "One": Cents = " and One Cent Only"
        Case Else
        Cents = " and " & Cents & " Cents Only"
    End Select
    
    If Dollars = "" Then
        XSPELLNUMBER = Trim(Replace(Replace(Dollars & Cents, "  ", " "), "and", ""))
    Else
        XSPELLNUMBER = Trim(Replace(Dollars & Cents, "  ", " "))
    End If
    Application.ScreenUpdating = True

End Function

