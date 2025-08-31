Private Function XGETDIGIT(Digit)
'   Purpose: Converts a number from 1 to 9 into text.

    Select Case Val(Digit)
        Case 1: XGETDIGIT = "One"
        Case 2: XGETDIGIT = "Two"
        Case 3: XGETDIGIT = "Three"
        Case 4: XGETDIGIT = "Four"
        Case 5: XGETDIGIT = "Five"
        Case 6: XGETDIGIT = "Six"
        Case 7: XGETDIGIT = "Seven"
        Case 8: XGETDIGIT = "Eight"
        Case 9: XGETDIGIT = "Nine"
        Case Else: XGETDIGIT = ""
    End Select
    
End Function

