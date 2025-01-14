Sub RemoveAccented()
'	Purpose: Remove accented characters in cell
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Selection.Replace ".", "_"
    Selection.Replace "'", ""
    Selection.Replace Chr(225), "a"
    Selection.Replace Chr(227), "a"
    Selection.Replace Chr(228), "a"
    Selection.Replace Chr(230), "ae"
    Selection.Replace Chr(223), "b"
    Selection.Replace ChrW(263), "c"
    Selection.Replace ChrW(269), "c"
    Selection.Replace Chr(199), "C"
    Selection.Replace Chr(232), "e"
    Selection.Replace Chr(233), "e"
    Selection.Replace Chr(234), "e"
    Selection.Replace Chr(235), "e"
    Selection.Replace ChrW(287), "g"
    Selection.Replace Chr(237), "i"
    Selection.Replace Chr(239), "i"
    Selection.Replace ChrW(304), "I"
    Selection.Replace ChrW(321), "L"
    Selection.Replace ChrW(324), "n"
    Selection.Replace Chr(246), "o"
    Selection.Replace Chr(248), "o"
    Selection.Replace Chr(214), "O"
    Selection.Replace Chr(216), "O"
    Selection.Replace ChrW(345), "r"
    Selection.Replace ChrW(352), "S"
    Selection.Replace Chr(250), "u"
    Selection.Replace Chr(251), "u"
    Selection.Replace Chr(252), "u"
    
ErrorHandler:
    Exit Sub

End Sub