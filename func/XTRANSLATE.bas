Function XTRANSLATE(strInput As String, strFromSourceLanguage As String, strToTargetLanguage As String) As String
'   Purpose: Translate with Google Translate
'   Reference: https://www.youtube.com/watch?v=RsyqqzholVk&ab_channel=DineshKumarTakyar
'   Usage: = Translate(Range("A2"), "en", "es")
'   Todo: Create optional source language input. Default to auto detect language
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save

    Dim strURL As String
    Dim objHTTP As Object
    Dim objHTML As Object
    Dim objDivs As Object, objDiv As Object
    Dim strTranslated As String

    ' send query to web page (google translate mobile)
    strURL = "https://translate.google.com/m?hl=" & strFromSourceLanguage & _
        "&sl=" & strFromSourceLanguage & _
        "&tl=" & strToTargetLanguage & _
        "&ie=UTF-8&prev=_m&q=" & strInput

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP") 'late binding
    objHTTP.Open "GET", strURL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ""

    ' create an html document
    Set objHTML = CreateObject("htmlfile")
    With objHTML
        .Open
        .Write objHTTP.responsetext
        .Close
    End With

    Set objDivs = objHTML.getElementsByTagName("div")

    For Each objDiv In objDivs

        If objDiv.className = "result-container" Then
            strTranslated = objDiv.innerText
            Translate = strTranslated
        End If

    Next objDiv

    Set objHTML = Nothing
    Set objHTTP = Nothing
    
End Function

