Private Function GetUTCTimeDate() As Date
'   Purpose: Get Internet Time
'   Reference: https://stackoverflow.com/questions/48371398/get-date-from-internet-and-compare-to-system-clock-on-workbook-open
'   Reference: https://stackoverflow.com/questions/551613/check-for-active-internet-connection
'   Reference: http://excelerator.solutions/2017/08/28/excel-http-get-request/
'   Requirement: Microsoft Scripting Runtime, Microsoft Internet Controls, and Microsoft WinHTTP

    Dim UTCDateTime As String
    Dim arrDT() As String
    Dim http As Object
    Dim UTCDate As String
    Dim UTCTime As String

    Const NetTime As String = "https://www.time.gov/"
    
    On Error Resume Next
    Set http = CreateObject("Microsoft.XMLHTTP")
    On Error GoTo 0

    On Error Resume Next
    http.Open "GET", NetTime & Now(), False, "", ""
    http.send
    
    UTCDateTime = http.GetResponseHeader("Date")
    UTCDate = Mid(UTCDateTime, InStr(UTCDateTime, ",") + 2)
    UTCDate = Left(UTCDate, InStrRev(UTCDate, " ") - 1)
    UTCTime = Mid(UTCDate, InStrRev(UTCDate, " ") + 1)
    UTCDate = Left(UTCDate, InStrRev(UTCDate, " ") - 1)
    GetUTCTimeDate = DateValue(UTCDate)
    On Error GoTo 0

End Function

