Attribute VB_Name = "Module1"
Option Explicit
Public Sub Http_Request()

Dim xmlHttp As New MSXML2.XMLHTTP60
Dim aw As Workbook
Dim response As String, url As String
Dim ind As Long, i As Integer, latestDate As Date

Set aw = ActiveWorkbook
url = aw.Sheets(1).Range("C6").Value 'URL to send requests to (cell C6)
latestDate = aw.Sheets(1).Range("C5").Value
'if blank, use yesterday's date
If IsNull(latestDate) Or latestDate = 0 Then latestDate = Date - 1

aw.ActiveSheet.Range("E5:F25").ClearContents
For i = 1 To days 'send HTTP requests with parameter in a loop for # of days
    response = ""
    With xmlHttp
        .Open "GET", url & Format(latestDate, "yyyy-mm-dd"), False 'Sync code
        .setRequestHeader "Content-Type", "text/json"
        .send
        'some exception handling
        If .readyState = 4 Then
            If .Status = 200 Then
                response = .responseText
            Else
                MsgBox "HTTP 300 or 400 error recieved"
            End If
        Else
            MsgBox ("Request has not completed")
        End If
    End With
    If response <> "" Then ind = InStr(response, "INDEX PUT")
    
    aw.ActiveSheet.Range("E4").Offset(i, 0).Value = latestDate
    If ind > 0 Then
        aw.ActiveSheet.Range("E4").Offset(i, 1) = _
        Round(CSng(Mid(response, ind + 54, 4)), 2)
    End If
    
    'decrement date by 1; 2 or 3 for weekend
    latestDate = latestDate - IIf(Weekday(latestDate) >= 3, 1, 1 + Weekday(latestDate))
Next i

End Sub





