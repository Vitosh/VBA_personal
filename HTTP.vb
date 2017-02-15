Option Explicit

Public Sub TestMe()

    Dim oRequest    As Object
    Dim strOb       As String
    Dim strInfo     As String: strInfo = "class=""question-hyperlink"">"
    Dim lngStart    As Long
    Dim lngEnd      As Long

    Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    With oRequest
        .Open "GET", "http://stackoverflow.com/questions/42254051/vba-open-website-find-specific-value-and-return-value-to-excel#42254254", True
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .Send "{range:9129370}"
        .WaitForResponse
        strOb = .ResponseText

    End With

    lngStart = InStr(1, strOb, strInfo)
    lngEnd = InStr(lngStart, strOb, "<")

    Debug.Print Mid(strOb, lngStart + Len(strInfo), lngEnd - lngStart - Len(strInfo))

End Sub
