Attribute VB_Name = "AmazonInternet"
Option Explicit

Public Function PageWithResultsExists(appIE As Object, keyword As String) As Boolean

    On Error GoTo PageWithResultsExists_Error
    
    Dim allData As Object
    Set allData = appIE.document.getElementById("s-results-list-atf")
    PageWithResultsExists = True
    IeErrors = 0
    
    On Error GoTo 0
    Exit Function

PageWithResultsExists_Error:
    
    WaitSomeMilliseconds
    IeErrors = IeErrors + 1
    
    Select Case Err.Number
        
        Case 424
            
            If IeErrors > MAX_IE_ERRORS Then
                PageWithResultsExists = False
                IeErrors = 0
            Else
                LogMe "PageWithResultsExists", IeErrors, keyword, IeErrors
                PageWithResultsExists appIE, keyword
            End If
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    
End Function

Public Function MakeUrl(i As Long, keyword As String) As String

    MakeUrl = "https://www.amazon.com/s/ref=sr_pg_" & i & "?rh=i%3Aaps%2Ck%3A" & keyword & "&page=" & i & "&keywords=" & keyword

End Function

Public Sub Navigate(i As Long, appIE As Object, keyword As String)
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    With appIE
        .Navigate MakeUrl(i, keyword)
        .Visible = False
    End With
    
    Do While appIE.Busy
        DoEvents
    Loop
    
End Sub
