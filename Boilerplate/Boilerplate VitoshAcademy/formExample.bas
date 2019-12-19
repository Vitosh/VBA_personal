Attribute VB_Name = "formExample"
Option Explicit
Option Private Module

Private presenter As formSummaryPresenter

Public Sub FormExampleMain()
    
    presenter.ChangeLabelAndCaption "Starting and running...", "Running..."
    GenerateNumbers

End Sub

Public Sub GenerateNumbers(Optional outerLoopLimit As Long = 2, Optional innerLoopLimit As Long = 4)
    
    Dim a As Long
    Dim b As Long
    
    For a = 1 To outerLoopLimit
        For b = 1 To innerLoopLimit
            Debug.Print a * b
        Next
    Next
    Debug.Print "-------END-------" & vbCrLf & Now
    
End Sub

Public Sub ShowMainForm()

    If (presenter Is Nothing) Then
        Set presenter = New formSummaryPresenter
    End If

    presenter.Show

End Sub

Public Sub CheckHowManyWbAreOpened()

    On Error GoTo CheckHowManyWbAreOpened_Error

    If Workbooks.Count > 1 Then
        PUB_STR_ERROR_REPORT = True
        frmInfo.Show (vbModeless)
        Application.Wait (Now + TimeValue("00:00:02"))
        Unload frmInfo
    End If
    
    PUB_STR_ERROR_REPORT = False

    On Error GoTo 0
    Exit Sub

CheckHowManyWbAreOpened_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckHowManyWbAreOpened of Sub DieseArbeitsmappe"

End Sub

