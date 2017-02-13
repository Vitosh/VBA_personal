 Option Explicit

Private WithEvents objSummaryForm As frmMain

Private Sub Class_Initialize()
    
    Set objSummaryForm = New frmMain

End Sub

Private Sub Class_Terminate()
    
    Set objSummaryForm = Nothing
    
End Sub

Public Sub Show()

    If Not objSummaryForm.Visible Then
        objSummaryForm.Show vbModeless
        objSummaryForm.InformationText = "Press Run to Start"
        objSummaryForm.InformationCaption = "Starting"
    End If

End Sub

Public Sub Hide()

    If objSummaryForm.Visible Then objSummaryForm.Hide

End Sub

Public Sub ChangeLabelAndCaption(strLabelInfo As String, strCaption As String)

    objSummaryForm.InformationText = strLabelInfo
    objSummaryForm.InformationCaption = strCaption
    objSummaryForm.Repaint

End Sub


Private Sub objSummaryForm_OnRunReport()

    MainGenerateReport
    Refresh

End Sub

Private Sub objSummaryForm_OnExit()
    
    Hide

End Sub

Public Sub Refresh()
    
    With objSummaryForm
        .lblInfo = "Ready"
        .Caption = "Task performed"
    End With

End Sub

