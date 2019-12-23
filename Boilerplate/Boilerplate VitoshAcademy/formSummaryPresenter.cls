VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "formSummaryPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents summaryForm As frmExample
Attribute summaryForm.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set summaryForm = New frmExample

End Sub

Private Sub Class_Terminate()

    Set summaryForm = Nothing

End Sub

Public Sub Show()

    If Not summaryForm.Visible Then
        summaryForm.Show vbModeless
        ChangeLabelAndCaption "Press Run to Start", "Starting"
    End If

    With summaryForm
        .Top = CLng((Application.Height / 2 + Application.Top) - .Height / 2)
        .Left = CLng((Application.Width / 2 + Application.Left) - .Width / 2)
        .caption = CON_STR_APP_NAME
    End With

End Sub

Private Sub Hide()

    If summaryForm.Visible Then summaryForm.Hide

End Sub

Public Sub ChangeLabelAndCaption(labelInfo As String, caption As String)

    summaryForm.InformationText = labelInfo
    summaryForm.InformationCaption = caption
    summaryForm.Repaint

End Sub

Private Sub summaryForm_OnRunReport()
    
    FormExampleMain
    Refresh

End Sub

Private Sub summaryForm_OnExit()

    Hide

End Sub

Public Sub Refresh()
    
    With summaryForm
        .lblInfo = "Ready"
        .caption = "Task performed"
    End With

End Sub
