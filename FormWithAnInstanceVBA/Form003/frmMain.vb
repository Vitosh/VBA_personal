Option Explicit

Public Event OnRunReport()
Public Event OnExit()

Public Property Get InformationText() As String
    
    InformationText = lblInfo.Caption

End Property

Public Property Let InformationText(ByVal value As String)
    
    lblInfo.Caption = value

End Property

Public Property Get InformationCaption() As String
    
    InformationCaption = Caption

End Property

Public Property Let InformationCaption(ByVal value As String)
    
    Caption = value

End Property


Private Sub btnRun_Click()
    RaiseEvent OnRunReport
End Sub

Private Sub btnExit_Click()
    RaiseEvent OnExit
End Sub

Private Sub UserForm_QueryClose(CloseMode As Integer, Cancel As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If

End Sub
