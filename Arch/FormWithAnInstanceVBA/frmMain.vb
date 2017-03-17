Option Explicit

Public Event OnRunReport()
Public Event OnExit()

Private Sub btnRun_Click()
    RaiseEvent OnRunReport
End Sub

Private Sub btnExit_Click()
    RaiseEvent OnExit
End Sub
