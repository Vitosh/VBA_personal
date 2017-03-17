Option Explicit

Private objPresenter   As clsSummaryPresenter

Public Sub MainGenerateReport()
    
    Call objPresenter.ChangeLabelAndCaption("Starting and running...", "Running...")
    GenerateNumbers
        
End Sub

Public Sub GenerateNumbers()
    
    Dim lngLong         As Long
    Dim lngLong2        As Long
    
    tblMain.Cells.Clear
    
    For lngLong = 1 To 3000
        For lngLong2 = 1 To 10
            tblMain.Cells(lngLong, lngLong2) = lngLong * lngLong2
        Next lngLong2
    Next lngLong

End Sub

Public Sub ShowMainForm() 'CTRL+E

    If (objPresenter Is Nothing) Then Set objPresenter = New clsSummaryPresenter
    objPresenter.Show

End Sub

