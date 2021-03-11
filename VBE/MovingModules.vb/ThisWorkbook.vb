Option Explicit

Private Sub Workbook_Open()

    Dim i As Long
    
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
        ActiveWorkbook.Worksheets(i).Protect Password:=s_CONST
    Next
    
    Application.DisplayAlerts = False
    
End Sub
