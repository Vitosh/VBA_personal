Option Explicit

Private Sub Workbook_NewSheet(ByVal Sh As Object)

   On Error GoTo Workbook_NewSheet_Error

    Sheets(1).Rows("1:2").Copy
    Sh.Paste
    Application.CutCopyMode = False
    
    'Sheets(1).Columns(1).Copy
    Sheets(1).Columns("A:D").Copy
    Sh.Paste
    Application.CutCopyMode = False
    
    Sh.Cells(1, 1).Select
    
   On Error GoTo 0
   Exit Sub

Workbook_NewSheet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_NewSheet of VBA Document DieseArbeitsmappe"
End Sub

