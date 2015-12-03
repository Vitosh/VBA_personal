Option Explicit

Private Sub Workbook_NewSheet(ByVal Sh As Object)

   On Error GoTo Workbook_NewSheet_Error

    Tabelle1.Rows("1:2").Copy
    Sh.Paste
    Application.CutCopyMode = False
    Tabelle1.Columns("A:A").Copy
    Sh.Paste
    Application.CutCopyMode = False

   On Error GoTo 0
   Exit Sub

Workbook_NewSheet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_NewSheet of VBA Document DieseArbeitsmappe"
End Sub
