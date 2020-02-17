'Create Make Sheet Worksheet
'Remove Sheet Worksheet
'Delete Sheet Worksheet

Sub CreateWorksheet(sheetName As String)

    ThisWorkbook.Worksheets.Add.Name = sheetName
        
End Sub

Sub DeleteWorksheet(sheetName As String)

    Dim displayAlert As Boolean
    Dim mySheet As Worksheet
    
    displayAlert = Application.DisplayAlerts
    
    For Each mySheet In ThisWorkbook.Worksheets
        If mySheet.Name = sheetName Then
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets(sheetName).Delete
            Application.DisplayAlerts = displayAlert
        End If
    Next
    
End Sub

Sub DeleteAllButLast()

    Dim wksToStay As Worksheet
    Dim wksToDelete As Worksheet
    Dim i As Long

    Set wksToStay = ThisWorkbook.Worksheets(Worksheets.Count)

    For i = Worksheets.Count To 1 Step -1
        Set wksToDelete = ThisWorkbook.Worksheets(i)
        If wksToDelete.Name <> wksToStay.Name Then
            Application.DisplayAlerts = False
            wksToDelete.Delete
            Application.DisplayAlerts = True
        End If
    Next

End Sub

