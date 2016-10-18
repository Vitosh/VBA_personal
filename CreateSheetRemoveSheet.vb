'Create Sheet
'Make Sheet
'Remove Sheet

Sub CreateSheet(str_name As String)

    Sheets.Add.Name = str_name
    tbl_total_s = Worksheets(str_name).CodeName
    Stop
End Sub

Sub DeleteSheet(str_name As String)

    Dim b_display_alerts    As Boolean
    Dim my_sheet            As Worksheet
    
    b_display_alerts = Application.DisplayAlerts
    
    For Each my_sheet In ActiveWorkbook.Worksheets
        If my_sheet.Name = str_name Then
            Application.DisplayAlerts = False
            Worksheets(str_name).Delete
            Application.DisplayAlerts = b_display_alerts
        End If
    Next my_sheet
    
End Sub
