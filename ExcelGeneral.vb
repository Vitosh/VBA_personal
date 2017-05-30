Public Sub CloseAllExcelFilesExceptCurrent()

    Dim wb As Workbook
    
    Application.ScreenUpdating = False
    
    For Each wb In Workbooks

        If Not wb.ReadOnly Then wb.Save
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close
        End If
    Next wb
    
End Sub
