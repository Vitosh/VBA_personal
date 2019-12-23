Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If ActiveWindow.Zoom > 100 Or ActiveWindow.Zoom < 70 Then
        ActiveWindow.Zoom = 100
    End If
    
End Sub
