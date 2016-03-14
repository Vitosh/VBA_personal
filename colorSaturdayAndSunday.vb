Public Sub ColorSS()
    
    On Error GoTo ColorSS_Error
    
    'Colors Saturdays and Sundays.
    
    Dim r_cell      As Range
    Dim r_range     As Range
    
    For Each r_cell In Selection
        If Weekday(r_cell.Value) = 1 Or Weekday(r_cell.Value) = 7 Then
            Set r_range = ActiveSheet.Range(Cells(4, r_cell.Column), Cells(340, r_cell.Column))
            r_range.Interior.Color = 13434828
        End If
    Next r_cell
    
    Set r_range = Nothing

    On Error GoTo 0
    Exit Sub

ColorSS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ColorSS of Sub mod_play_with_me"
End Sub
