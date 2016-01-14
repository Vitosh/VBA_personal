Public Sub remove_space_in_string()

    Dim r_range As Range
        
    For Each r_range In Selection
        r_range.Value = RTrim(r_range.Value)
        r_range.Cells.Font.Bold = False
        r_range = Replace(r_range, " ", "")
        
    Next r_range

End Sub
