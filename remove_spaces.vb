Public Sub remove_space_in_string()

    Dim r_range As Range
        
    For Each r_range In Selection
        r_range = Trim(r_range)
        r_range = Replace(r_range, vbTab, "")
        r_range = Replace(r_range, " ", "")
        r_range = Replace(r_range, Chr(160), "")
    Next r_range

End Sub
