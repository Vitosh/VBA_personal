'Select and run, values are printed

Public Sub MakeValues()

    Dim my_cell         As Range
    Dim str             As String
    Dim l_counter       As Long

    For Each my_cell In Selection
        Call Increment(l_counter)
        str = "my_arr(" & l_counter & ")= "

        If Len(my_cell) > 0 Then
            str = str & change_commas(my_cell.value)
        Else
            str = str & 0
        End If

        Debug.Print str

    Next my_cell

End Sub
