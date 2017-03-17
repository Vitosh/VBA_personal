Public Sub BorderMe(my_range)

    Dim l_counter   As Long

    For l_counter = 7 To 10 '7 to 10 are the magic numbers for xlEdgeLeft etc
        With my_range.Borders(l_counter)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next l_counter

End Sub
