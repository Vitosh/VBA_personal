Public Function l_locate_value_row(target As String, ByRef target_sheet As Worksheet, _
                                   Optional l_col As Long = 2, _
                                   Optional l_more_values_found As Long = 1, _
                                   Optional b_look_for_part = False) As Long

    Dim l_values_found  As Long
    Dim r_local_range   As Range
    Dim my_cell         As Range
    
    l_values_found = l_more_values_found

    Set r_local_range = Nothing
    wb_source.Activate
    Set r_local_range = target_sheet.Range(target_sheet.Cells(1, l_col), target_sheet.Cells(Rows.Count, l_col))

    For Each my_cell In r_local_range

        'The b_look_for_part is for the vertriebscase
        If b_look_for_part Then
            If target = Left(my_cell, Len(target)) Then
                If l_values_found = 1 Then
                    l_locate_value_row = my_cell.Row
                    Exit Function
                Else
                    Call Decrement(l_values_found)
                End If
            End If
        Else
            If target = Trim(my_cell) Then
                If l_values_found = 1 Then
                    l_locate_value_row = my_cell.Row
                    Exit Function
                Else
                    Call Decrement(l_values_found)
                End If
            End If
        End If
    Next my_cell

    l_locate_value_row = -1

End Function
