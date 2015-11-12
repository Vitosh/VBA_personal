Public Sub Increment(ByRef value_to_increment, Optional l_plus As Long = 1)
    
    value_to_increment = value_to_increment + l_plus

End Sub

Public Sub Decrement(value_to_decrement, Optional l_minus As Long = 1)
    
    value_to_decrement = value_to_decrement - l_minus

End Sub

Public Function l_locate_value_row(target As String, target_sheet As Worksheet, Optional l_col As Long = 1, Optional l_values_found As Long = 1) As Long

    For Each my_cell In target_sheet.Range(target_sheet.Cells(l_col, 1), target_sheet.Cells(Rows.Count, l_col))
        
        If target = my_cell Then
            If l_values_found = 1 Then
                l_locate_value_row = my_cell.Row
                Exit Function
            Else
                Call Decrement(l_values_found)
            End If
        End If
    Next my_cell
    
    l_locate_value_row = -1

End Function
