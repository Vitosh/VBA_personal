'[hon_br_kosten].rows.count-1+[hon_br_kosten].row
'last row of named range

Public Function get_last_row_of_named_range(my_range As Range) As Long
    
    get_last_row_of_named_range = my_range.Rows.Count - 1 + my_range.Row

End Function
