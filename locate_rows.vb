Public Function locate_last_col_in_row(target_sheet As Worksheet, Optional l_row As Long = 1) As Long
    
    locate_last_col_in_row = target_sheet.Cells(l_row, Columns.Count).End(xlToLeft).Column
    
End Function
