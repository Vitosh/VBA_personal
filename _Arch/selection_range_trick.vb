Public Sub SelectAndChange()
        
    Dim current_cells_range         As Range
    Dim my_array                    As Variant
    
    Dim l_step_between_BA           As Long
    Dim l_counter                   As Long
    Dim l_counter_2                 As Long
    Dim l_counter_3                 As Long
    Dim col                         As Long
    Dim row                         As Long
    
    l_step_between_BA = 17
    col = Selection.Column
    row = Selection.row
    'Beware what you select, for it would stay selected! :)
    
    Set current_cells_range = Selection
    
    For l_counter = 0 To 9
        Set current_cells_range = Union(current_cells_range, ActiveSheet.Cells(row + l_step_between_BA * l_counter, col))
        
'        Set current_cells_range = Union(current_cells_range, ActiveSheet.Cells(row + 1 + l_step_between_BA * l_counter, col))
'
'        Set current_cells_range = Union(current_cells_range, ActiveSheet.Cells(row + 2 + l_step_between_BA * l_counter, col))
'
'        Set current_cells_range = Union(current_cells_range, ActiveSheet.Cells(row + 3 + l_step_between_BA * l_counter, col))
'
'        Set current_cells_range = Union(current_cells_range, ActiveSheet.Cells(row + 4 + l_step_between_BA * l_counter, col))
        
    Next l_counter
    
    current_cells_range.Select
    
End Sub
