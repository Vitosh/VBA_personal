Option Explicit

Public Function return_line(my_range As Range, percentage As Double) As Integer
    
    Dim my_cell     As Range
    Dim my_result   As Double
    
    For Each my_cell In my_range
        my_result = my_result + my_cell.Value
        If my_result >= (Application.WorksheetFunction.Sum(my_range) * percentage) Then
            return_line = my_cell.Row - my_range.Row + 1
            Exit For
            
        End If
    Next my_cell
    
End Function
