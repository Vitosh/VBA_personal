Public Function sum_array(my_array As Variant, Optional last_values_not_to_calculate As Long = 0) As Double
    'For unknown reasons, WorksheetFunction.sum(my_array) does not work always,
    'when we sum currency, long and double...
    
    Dim l_counter           As Long
    
    For l_counter = LBound(my_array) To UBound(my_array) - last_values_not_to_calculate
        sum_array = sum_array + my_array(l_counter)
    Next l_counter
    
End Function
