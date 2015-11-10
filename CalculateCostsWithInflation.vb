'             If we use the optional argument, -> calculate_total_month_value_with_inflation(100,1.06,37,2),
'             this would return us the money for a month in the second period. -> 106 (100 + 1.06 inflation rate per year)



Public Function calculate_total_month_value_with_inflation(ByVal dbl_per_month As Double, ByVal dbl_inflation As Double, ByVal int_total_length, Optional ByVal int_period As Long = 0) As Double

    Dim months_left             As Long
    Dim years                   As Long
    Dim i_counter               As Long
    
    Dim dbl_result              As Double

    Dim previous_period         As Double
    
    
   On Error GoTo calculate_total_month_value_with_inflation_Error
   

    years = int_total_length \ MONTHS_IN_YEAR
    months_left = int_total_length - MONTHS_IN_YEAR * years
    
    For i_counter = 0 To years - 1
    
        If i_counter > 0 Then
            previous_period = dbl_result
        End If
        
        dbl_result = dbl_result + dbl_per_month * MONTHS_IN_YEAR * dbl_inflation ^ i_counter
        
        If int_period = i_counter + 1 Then
            calculate_total_month_value_with_inflation = (dbl_result - previous_period) / MONTHS_IN_YEAR
            Exit Function
        End If
        
    Next i_counter
    
    previous_period = dbl_result
    'adding values for months_left
    dbl_result = dbl_result + dbl_per_month * months_left * dbl_inflation ^ i_counter
    
    'checking if we need the values for the not filled months:
    
    If int_period > 0 Then
        If months_left = 0 Then
            calculate_total_month_value_with_inflation = dbl_per_month * dbl_inflation ^ (i_counter - 1)
            Exit Function
        Else
            calculate_total_month_value_with_inflation = (dbl_result - previous_period) / months_left
            Exit Function
        End If
    End If
    
    calculate_total_month_value_with_inflation = dbl_result

   On Error GoTo 0
   Exit Function

calculate_total_month_value_with_inflation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calculate_total_month_value_with_inflation of Modul mod_GeneralFunctions"
    
End Function
