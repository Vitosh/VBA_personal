Option Explicit

Sub GreedyAlgorithm()
    
    Dim rowsCount           As Long
    Dim colCount            As Long
    Dim l_row_counter       As Long
    Dim l_col_counter       As Long
    Dim l_min_value         As Long
    Dim max_prev_cell       As Long
    
    Dim arr_sum             As Variant
    Dim arr_reverse         As Variant

    Dim rng                 As Range
    Dim rng2                As Range
    
    Calculate
    Application.Calculation = xlCalculationManual
    
    Set rng = [matrix]
    Set rng2 = [matrix2]
    
    rowsCount = [matrix].Rows.Count
    colCount = [matrix].Columns.Count
    rng2.Clear
    
    l_min_value = Application.WorksheetFunction.Min([matrix]) - 1
    ReDim arr_sum(rowsCount, colCount)
    ReDim arr_reverse(rowsCount, colCount)
    For l_row_counter = 1 To rowsCount
        For l_col_counter = 1 To colCount
                
            max_prev_cell = l_min_value
            
            If l_row_counter > 1 Then
                If arr_sum(l_row_counter - 1, l_col_counter) > max_prev_cell Then
                    max_prev_cell = arr_sum(l_row_counter - 1, l_col_counter)
                End If
            End If
            
            If l_col_counter > 1 Then
                If arr_sum(l_row_counter, l_col_counter - 1) > max_prev_cell Then
                    max_prev_cell = arr_sum(l_row_counter, l_col_counter - 1)
                End If
            End If
        
            arr_sum(l_row_counter, l_col_counter) = rng.Item(l_row_counter, l_col_counter)
            rng2.Item(l_row_counter, l_col_counter) = rng.Item(l_row_counter, l_col_counter)
            
            If max_prev_cell <> l_min_value Then
                arr_sum(l_row_counter, l_col_counter) = arr_sum(l_row_counter, l_col_counter) + max_prev_cell
                rng2.Item(l_row_counter, l_col_counter) = arr_sum(l_row_counter, l_col_counter)
            End If
            
        Next l_col_counter
    Next l_row_counter
    
    l_col_counter = l_col_counter - 1
    l_row_counter = l_row_counter - 1
    
    While (l_row_counter > 0) And (l_col_counter > 0)
        arr_reverse(l_row_counter, l_col_counter) = True
        If arr_sum(l_row_counter - 1, l_col_counter) > arr_sum(l_row_counter, l_col_counter - 1) Then
            l_row_counter = l_row_counter - 1
        Else
            l_col_counter = l_col_counter - 1
        End If

    Wend
    
    For l_row_counter = 1 To rowsCount
        For l_col_counter = 1 To colCount
            If arr_reverse(l_row_counter, l_col_counter) Then
                rng2.Item(l_row_counter, l_col_counter).Font.Color = vbRed
            End If
        Next l_col_counter
    Next l_row_counter
    
    rng.Columns.EntireColumn.AutoFit
    rng2.Columns.EntireColumn.AutoFit
    
    'Application.Calculation = xlAutomatic

End Sub
