Option Explicit

Public Const NO_PREVIOUS = -1

Sub Main()

    Dim arr_seq         As Variant
    Dim arr_len         As Variant
    Dim arr_pre         As Variant
    
    Dim lng_best        As Long
    
    arr_seq = Array(1, 2, -6, -5, -3, 23, 123, 3, 2, -23, -5, 54, 100, 200, 300, 1111, 23412, 3, 4, 5, 6, 7, 8, 9, 19, 65, 2)
    ReDim arr_len(UBound(arr_seq))
    ReDim arr_pre(UBound(arr_seq))
    
    lng_best = CalculateLongestIncreasingSubsequence(arr_seq, _
                                                    arr_len, _
                                                    arr_pre)
    Call print_array(arr_seq)
    Call print_array(arr_len)
    Call print_array(arr_pre)
    
    Call PrintLongestIncreasingSubsequance(arr_seq, arr_pre, lng_best)
    
End Sub

Public Sub PrintLongestIncreasingSubsequance(ByRef arr_seq As Variant, _
                                            ByRef arr_pre As Variant, _
                                            lng_best As Long)
                                           
    Dim arr_result  As Variant
    Dim l_counter   As Long: l_counter = 0
    
    ReDim arr_result(1)
    
    While (lng_best <> NO_PREVIOUS)
        
        
        ReDim Preserve arr_result(l_counter)
        l_counter = l_counter + 1
        arr_result(l_counter - 1) = arr_seq(lng_best)
        lng_best = arr_pre(lng_best)
    
    Wend
    
    Debug.Print Join(reverse_array(arr_result), " ")
    
End Sub


Public Function CalculateLongestIncreasingSubsequence(ByRef arr_seq As Variant, _
                                                    ByRef arr_len As Variant, _
                                                    ByRef arr_pre As Variant) As Long

    Dim lng_best_len    As Long: lng_best_len = 0
    Dim lng_best_ind    As Long: lng_best_ind = 0
    Dim x               As Long
    Dim i               As Long
    
    For x = LBound(arr_seq) To (UBound(arr_seq)) Step 1
        arr_len(x) = 1
        arr_pre(x) = NO_PREVIOUS
        
        For i = 0 To x Step 1
            If (arr_seq(i) < arr_seq(x)) And (arr_len(i) + 1 > arr_len(x)) Then
                
                arr_len(x) = arr_len(i) + 1
                arr_pre(x) = i
                
                If arr_len(x) > lng_best_len Then
                    lng_best_len = arr_len(x)
                    lng_best_ind = x
                End If
            End If
            
        Next i
    Next x
        
    CalculateLongestIncreasingSubsequence = lng_best_ind
    
End Function

Public Sub print_array(ByRef my_array As Variant)
    Dim counter As Long
    
    For counter = LBound(my_array) To UBound(my_array)
        Debug.Print counter & " --> " & my_array(counter)
    Next counter
    Debug.Print "------------------------------"
End Sub

Public Function reverse_array(ByVal my_array As Variant) As Variant

    Dim counter     As Long
    Dim counter_2   As Long
    Dim arr_new     As Variant
    
    ReDim arr_new(UBound(my_array) + 1)
    
    For counter = LBound(arr_new) To UBound(arr_new) - 1
        counter_2 = UBound(arr_new) - counter - 1
        arr_new(counter) = my_array(counter_2)
    Next counter

    reverse_array = arr_new

End Function
