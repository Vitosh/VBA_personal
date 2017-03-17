Option Explicit

Sub EmbeddedLoops()
    
    Static size         As Long
    Static c            As Variant
    Static arr          As Variant
    Static n            As Long
    
    size = 4
    c = Array(1, 2, 3, 4, 5, 6)
    n = UBound(c) + 1
    ReDim arr(size - 1)
    
    Call embedded_loops(0, size, c, n, arr)
    
End Sub

Function embedded_loops(index, k, c, n, arr)
    
    Dim i                   As Variant
    
    If index >= k Then
        Call print_array_one_line(arr)
    Else
        For Each i In c
            arr(index) = i
            Call embedded_loops(index + 1, k, c, n, arr)
        Next i
    End If

End Function

Public Sub print_array_one_line(my_array As Variant)

    Dim counter     As Integer
    Dim s_array     As String
    
    For counter = LBound(my_array) To UBound(my_array)
        
        s_array = s_array & my_array(counter)
    
    Next counter
    Debug.Print s_array
    
    End Sub
