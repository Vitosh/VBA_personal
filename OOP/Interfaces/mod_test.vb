Option Explicit

Sub Test()
    
    Dim arr_collection(1 To 4)          As IGeneral
    Dim l_counter                       As Long
    Dim s_result                        As String
    
    Set arr_collection(1) = New cls_carport
    Set arr_collection(2) = New cls_tg
    Set arr_collection(3) = New cls_carport
    Set arr_collection(4) = New cls_tg
    
    For l_counter = LBound(arr_collection) To UBound(arr_collection)
        Call arr_collection(l_counter).Info
        Debug.Print arr_collection(l_counter).CalculatePrice(l_counter * 100)
    Next l_counter
    
End Sub
