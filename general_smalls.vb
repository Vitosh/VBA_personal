Option Explicit

Sub FixRangeError() 
'Fix bezug fehler

    Dim r_range         As Range
    Dim str_text        As String
    Dim l_counter       As Long
    Dim str_result      As String
    
    Dim arr_result      As Variant
    Dim arr_range       As Variant
    
    ReDim arr_result(0)
    Set r_range = Selection
    str_text = Replace(r_range.Formula, "=", "")
    
    arr_range = Split(str_text, "+")
    
    For l_counter = LBound(arr_range) To UBound(arr_range)
        If Not InStr(arr_range(l_counter), "#") > 0 Then
            ReDim Preserve arr_result(UBound(arr_result) + 1)
            arr_result(UBound(arr_result)) = arr_range(l_counter)
        End If
    Next l_counter
    
    For l_counter = LBound(arr_result) + 1 To UBound(arr_result)
        str_result = str_result & "+" & arr_result(l_counter)
    Next l_counter
    
    Debug.Print str_result
    
End Sub
