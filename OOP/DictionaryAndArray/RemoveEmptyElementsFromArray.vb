Public Function RemoveEmptyElementsFromArray(myArray As Variant) As Variant
    
    Dim i As Long, j As Long
    ReDim newArray(LBound(myArray) To UBound(myArray))
    
    For i = LBound(myArray) To UBound(myArray)
        If Trim(myArray(i)) <> "" Then
            j = j + 1
            newArray(j) = myArray(i)
        End If
    Next i
    
    ReDim Preserve newArray(LBound(myArray) To j - 1)
    RemoveEmptyElementsFromArray = newArray
    
End Function
